Attribute VB_Name = "SendEmail"
Option Explicit

Sub SendMail2Persons(persons As Scripting.Dictionary, changes As Scripting.Dictionary, wsToDo As Worksheet)
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
'Don't forget to copy the function RangetoHTML in the module.
'Working in Excel 2000-2016
' Refaktorováno - výpis zmìn vloženo do funkce
    Dim OutApp As Object
    Dim OutMail As Object
    
    Dim StrBody, StrHeading, StrCreated, StrChanged, StrFootnote As String
    Dim projectName, projectNumber, filePath, fileName As String
    
    Dim emails As String
    Dim key As Variant
    
     For Each key In persons.Keys
        emails = emails & persons(key)("email") & ";"
    Next key
    emails = Left(emails, Len(emails) - 1)
    
    projectName = wsToDo.Range("Project_name")
    projectNumber = wsToDo.Range("Project_number")
    fileName = ActiveWorkbook.Name
    filePath = ActiveWorkbook.Path & "\" & fileName
     
    Application.ScreenUpdating = False
     
    StrHeading = "<span style=""font-size:14pt"">Changes in ToDO list for <b>" & projectNumber & " " & projectName & "</b></span>"
    StrCreated = GetHtmlFromChanges(changes, ChangeType.Created, "New tasks")
    StrChanged = GetHtmlFromChanges(changes, ChangeType.Changed, "Task changes")
    
    StrFootnote = "<p> Todo list path:<br /><a href=""" & filePath & """>" & filePath & "</a></p><br />"
    StrFootnote = StrFootnote & "<p>(this message is generated automatically| contact: <a href=""mailto:admin@domain.com"">admin@domain.com</a>)</p>"
     
    StrBody = "<BODY style=""font-size:11pt;font-family:Calibri""><div>" & _
              StrHeading & "<br /><br />" & _
              StrCreated & "<br />" & _
              StrChanged & "<br /><br />" & _
              StrFootnote & _
              "</div></BODY>"

    On Error Resume Next

    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With OutMail
        .To = emails
        .CC = ""
        .BCC = ""
        .Subject = projectNumber & " " & projectName & "_" & "ToDo list changes"
        .HTMLBody = StrBody
        '.Send
        .Display
    End With

    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub


Function RangetoHTML(rng As Range, rowHeight As Integer)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2016
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook, newWorkbook As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    
    
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).rowHeight = rowHeight
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        .Cells(1).PasteSpecial Paste:=xlPasteAll
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With
    
    ' Set pasted range
    Dim pastedRange As Range
    With TempWB.Sheets(1)
        Set pastedRange = .Range(.Cells(1, 1), .Cells(rng.Rows.Count, rng.Columns.Count))
    End With
    
    pastedRange.FormatConditions.Delete
    
    ' Copy background color from TodoList to temp range (influenced by conditional formating)
    Dim i, j As Integer
    Dim sourceCell, targetCell As Range
    For i = 1 To rng.Rows.Count
        For j = 1 To rng.Columns.Count
            Set sourceCell = rng.Cells(i, j)
            Set targetCell = TempWB.Worksheets(1).Cells(i, j)
            With targetCell
              .Font.FontStyle = sourceCell.DisplayFormat.Font.FontStyle
              .Interior.Color = sourceCell.DisplayFormat.Interior.Color
              .Font.Strikethrough = sourceCell.DisplayFormat.Font.Strikethrough
            End With
        Next j
    Next i

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         fileName:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function

Function GetHtmlFromChanges(dictChanges As Scripting.Dictionary, _
                            chType As ChangeType, _
                            headline As String)
                            
    Dim tempRange As Range, key As Variant
    Dim strHtml As String
    Dim isAny As Boolean
    isAny = False

    If dictChanges.Count > 0 Then
        strHtml = strHtml & "<u>" & headline & "</u>"
        For Each key In dictChanges.Keys
            If dictChanges(key)("changeType") = chType Then
                isAny = True
                Set tempRange = dictChanges(key)("range")
                strHtml = strHtml + RangetoHTML(tempRange, 30)
                Set tempRange = Nothing
            End If
        Next key
    End If
    
    If isAny Then
        GetHtmlFromChanges = strHtml
    Else
        GetHtmlFromChanges = ""
    End If
End Function

