Attribute VB_Name = "DiffTwoLists"
Option Explicit

' GLOBAL VARIABLES
Dim gShOld As Worksheet
Dim gShNew As Worksheet
Dim gCollectionForbidenBc As Collection

' CONSTANTS

' ENUMS
Enum changeType
    ADDED = 1
    REMOVED
    PLUS
    MINUS
End Enum

Function GetEnumName(i) As String
    Select Case i
        Case 1:      GetEnumName = "ADDED"
        Case 2:      GetEnumName = "REMOVED"
        Case 3:      GetEnumName = "PLUS"
        Case 4:      GetEnumName = "MINUS"
    End Select
End Function

Function HasForbiddenBc(rng As Range) As Boolean
    Dim item
    Dim rngBcColor
    rngBcColor = rng.Interior.Color
    'Debug.Print rngBcColor
    For Each item In gCollectionForbidenBc
            If item = rngBcColor Then
                HasForbiddenBc = True
                Exit Function
            End If
    Next item
    HasForbiddenBc = False
End Function

Sub GetStaticInfo()
    Set gShOld = Worksheets("Old")
    Set gShNew = Worksheets("New")
    Set gCollectionForbidenBc = New Collection
    gCollectionForbidenBc.Add gShNew.Range("K8").Interior.Color
    gCollectionForbidenBc.Add gShNew.Range("K9").Interior.Color
    gCollectionForbidenBc.Add gShNew.Range("K10").Interior.Color
End Sub

Sub btnCompare_handler()
    Call Compare
End Sub

Sub Compare()
    Dim key, i
    Dim dictOld As Object
    Dim dictNew As Object
    Dim dictDiff As Object
    Dim shDiff As Worksheet
    
    Call GetStaticInfo

    Set dictOld = Data2dict(gShOld)
    Set dictNew = Data2dict(gShNew)

    Set dictDiff = GetDictDiff(dictOld, dictNew)

    Call DeleteDiffSh
    Set shDiff = CreateDiffSh
    
    shDiff.Range("A1").value = "Desc"
    shDiff.Range("B1").value = "Order number"
    shDiff.Range("C1").value = "Manufacturer"
    shDiff.Range("D1").value = "Change"
    shDiff.Range("E1").value = "Unit"
    shDiff.Range("F1").value = "Change type"
    shDiff.Range("A1:F1").Interior.Color = RGB(217, 217, 217)
    
    i = 2
    For Each key In dictDiff.Keys
        shDiff.Range("A" + CStr(i)).value = dictDiff(key)("Description")
        shDiff.Range("B" + CStr(i)).value = key
        shDiff.Range("C" + CStr(i)).value = dictDiff(key)("Producer")
        shDiff.Range("D" + CStr(i)).value = dictDiff(key)("ChangeCount")
        shDiff.Range("E" + CStr(i)).value = dictDiff(key)("Unit")
        shDiff.Range("F" + CStr(i)).value = GetEnumName(dictDiff(key)("ChangeType"))
        i = i + 1
    Next key
    With shDiff.Columns("A:F")
        .HorizontalAlignment = xlLeft
        .AutoFit
    End With
End Sub

Function Data2dict(sh As Worksheet)
    Dim i As Integer

    Dim Dict As Object
    Set Dict = CreateObject("Scripting.Dictionary")

    Dim sItemIdColl As String
    Dim sItemCountColl As String
    Dim sItemDescColl As String
    Dim sItemProducerColl As String
    Dim sItemUnitColl As String
    
    sItemIdColl = "C"
    sItemCountColl = "F"
    sItemDescColl = "B"
    sItemProducerColl = "D"
    sItemUnitColl = "G"
    
    Dim iFirstRow As Integer
    Dim iLastRow As Integer
    
    iFirstRow = 3
    iLastRow = sh.Cells(sh.Rows.Count, sItemIdColl).End(xlUp).row
    
    Dim sItemId As String
    Dim iItemCount As Integer
    Dim sItemDesc As String
    Dim sItemProducer As String
    Dim sItemUnit As String
    
    For i = iFirstRow To iLastRow
        ' Filter
        If HasForbiddenBc(sh.Range(sItemIdColl + CStr(i))) Then GoTo NextIteration
        If sh.Range(sItemIdColl + CStr(i)) = "" Then GoTo NextIteration
        ' Get data
        sItemId = sh.Range(sItemIdColl + CStr(i)).value
        iItemCount = sh.Range(sItemCountColl + CStr(i)).value
        sItemDesc = sh.Range(sItemDescColl + CStr(i)).value
        sItemProducer = sh.Range(sItemProducerColl + CStr(i)).value
        sItemUnit = sh.Range(sItemUnitColl + CStr(i)).value
        
        If Dict.Exists(sItemId) Then
            Dict(sItemId)("Count") = Dict(sItemId)("Count") + iItemCount
        Else
            Dim itemDict As Object
            Set itemDict = CreateObject("Scripting.Dictionary")
            
            itemDict.Add "Count", iItemCount
            itemDict.Add "Description", sItemDesc
            itemDict.Add "Producer", sItemProducer
            itemDict.Add "Unit", sItemUnit
            Dict.Add sItemId, itemDict
        End If
NextIteration:
    Next i
    Set Data2dict = Dict
End Function

Function GetDictDiff(dictOld As Scripting.Dictionary, dictNew As Scripting.Dictionary)
    Dim key
    
    Dim dictDiff As Object
    Set dictDiff = CreateObject("Scripting.Dictionary")
    
    ' Foreach dictOld
    For Each key In dictOld.Keys
        If dictNew.Exists(key) Then
            If dictNew(key)("Count") > dictOld(key)("Count") Then
                'PLUS
                dictDiff.Add key, GetChange(dictOld(key), CInt(dictNew(key)("Count") - dictOld(key)("Count")), changeType.PLUS)
            ElseIf dictNew(key)("Count") < dictOld(key)("Count") Then
                'MINUS
                dictDiff.Add key, GetChange(dictOld(key), CInt(dictNew(key)("Count") - dictOld(key)("Count")), changeType.MINUS)
            End If
        Else
            ' REMOVED
            dictDiff.Add key, GetChange(dictOld(key), 0 - CInt(dictOld(key)("Count")), changeType.REMOVED)
        End If
    Next key
    
    'Foreach dictNew
    For Each key In dictNew.Keys
        If Not dictOld.Exists(key) Then
            'ADDED
            dictDiff.Add key, GetChange(dictNew(key), CInt(dictNew(key)("Count")), changeType.ADDED)
        End If
    Next key
    
'    For Each key In dictDiff.Keys
'        Debug.Print key; dictDiff(key)("ChangeCount"); GetEnumName(dictDiff(key)("Type"))
'    Next key
    Set GetDictDiff = dictDiff
End Function

Function GetChange(itemDict As Scripting.Dictionary, changeCount As Integer, aChangeType As changeType)
    Dim changeDict  As Scripting.Dictionary
    Set changeDict = New Dictionary
    changeDict.Add "ChangeCount", changeCount
    changeDict.Add "ChangeType", aChangeType
    changeDict.Add "Description", itemDict("Description")
    changeDict.Add "Producer", itemDict("Producer")
    changeDict.Add "Unit", itemDict("Unit")
    Set GetChange = changeDict
End Function

Function CreateDiffSh()
    Dim shDiff As Worksheet
    Set shDiff = ThisWorkbook.Sheets.Add(After:= _
             ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    shDiff.Name = "Diff"
    Set CreateDiffSh = shDiff
End Function

Sub DeleteDiffSh()
    On Error Resume Next
    Application.DisplayAlerts = False
    If WorksheetExists("Diff") Then
        Worksheets("Diff").Delete
    End If
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

' HELPERS
 Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

     If wb Is Nothing Then Set wb = ThisWorkbook
     On Error Resume Next
     Set sht = wb.Sheets(shtName)
     On Error GoTo 0
     WorksheetExists = Not sht Is Nothing
 End Function
