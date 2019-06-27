Attribute VB_Name = "Main"
Option Explicit

' CONSTS
Public Const ERROR_IS_NOTHING As Long = vbObjectError + 514
Public Const ERROR_DUPLICATE As Long = vbObjectError + 515

' ENUMS
Enum CollType
  id = 1
  StartDate
  writePerson
  Description
  Priority
  responsiblePerson
  State
  EndDate
  Note
  Attachment
End Enum

Enum ChangeType
    Created
    Changed
End Enum

Enum ReturnType
    ValArray
    RefRange
End Enum

' PRIVATE variables
Private dictTasksOnOpen As Scripting.Dictionary
Private dictTasksOnClose As Scripting.Dictionary
Private dictRngTasksOnClose As Scripting.Dictionary

Private wsConfig As Worksheet, wsToDo As Worksheet


Sub onOpen()
    'On Error GoTo EH ' EH Topmost
    On Error GoTo 0
       
    Set wsConfig = ThisWorkbook.Worksheets("config")
    Set wsToDo = ThisWorkbook.Worksheets("ToDo")
    
    Set dictTasksOnOpen = Tasks2Dict(wsToDo, ReturnType.ValArray)
    
Done:
    Exit Sub
EH:
    DisplayError Err.Source, Err.Description, "Module1.Topmost", Erl
End Sub

Sub onClose()
    'On Error GoTo EH ' EH Topmost
    On Error GoTo 0
    
     If wsToDo.Range("A" & wsToDo.Rows.Count).End(xlUp).row = 4 Then
        Exit Sub
    End If
    
    ' EH
    If wsConfig Is Nothing Then
        Err.Raise ERROR_IS_NOTHING, "onClose", "wsConfig is Nothing"
    End If

    Dim dictChanges As New Scripting.Dictionary
    Dim dictPersons As New Scripting.Dictionary
    Dim dictSupervisors As New Scripting.Dictionary
    Dim dictEmails As New Scripting.Dictionary
    
    Set dictSupervisors = GetSupervisors(wsConfig, "Users_table", "yes")
    Set dictEmails = GetEmailList(wsConfig, "Users_table")
    
    Set dictTasksOnClose = Tasks2Dict(wsToDo, ReturnType.ValArray)
    Set dictRngTasksOnClose = Tasks2Dict(wsToDo, ReturnType.RefRange)
    Set dictChanges = CompareOpenClose(dictTasksOnOpen, dictTasksOnClose, dictSupervisors)
    Set dictPersons = getPersons(dictChanges)
    Call AddEmail2Person(dictEmails, dictPersons)
    
    If dictChanges.Count > 0 Then
        Call SendMail2Persons(dictPersons, dictChanges, wsToDo)
    End If
    
Done:
    Exit Sub
EH:
    DisplayError Err.Source, Err.Description, "Module1.Topmost", Erl
End Sub

Function Tasks2Dict(wsToDo As Worksheet, ReturnType As ReturnType) As Scripting.Dictionary
    Dim dictTasks As New Scripting.Dictionary
    Dim i As Integer, row As Range, rowId As Integer
    Dim iLastRow As Integer, rngToDo As Range
    
    iLastRow = wsToDo.Range("A" & wsToDo.Rows.Count).End(xlUp).row
    If iLastRow < 5 Then ' pokud neexistuje žádný úkol, mohlo by dojít k situaci A5:J4, což vyhodí chybu, proto
        iLastRow = 5
    End If
    Set rngToDo = wsToDo.Range("A5:J" & iLastRow)
    
    For Each row In rngToDo.Rows
        rowId = row.Cells(1, 1)
        
        'EH
        If dictTasks.Exists(rowId) Then
            Err.Raise ERROR_DUPLICATE, "Tasks2Dict", "Task number dupplicate"
        End If
        
        If ReturnType = RefRange Then
            dictTasks.Add rowId, row
        Else
            dictTasks.Add rowId, row.value
        End If
    Next row
    
    Set Tasks2Dict = dictTasks
End Function

Function CompareOpenClose(dict1 As Scripting.Dictionary, _
                          dict2 As Scripting.Dictionary, _
                          dictSupervisors As Scripting.Dictionary) As Scripting.Dictionary
            
    Dim dictChanges As New Scripting.Dictionary
    Dim dictPersons As Scripting.Dictionary
    Dim originalWritePerson, originalResponsiblePerson, newWritePerson, newResponsiblePerson
    Dim key As Variant
    Dim sRowOld As String, sRowNew As String
    
    For Each key In dict2.Keys
        Set dictPersons = CreateObject("Scripting.Dictionary")
    
        newWritePerson = LCase(dict2(key)(1, CollType.writePerson))
        newResponsiblePerson = LCase(dict2(key)(1, CollType.responsiblePerson))
        
        If dict1.Exists(key) Then
        
            originalWritePerson = LCase(dict1(key)(1, CollType.writePerson))
            originalResponsiblePerson = LCase(dict1(key)(1, CollType.responsiblePerson))
        
            sRowOld = ArrayToDelimitedString(dict1(key), ",")
            sRowNew = ArrayToDelimitedString(dict2(key), ",")
            
            If sRowOld <> sRowNew Then
            
            dictPersons.Add originalWritePerson, False
            Call dictAddIfNotContain(dictPersons, originalResponsiblePerson, False)
            Call dictAddIfNotContain(dictPersons, newWritePerson, False)
            Call dictAddIfNotContain(dictPersons, newResponsiblePerson, False)
            
            dictChanges.Add key, GetChange(key, _
                                           dictRngTasksOnClose(key), _
                                           ChangeType.Changed, _
                                           MergeDicts(dictPersons, dictSupervisors))
            
            End If
        Else
            dictPersons.Add newWritePerson, False
            Call dictAddIfNotContain(dictPersons, newResponsiblePerson, False)
            
            dictChanges.Add key, GetChange(key, _
                                           dictRngTasksOnClose(key), _
                                           ChangeType.Created, _
                                           MergeDicts(dictPersons, dictSupervisors))
        End If
        
        dictPersons.RemoveAll
        Set dictPersons = Nothing
    Next key
    Set CompareOpenClose = dictChanges
End Function

Function GetChange(ByVal id As String, _
                   ByRef rng As Range, _
                   ByVal ChangeType As ChangeType, _
                   ByRef dictPersons As Scripting.Dictionary) As Scripting.Dictionary
                   
    Dim dictChange As New Scripting.Dictionary
    Dim key As Variant
    
    dictChange.Add "id", id
    dictChange.Add "range", rng
    dictChange.Add "changeType", ChangeType
    dictChange.Add "persons", dictPersons
    
    Set GetChange = dictChange
End Function

Function getPersons(dictChanges As Scripting.Dictionary)
    Dim dictPersons As New Scripting.Dictionary
    
    Dim key, key2 As Variant
    Dim personId As String
    

    For Each key In dictChanges.Keys
        For Each key2 In dictChanges(key)("persons").Keys
            personId = key2
            If Not dictPersons.Exists(personId) Then
                dictPersons.Add personId, New Scripting.Dictionary
            End If
        Next key2
    Next key
    
    Set getPersons = dictPersons
End Function

Function GetSupervisors(ws As Worksheet, strRangeName As String, strYes As String)
    Dim dictSupervisors As New Scripting.Dictionary
    Dim rngTabNames, row As Range
    
    Set rngTabNames = ws.Range(strRangeName)
    
    For Each row In rngTabNames.Rows
        If row.Cells(1, 3) = strYes Then
            dictSupervisors.Add LCase(row.Cells(1, 1).value), True
        End If
    Next row
    Set GetSupervisors = dictSupervisors
End Function

Function GetEmailList(ws As Worksheet, strRangeName As String)
    Dim dictEmailList As New Scripting.Dictionary
    Dim rngTabNames, row As Range
    Set rngTabNames = ws.Range(strRangeName)
    
    For Each row In rngTabNames.Rows
        dictEmailList.Add row.Cells(1, 1).value, row.Cells(1, 2).value
    Next row
    Set GetEmailList = dictEmailList
End Function

Function AddEmail2Person(dictEmails, dictPersons)
    Dim personKey As Variant
    
    For Each personKey In dictPersons.Keys
        dictPersons(personKey).Add "email", dictEmails(personKey)
    Next personKey
End Function

Public Function ArrayToDelimitedString(variantArray As Variant, separator As String) As String
   Dim delimitedString As String, index As Integer

   For index = 1 To UBound(variantArray, 2)
      delimitedString = delimitedString & CStr(variantArray(1, index)) & separator
   Next
   ArrayToDelimitedString = Left(delimitedString, Len(delimitedString) - 1)
End Function

Function MergeDicts(Dct1, Dct2)
    'Merge 2 dictionaries. The second dictionary will override the first if they have the same key

    Dim Res, key

    Set Res = CreateObject("Scripting.Dictionary")

    For Each key In Dct1.Keys()
        Res.Item(key) = Dct1(key)
    Next

    For Each key In Dct2.Keys()
        Res.Item(key) = Dct2(key)
    Next

    Set MergeDicts = Res
End Function

Function CloneDictionary(Dict)
 ' https://stackoverflow.com/a/3022349
  Dim newDict As New Scripting.Dictionary
  Dim key

  For Each key In Dict.Keys
    newDict.Add key, Dict(key)
  Next
  newDict.CompareMode = Dict.CompareMode

  Set CloneDictionary = newDict
End Function

Function dictAddIfNotContain(Dict As Scripting.Dictionary, key, value)
    If Not Dict.Exists(key) Then
        Dict.Add key, value
    End If
End Function


Sub CalculateRunTime(strFuncName)
    Dim StartTime As Double
    Dim SecondsElapsed As Double

    StartTime = Timer

    eval (strFuncName)

    SecondsElapsed = Round(Timer - StartTime, 2)

    Debug.Print SecondsElapsed & " s"
End Sub

Function eval(fcName As String) As String
        Application.Run (fcName)
End Function

