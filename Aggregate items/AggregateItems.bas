Attribute VB_Name = "AggregateItems"
Option Explicit

Sub GetInfo()

' sheet Data
Dim sh As Worksheet
Set sh = Worksheets("Data")

' helpers
Dim i As Integer
Dim key As Variant
Dim articleNum As String
Dim articleCount As Integer

' In
Dim inFirstCol As String
Dim inLastCol As String

Dim inArticleCol As String
Dim inCountCol As String
Dim inDescCol As String
Dim inUnitCol As String
Dim inProducerCol As String

' Out
Dim outArticleCol As String
Dim outCountCol As String
Dim outDescCol As String
Dim outUnitCol As String
Dim outProducerCol As String

'... of Data
Dim firstRow As Integer
Dim lastRow As Integer

Dim itemCount As Integer

'Indentical
Dim idItemsDict As Object
Dim itemIdCount As Integer

inFirstCol = "A"
inLastCol = "H"

inArticleCol = "C"
inCountCol = "F"
inDescCol = "B"
inUnitCol = "G"
inProducerCol = "D"

outArticleCol = "B"
outCountCol = "D"
outDescCol = "A"
outUnitCol = "E"
outProducerCol = "C"

firstRow = 3
'https://www.thespreadsheetguru.com/blog/2014/7/7/5-different-ways-to-find-the-last-row-or-last-column-using-vba
lastRow = sh.Cells(sh.Rows.Count, inArticleCol).End(xlUp).Row

'Without Headings
itemCount = lastRow - 2

'1. Aggregate items
Set idItemsDict = CreateObject("Scripting.Dictionary")
For i = firstRow To lastRow
    articleNum = sh.Range(inArticleCol + CStr(i)).Value
    articleCount = sh.Range(inCountCol + CStr(i)).Value
    If idItemsDict.Exists(articleNum) Then
        idItemsDict(articleNum) = idItemsDict(articleNum) + articleCount
    Else
        idItemsDict.Add articleNum, articleCount
    End If
Next i

'2. Get information
Dim idItemsArr() As String
i = 0
'0 desc 1 orden number 2 manufacturer 3 qty 4 unit
ReDim idItemsArr(idItemsDict.Count - 1, 4)
For Each key In idItemsDict.Keys
   idItemsArr(i, 0) = CStr(WorksheetFunction.Index(sh.Range(inDescCol + CStr(firstRow), inDescCol + CStr(lastRow)), CStr(WorksheetFunction.Match(key, sh.Range(inArticleCol + CStr(firstRow), inArticleCol + CStr(lastRow)), 0))))
   idItemsArr(i, 1) = key
   idItemsArr(i, 2) = CStr(WorksheetFunction.Index(sh.Range(inProducerCol + CStr(firstRow), inProducerCol + CStr(lastRow)), CStr(WorksheetFunction.Match(key, sh.Range(inArticleCol + CStr(firstRow), inArticleCol + CStr(lastRow)), 0))))
   idItemsArr(i, 3) = CStr(idItemsDict(key))
   idItemsArr(i, 4) = CStr(WorksheetFunction.Index(sh.Range(inUnitCol + CStr(firstRow), inUnitCol + CStr(lastRow)), CStr(WorksheetFunction.Match(key, sh.Range(inArticleCol + CStr(firstRow), inArticleCol + CStr(lastRow)), 0))))
   i = i + 1
Next key

'3. delete old data
sh.Range(inFirstCol + CStr(firstRow), inLastCol + CStr(lastRow)).ClearContents

'4. write aggregate data
i = 0
For i = 0 To UBound(idItemsArr)
    sh.Range(inDescCol + CStr(firstRow + i)) = idItemsArr(i, 0)
    sh.Range(inArticleCol + CStr(firstRow + i)) = idItemsArr(i, 1)
    sh.Range(inProducerCol + CStr(firstRow + i)) = idItemsArr(i, 2)
    sh.Range(inCountCol + CStr(firstRow + i)) = idItemsArr(i, 3)
    sh.Range(inUnitCol + CStr(firstRow + i)) = idItemsArr(i, 4)
Next i
End Sub
