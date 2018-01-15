Attribute VB_Name = "STK1_Main"
Sub Stocktake(control As IRibbonControl)

Dim wbStock As Workbook
Dim wbImport As Workbook

Dim StockList As Variant
Dim ImportList As Variant
Dim SplitList As Variant

Dim dicStock As Scripting.Dictionary
Dim dicSplitKits As Scripting.Dictionary
Dim dicSplitKitsUnique As Scripting.Dictionary
Dim strKey As Variant

Dim rowStock As Integer
Dim LineNum As Integer

'Main
Application.ScreenUpdating = False

'Check that the active spreadsheet is a stocktake sheet else open a new document.
If CheckFormat_StockSheet = False Then
    'Open StockSheet
    MsgBox "hi"
    Exit Sub
End If

Set wbStock = ActiveWorkbook
rowStock = RowNum(1)    'Record the number of rows in Column A.

'Create Dictionary of all Stock Items
StockList = Range("A1:L" & RowNum(4))
Set dicStock = New Scripting.Dictionary: dicStock.CompareMode = vbTextCompare
For i = 2 To RowNum(4): dicStock(StockList(i, 1)) = i: Next i

'Create Dictionary of Known split components (from THIS WORKBOOOK)
SplitList = ThisWorkbook.Sheets("Stocktake_Exceptions").Range("A2:D" & ThisWorkbook.Sheets("Stocktake_Exceptions").Range("A" & 60000).End(xlUp).Row)
Set dicSplitKits = New Scripting.Dictionary: dicSplitKits.CompareMode = vbTextCompare

LineNum = 2
For i = 1 To UBound(SplitList, 1)
    If SplitList(i, 2) = wbStock.Sheets(1).Cells(2, 8).Value Then
        dicSplitKits(SplitList(i, 1)) = LineNum
        If CheckSheet("SplitKits") = False Then
            ActiveWorkbook.Sheets.Add(After:=Sheets(Sheets.Count)).Name = "SplitKits"
            Cells(1, 1).Value = "Scanned ID"
            Cells(1, 2).Value = "Location"
            Cells(1, 3).Value = "Converted ID"
            Cells(1, 4).Value = "Count Conversion"
            Cells(1, 5).Value = "Count"
            Cells(1, 6).Value = "Converted Count"
        End If
        Cells(LineNum, 1).Value = SplitList(i, 1)
        Cells(LineNum, 2).Value = SplitList(i, 2)
        Cells(LineNum, 3).Value = SplitList(i, 3)
        Cells(LineNum, 4).Value = SplitList(i, 4)
        LineNum = LineNum + 1
    End If
Next i
Sheets(1).Activate
Set dicSplitKitsUnique = New Scripting.Dictionary: dicSplitKitsUnique.CompareMode = vbTextCompare

'Open ImportList
   
    Set myfile = Application.FileDialog(msoFileDialogOpen)

    With myfile
        .Title = "Choose the file to Import data from"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            Exit Sub
        End If
        FileSelected = .SelectedItems(1)
    End With
    SampleFilePath = FileSelected
    

'Open File
Workbooks.Open Filename:=SampleFilePath
  
 Set wbImport = ActiveWorkbook
 If CheckFormat_StockSheet = False Then Exit Sub  'Checks that the active spreadsheet is an import sheet.
 ImportList = Range("B1:E" & RowNum(4))
 wbImport.Close
 
'Copy counts to Spreadsheet
rowStock = rowStock + 1
 
 For i = 2 To UBound(ImportList, 1)
        If dicStock.Exists(ImportList(i, 1)) Then
            'If item number found in list then add the count value (with comparision to prior figure checking)
            Cells(dicStock(ImportList(i, 1)), 12).Value = ImportList(i, 4)
            If ImportList(i, 4) > Cells(dicStock(ImportList(i, 1)), 11).Value Then
                'If current count is higher then previous count highlight value red.
                Cells(dicStock(ImportList(i, 1)), 12).Interior.Color = RGB(255, 0, 0)
            End If
        ElseIf dicSplitKits.Exists(ImportList(i, 1)) Then
            'If item number not found in the list but is found in the SplitKits list then add count to SplitKits list
            Sheets("SplitKits").Activate
            LineNum = dicSplitKits(ImportList(i, 1))            'Find line number
            Cells(LineNum, 5).Value = ImportList(i, 4)          'Add count
            Cells(LineNum, 6).Value = Cells(LineNum, 4) * Cells(LineNum, 5) 'Calculate count as a proportion of the entire parent kit (based on 4th column value).
            Sheets(1).Activate
        ElseIf ImportList(i, 4) <> 0 Then
            'If a completely Unqiue Item number not contained in the list (& not 0) then add a new record at the end of the list
            Cells(rowStock, 1).Value = ImportList(i, 1)        'Item Number
            Cells(rowStock, 12).Value = ImportList(i, 4)       'Current Quantity
            Cells(rowStock, 8).Value = Cells(2, 8).Value       'Region (Region defind as value from the location column the of the first record)
            rowStock = rowStock + 1
        End If
 Next i

'Combine Splitkit values and post result back to main list
Sheets("SplitKits").Activate
SplitList = Range("A1:F" & RowNum(4))
For i = 2 To RowNum(4)
    If dicSplitKitsUnique.Exists(Cells(i, 3).Value) Then
        dicSplitKitsUnique(Cells(i, 3).Value) = dicSplitKitsUnique(Cells(i, 3).Value) + Cells(i, 6).Value
    Else
        dicSplitKitsUnique(Cells(i, 3).Value) = Cells(i, 6).Value
    End If
Next i
Sheets(1).Activate
For Each strKey In dicSplitKitsUnique.Keys()
    Cells(dicStock(strKey), 12).Value = Round(dicSplitKitsUnique(strKey), 1)
Next


Application.ScreenUpdating = True

End Sub


Function CheckSheet(pName As String) As Boolean
'Updateby20140617
Dim IsExist As Boolean
IsExist = False
For i = 1 To Application.ActiveWorkbook.Sheets.Count
    If Application.ActiveWorkbook.Sheets(i).Name = pName Then
        IsExist = True
        Exit For
    End If
Next
CheckSheet = IsExist
End Function


