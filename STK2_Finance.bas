Attribute VB_Name = "STK2_Finance"
Sub Create_StocktakeSheets(control As IRibbonControl)
    
    'Variables
    Dim i As Single, j As Single, k As Single
    Dim wbImport As Workbook
    Dim ws As Worksheet
    Dim RowNum As Single
    Dim ImportData As Variant
    Dim Nodes As Variant
    Dim ImportPath As String
    
    'Main Program
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'Check if current sheet contains Navision data or open new file
    If Not (Range("A1").Value = "PHYS. INVE" And Range("B1").Value = "DEFAULT" And Range("O1").Value <> "" And Range("P1").Value = "") Then
        If GetNavData = False Then Exit Sub 'Open Import file.  If incorrect file opened (i.e. Func=FALSE) then exit sub
    End If

    Set wbImport = ActiveWorkbook
    ImportPath = Application.ActiveWorkbook.Path
    
    'Copy Data to variant
    RowNum = Range("A" & 65536).End(xlUp).Row  'Find the number of rows,  The columns are already known.
    ImportData = Range("A1:O" & RowNum)
    
    'Find Unique Nodes
    Range("L1:L" & RowNum).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("Q1"), Unique:=True
    Nodes = Range("Q2:Q" & Range("Q" & 65536).End(xlUp).Row)
    Columns("Q:Q").Delete
    
    'Create Mastersheet
    Workbooks.Add
    Sheets(1).Name = "Stocktake"
 
    Range("A1").Value = "Type"
    Range("B1").Value = "Default"
    Range("C1").Value = "Line Number"
    Range("D1").Value = "Item Number"
    Range("E1").Value = "Description"
    Range("F1").Value = "Description 2"
    Range("G1").Value = "UOM"
    Range("H1").Value = "Vendor"
    Range("I1").Value = "Location"
    Range("J1").Value = "Section"
    Range("K1").Value = "Region"
    Range("L1").Value = "Category"
    Range("M1").Value = "Shelf/Bin"
    Range("N1").Value = "Previous Qty"
    Range("O1").Value = "Current Qty"
    Range("P1").Value = "Unit Cost"
    
    j = 2 'Let j be the row counter.  Hold which rwo to print the next record on. (Main list)
    For i = 1 To UBound(ImportData, 1)
            Cells(j, 1).Value = ImportData(i, 1)
            Cells(j, 2).Value = ImportData(i, 2)
            Cells(j, 3).Value = ImportData(i, 3)
            Cells(j, 4).Value = ImportData(i, 4)        'Item Number
            Cells(j, 5).Value = ImportData(i, 5)        'Description
            Cells(j, 6).Value = ImportData(i, 14)       'Description 2
            Cells(j, 7).Value = ImportData(i, 8)        'UOM
            Cells(j, 8).Value = ImportData(i, 9)        'Vendor
            Cells(j, 9).Value = ImportData(i, 6)        'Location
            Cells(j, 10).Value = ImportData(i, 11)      'Section
            Cells(j, 11).Value = ImportData(i, 12)      'Region
            Cells(j, 12).Value = ImportData(i, 13)      'Category
            Cells(j, 13).Value = ImportData(i, 7)       'Shelf/Bin
            Cells(j, 14).Value = ImportData(i, 10)      'Previous Qty
            Cells(j, 16).Value = ImportData(i, 15)      'Unit Cost
            If Left(ImportData(i, 6), 1) <> Left(ImportData(i, 12), 1) Or Mid(ImportData(i, 6), 2) <> ImportData(i, 11) Then
                Range(Cells(j, 9), Cells(j, 11)).Interior.Color = RGB(255, 0, 0)
            End If
            j = j + 1
    Next i
    
    'Formating
    Range("A1:P1").Font.Bold = True
    Columns("A:C").EntireColumn.Hidden = True
    Columns("D:D").ColumnWidth = 20
    Columns("D:D").NumberFormat = "0"
    Columns("E:E").ColumnWidth = 35
    Columns("F:F").ColumnWidth = 22
    Columns("G:G").ColumnWidth = 10
    Columns("H:H").ColumnWidth = 35
    Columns("I:J").ColumnWidth = 12
    Columns("N:N").EntireColumn.Hidden = True
    Columns("O:O").ColumnWidth = 10
    Columns("P:P").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"    'Accounting Format
    Columns("D:G").HorizontalAlignment = xlLeft
    
    'Sort by Description
    Columns("A:M").Sort Key1:=Range("B2"), Order1:=xlAscending, _
    Header:=xlYes, OrderCustom:=1, MatchCase:=False, _
    Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2:=xlSortNormal                        'Sort Data
    
    'Setup Print Preferences
    ActiveSheet.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = "Please use pen | Sheets are a GUIDE ONLY | 2 person to sign off"
        .CenterHeader = "STOCKTAKE " & Format(Date, "mmmm yyyy")
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = "&P of &N"
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .PrintGridlines = True
        .PrintTitleRows = "$1:$1"
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 0
        .PrintErrors = xlPrintErrorsDisplayed
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
    End With
    Application.PrintCommunication = True
    
    ActiveWorkbook.SaveAs ImportPath & "\" & "Stocktake Master " & Format(Date, "mmmm yyyy")
    ActiveWorkbook.Close
    
    'Create Node sheets
    
    For k = 1 To UBound(Nodes, 1)
        Workbooks.Add
        Sheets(1).Name = "Stocktake"
        
        Range("A1").Value = "Item Number"
        Range("B1").Value = "Description"
        Range("C1").Value = "Description 2"
        Range("D1").Value = "UOM"
        Range("E1").Value = "Vendor"
        Range("F1").Value = "Location"
        Range("G1").Value = "Section"
        Range("H1").Value = "Region"
        Range("I1").Value = "Category"
        Range("J1").Value = "Shelf/Bin"
        Range("K1").Value = "Previous Qty"
        Range("L1").Value = "Current Qty"
        Range("M1").Value = "Unit Cost"
        
        j = 2 'Let j be the row counter.  Hold which rwo to print the next record on. (Main list)
        For i = 1 To UBound(ImportData, 1)
            If ImportData(i, 12) = Nodes(k, 1) And ImportData(i, 1) = "PHYS. INVE" Then
                Cells(j, 1).Value = ImportData(i, 4)       'Item Number
                Cells(j, 2).Value = ImportData(i, 5)       'Description
                Cells(j, 3).Value = ImportData(i, 14)      'Description 2
                Cells(j, 4).Value = ImportData(i, 8)       'UOM
                Cells(j, 5).Value = ImportData(i, 9)       'Vendor
                Cells(j, 6).Value = ImportData(i, 6)       'Location
                Cells(j, 7).Value = ImportData(i, 11)      'Section
                Cells(j, 8).Value = ImportData(i, 12)      'Region
                Cells(j, 9).Value = ImportData(i, 13)      'Category
                Cells(j, 10).Value = ImportData(i, 7)      'Shelf/Bin
                Cells(j, 11).Value = ImportData(i, 10)     'Previous Qty
                Cells(j, 13).Value = ImportData(i, 15)     'Unit Cost
                j = j + 1
            End If
        Next i
        
        'Formating
        Range("A1:M1").Font.Bold = True
        Columns("A:A").ColumnWidth = 20
        Columns("A:A").NumberFormat = "0"
        Columns("B:B").ColumnWidth = 35
        Columns("C:C").ColumnWidth = 22
        Columns("D:D").ColumnWidth = 10
        Columns("E:E").ColumnWidth = 35
        Columns("F:G").ColumnWidth = 12
        Columns("F:F").EntireColumn.Hidden = True 'Location
        Columns("K:K").EntireColumn.Hidden = True 'Previous Quantity
        Columns("L:L").ColumnWidth = 10
        Columns("M:M").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"    'Accounting Format
        Columns("A:D").HorizontalAlignment = xlLeft
        
        'Sort by Description
        Columns("A:M").Sort Key1:=Range("B2"), Order1:=xlAscending, _
        Header:=xlYes, OrderCustom:=1, MatchCase:=False, _
        Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2:=xlSortNormal                        'Sort Data
        
        'Setup Print Preferences
        ActiveSheet.PageSetup.PrintArea = ""
        Application.PrintCommunication = False
        With ActiveSheet.PageSetup
            .LeftHeader = "Please use pen | Sheets are a GUIDE ONLY | 2 person to sign off"
            .CenterHeader = "STOCKTAKE " & Nodes(k, 1) & " " & Format(Date, "mmmm yyyy")
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = ""
            .RightFooter = "&P of &N"
            .LeftMargin = Application.InchesToPoints(0.7)
            .RightMargin = Application.InchesToPoints(0.7)
            .TopMargin = Application.InchesToPoints(0.75)
            .BottomMargin = Application.InchesToPoints(0.75)
            .HeaderMargin = Application.InchesToPoints(0.3)
            .FooterMargin = Application.InchesToPoints(0.3)
            .PrintComments = xlPrintNoComments
            .PrintQuality = 600
            .PrintGridlines = True
            .PrintTitleRows = "$1:$1"
            .Orientation = xlLandscape
            .PaperSize = xlPaperA4
            .FirstPageNumber = xlAutomatic
            .Order = xlDownThenOver
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 0
            .PrintErrors = xlPrintErrorsDisplayed
            .ScaleWithDocHeaderFooter = True
            .AlignMarginsHeaderFooter = True
        End With
        Application.PrintCommunication = True
            
        ActiveWorkbook.SaveAs ImportPath & "\" & "Stocktake " & Nodes(k, 1) & " " & Format(Date, "mmmm yyyy")
        ActiveWorkbook.Close
    Next k
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox (UBound(Nodes, 1) + 1 & " Files Created" & vbCrLf & "Files saved to " & ImportPath)
    
End Sub


Function GetNavData() As Boolean
    'Open ImportList
    Set myfile = Application.FileDialog(msoFileDialogOpen)
    With myfile
        .Title = "Choose the Navision export file"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            Exit Function
        End If
        FileSelected = .SelectedItems(1)
    End With
    SampleFilePath = FileSelected
    
    'Open File
    Workbooks.Open Filename:=SampleFilePath, Format:=2  'Format 2 is a Comma separated file (CSV)

    If Not (Range("A1").Value = "PHYS. INVE" And Range("B1").Value = "DEFAULT") Then
        MsgBox "Incorrect File"
        Exit Function
    End If
    GetNavData = True
End Function

Function GetStockSheetData() As Boolean
    'Open ImportList
    Set myfile = Application.FileDialog(msoFileDialogOpen)
    With myfile
        .Title = "Choose the StockSheet file to import"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            Exit Function
        End If
        FileSelected = .SelectedItems(1)
    End With
    SampleFilePath = FileSelected
    
    'Open File
    Workbooks.Open Filename:=SampleFilePath

    If Not (Range("A1").Value = "Item Number" And Range("B1").Value = "Description") Then
        MsgBox "Incorrect File"
        Exit Function
    End If
    GetStockSheetData = True
End Function


Sub Import_to_Master(control As IRibbonControl)
    
    Dim wbData As Workbook
    Dim Data As Variant
    Dim SS_Data As Variant
    Dim RowNum As Single
    Dim key As String
    
    Dim dicMASTER As Scripting.Dictionary
    
    'Main
    Application.ScreenUpdating = False
     
    Set wbData = ActiveWorkbook
    If CheckFormat_Master = False Then Exit Sub  'Checks that the active spreadsheet is a stocktake sheet.
    
    'Create Dictionary of all Stock Items
    RowNum = Range("A" & 65536).End(xlUp).Row  'Find the number of rows,  The columns are already known.
    Data = Range("A1:P" & RowNum)
    Set dicMASTER = New Scripting.Dictionary: dicMASTER.CompareMode = vbTextCompare
        For i = 2 To RowNum
        key = Data(i, 1)          'TYPE
        key = key & Data(i, 4)    'Item Number
        key = key & Data(i, 9)    'Location
        key = key & Data(i, 10)   'Section
        key = key & Data(i, 11)   'Node
        key = key & Data(i, 12)   'Category
        key = key & Data(i, 13)   'Shelf/Bin
        dicMASTER(key) = i
    Next i
    
    'Open Stocksheet to Import
    If GetStockSheetData = False Then Exit Sub 'Open Import file.  If incorrect file opened (i.e. Func=FALSE) then exit sub
    
    'Copy SS Data to Variant
    RowNum = Range("A" & 65536).End(xlUp).Row  'Find the number of rows,  The columns are already known.
    SS_Data = Range("A2:M" & RowNum)
    ActiveWorkbook.Close
    
    'Copy Counts to MASTER
    wbData.Activate
    For i = 1 To UBound(SS_Data, 1)
        key = "PHYS. INVE"           'TYPE
        key = key & SS_Data(i, 1)    'Item Number
        key = key & SS_Data(i, 6)    'Location
        key = key & SS_Data(i, 7)    'Section
        key = key & SS_Data(i, 8)    'Node
        key = key & SS_Data(i, 9)    'Category
        key = key & SS_Data(i, 10)   'Shelf/Bin

        Cells(dicMASTER(key), 15).Value = SS_Data(i, 12)
    Next i

    Application.ScreenUpdating = True


End Sub



Sub Export_to_Nav(control As IRibbonControl)
    
    Dim i As Single
    Dim RowNum As Single
    Dim Data As Variant
    Dim FilePath As String
    
    'Main Program
    If CheckFormat_Master = False Then
        MsgBox "Incorrect Format"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    FilePath = Application.ActiveWorkbook.Path
    
    'Copy Data to variant
    RowNum = Range("A" & 65536).End(xlUp).Row  'Find the number of rows,  The columns are already known.
    Data = Range("A1:P" & RowNum)
    
    'Open New Workbook
    Workbooks.Add
    
    'Copy in Data
    For i = 2 To UBound(Data, 1)
            Cells(i, 1).Value = Data(i, 1)        'PHYS. INVE
            Cells(i, 2).Value = Data(i, 2)        'Default
            Cells(i, 3).Value = Data(i, 3)        'Line Number
            Cells(i, 4).Value = Data(i, 4)        'Item Number
            Cells(i, 5).Value = Data(i, 5)        'Description
            Cells(i, 6).Value = Data(i, 9)        'Location
            Cells(i, 7).Value = Data(i, 13)       'Self/Bin
            Cells(i, 8).Value = Data(i, 7)        'UOM
            Cells(i, 9).Value = Data(i, 8)        'Vendor
            Cells(i, 10).Value = Data(i, 15)      'Current Qty
            Cells(i, 11).Value = Data(i, 10)      'Section
            Cells(i, 12).Value = Data(i, 11)      'Region
            Cells(i, 13).Value = Data(i, 12)      'Category
            Cells(i, 14).Value = Data(i, 6)       'Description 2
            Cells(i, 15).Value = Data(i, 16)      'Unit Cost
    Next i
    
    'Sort by PHYS. INVE then Line Number
    Columns("A:O").Sort Key1:=Range("A1"), Order1:=xlAscending, Key2:=Range("C1"), _
    Order2:=xlAscending, Header:=xlNo, OrderCustom:=1, MatchCase:=False, _
    Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, DataOption2:=xlSortNormal                        'Sort Data
    
    'Formatting
    Columns("A:O").NumberFormat = "General"
    
    'Save File
    ActiveWorkbook.SaveAs FilePath & "\" & "Stocktake Import " & Format(Date, "mmmm yyyy"), FileFormat:=xlCSV, CreateBackup:=False
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub



