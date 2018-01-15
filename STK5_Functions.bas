Attribute VB_Name = "STK5_Functions"
'Row Number
'-----------
'Determines the number of used rows within a column of a worksheet.  Requires the column number as input.
Function RowNum(Col As Single) As Double
    Dim RowMax As Double
    'Determine the number of rows
    If ActiveSheet.UsedRange.Rows.Count > 65536 Then RowMax = 1048576 Else RowMax = 65536    'Set RowMax
    RowNum = Cells(RowMax, Col).End(xlUp).Row
End Function


