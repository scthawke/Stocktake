Attribute VB_Name = "STK3_Check_FileFormats"
'*****************
'* CHECK FORMAT  *
'*****************
'These functions check the contents of a file to ensure it is what we think it is.

'Stocktake MASTER
'-----------------
'This is the MASTER stocktake document held by finance.  All counts from each node need to be combined to this document prior to import into Nav.
Function CheckFormat_Master() As Boolean
    If Cells(1, 1).Value = "Type" And _
            Cells(1, 2).Value = "Default" And _
            Cells(1, 3).Value = "Line Number" And _
            Cells(1, 4).Value = "Item Number" And _
            Cells(1, 5).Value = "Description" And _
            Cells(1, 6).Value = "Description 2" And _
            Cells(1, 7).Value = "UOM" And _
            Cells(1, 11).Value = "Region" And _
            Cells(1, 15).Value = "Current Qty" Then
                CheckFormat_Master = True   'Correct Format
    Else
                CheckFormat_Master = False 'Incorrect Format
    End If
End Function

'Stocktake Sheet
'-----------------
'This is the stocktake document issued to each node.  The document may be printed, however if it is returned electronically,
'counts can be automatically filled into the MASTER documnt.
Function CheckFormat_StockSheet() As Boolean
    If Cells(1, 1).Value = "Item Number" And _
            Cells(1, 2).Value = "Description" And _
            Cells(1, 3).Value = "Description 2" And _
            Cells(1, 4).Value = "UOM" And _
            Cells(1, 8).Value = "Region" And _
            Cells(1, 12).Value = "Current Qty" Then
                CheckFormat_StockSheet = True   'Correct Format
    Else
                CheckFormat_StockSheet = False 'Incorrect Format
    End If
End Function

'Navision (NAV) stock export file
'---------------------------------
'This is the file exported from Navision on which all Stocktake sheets are based.
Function CheckFormat_NAV() As Boolean
    If Cells(1, 1).Value = "PHYS. INVE" And _
            Cells(1, 2).Value = "DEFAULT" And _
            Cells(1, 15).Value <> "" And _
            Cells(1, 16).Value = "" Then
                CheckFormat_NAV = True   'Correct Format
    Else
                CheckFormat_NAV = False 'Incorrect Format
    End If
End Function

'ORCA Scan file
'---------------
'This is the file exported from the ORCA scan phone app.
Function CheckFormat_ORCA() As Boolean
    If Cells(1, 1).Value = "ITEM" And _
            Cells(1, 2).Value = "BARCODE" And _
            Cells(1, 5).Value = "DESCRIPTION" Then
                CheckFormat_wbImport = True   'Correct Format
    Else
                CheckFormat_wbImport = False 'Incorrect Format
    End If
End Function

