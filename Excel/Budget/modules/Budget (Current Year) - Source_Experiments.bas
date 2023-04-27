Attribute VB_Name = "Experiments"
Sub ApplyFilter()
    Range("tblTransactions[#All]").AdvancedFilter Action:=xlFilterInPlace, _
        CriteriaRange:=Sheets("Criteria").Range("A1:W2"), Unique:=False
End Sub


Sub ClearFilter()
    Sheets("Transactions").ShowAllData
End Sub

Sub ListVisibleValues()
    Dim ws As Worksheet
    Dim row As Range
    Dim transactions As Range
    Dim visible_rows As Range
    
    
    Set transactions = Range("tblTransactions")
    Set visible_rows = transactions.Rows.SpecialCells(xlCellTypeVisible)
    
    For Each row In visible_rows.Areas
        Debug.Print (row.Address + ":" + CStr(row.Cells(4).Value))
        row.Cells(4).Value = "Something on the Staging account"
    Next
End Sub
