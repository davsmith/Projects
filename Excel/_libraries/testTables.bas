Attribute VB_Name = "testTables"
Option Explicit

Public Const TABLE_NAME = "tblMicrosoftStock"


Sub testTableRangeFromName()
    Dim rng As Range
    
    Set rng = tableRangeFromName(TABLE_NAME)
    Debug.Print ("Data from table " + TABLE_NAME + " is at address " + rng.Address)
    
    Set rng = tableRangeFromName(TABLE_NAME + "[#ALL]")
    Debug.Print ("Complete table " + TABLE_NAME + " is at address " + rng.Address)
End Sub

Sub testTableListObjectFromName()
    Dim loTable As ListObject
    Dim strErrorTable As String
    
    strErrorTable = TABLE_NAME + "_Error"
    
    ' Test 1: Get a ListObject representing a table
    Set loTable = tableListObjectFromName(TABLE_NAME)
    Debug.Print "Table: " + loTable.Name + " has " + CStr(loTable.ListRows.Count) + " rows and " + CStr(loTable.ListColumns.Count) + " columns of data."

    ' Test 2: Try to retrieve a non-existent table
    Set loTable = tableListObjectFromName(strErrorTable)
    If loTable Is Nothing Then
        Debug.Print "Table " + strErrorTable + " doesn't exist"
    End If
End Sub

Sub testTableArrayFromName()
    Dim arrTable As Variant
    Dim nNumRows As Long
    Dim nNumColumns As Long
    Dim vValue As Variant
    
    arrTable = tableArrayFromName(TABLE_NAME)
    nNumRows = UBound(arrTable) - LBound(arrTable) + 1
    nNumColumns = UBound(arrTable, 2) - LBound(arrTable, 2) + 1
    vValue = arrTable(1, 1)
    Debug.Print "Array from table " + TABLE_NAME + " has " + CStr(nNumRows) + " rows and " + CStr(nNumColumns) + " columns of data."
End Sub

Sub testArrayDimensions()
    Dim nNumRows As Long
    Dim nNumCols As Long
    Dim arrTable As Variant
    Dim arrList(3 To 5) As Long
    
    arrTable = tableArrayFromName(TABLE_NAME)
    
    ' Test happy path
    arrayDimensions arrTable, nNumRows, nNumCols
    
    ' Test single dimension array (not a table)
    arrayDimensions arrList, nNumRows, nNumCols
    
    ' Test non-existent table
    arrTable = tableArrayFromName(TABLE_NAME + "_error")
    arrayDimensions arrTable, nNumRows, nNumCols
    
End Sub

Sub testTableDimensions()
    Dim nNumRows As Long
    Dim nNumCols As Long
    Dim arrTable As Variant
    Dim arrList(3 To 5) As Long
    
    ' Test happy path
    tableDimensions TABLE_NAME, nNumRows, nNumCols
    
    ' Test non-existent table
    arrayDimensions TABLE_NAME + "_error", nNumRows, nNumCols

End Sub

Sub testTableArrayDimensionCount()
    Dim arrList(3 To 5) As Long
    Dim arrTable As Variant
    Dim nNumDimensions As Long
    
    arrTable = tableArrayFromName(TABLE_NAME)
    
    ' The list should have 1 dimension
    nNumDimensions = arrayDimensionCount(arrList)
    
    ' The table should have 2 dimensions
    nNumDimensions = arrayDimensionCount(arrTable)
    
    
End Sub

Sub testTableArrayToName()
    Dim arr As Variant
    Dim loTable As ListObject
    Dim nRowIndex As Long
    Dim nColIndex As Long
    Dim origValue As Double
    
    ' Get an array to work with
    arr = tableArrayFromName(TABLE_NAME)
    
    ' Figure out the first item in the array
    nRowIndex = LBound(arr, 1)
    nColIndex = LBound(arr, 2)
    
    ' Store the original value
    origValue = arr(nRowIndex + 1, nColIndex + 1)
    
    ' Change the value in the array and put it back in the sheet
    arr(nRowIndex + 1, nColIndex + 1) = 999
    tableArrayToName TABLE_NAME, arr
    
    ' Restore the original value
    arr(nRowIndex + 1, nColIndex + 1) = origValue
    tableArrayToName TABLE_NAME, arr
    
End Sub


Sub testTableIndexFromColumnName()
    Dim nColumnIndex As Long
    
    nColumnIndex = tableIndexFromColumnName(TABLE_NAME, "Adj Close")
    nColumnIndex = tableIndexFromColumnName(TABLE_NAME, "Abadacus")
    nColumnIndex = tableIndexFromColumnName(TABLE_NAME, "")
    
End Sub

Sub testTableApplyFilterFormula()
    Dim strRule As String
    Dim strTable As String
    Dim rngCriteria As Range
    
    ' The data range and criteria range both require headers
    strTable = TABLE_NAME
    Set rngCriteria = Range("tblCriteria")
    
    ' Filter for Close > 300 (April 5 and April 7)
    strRule = "=E2>300"
    Range("tblCriteria").value = strRule
    tableApplyFilterFormula strTable, rngCriteria
    
    ' Filter using dates
    strRule = "=A2>=DATE(2023,1,1)"
    Range("tblCriteria").value = strRule
    tableApplyFilterFormula strTable, rngCriteria

    strRule = "=A2>=DATEVALUE(""March 15, 2023"")"
    Range("tblCriteria").value = strRule
    tableApplyFilterFormula strTable, rngCriteria

    ' Filter for specific text in the Comments
    strRule = "=H2=""Pluribus"""
    Range("tblCriteria").value = strRule
    tableApplyFilterFormula strTable, rngCriteria
    
    ' Filter for text in Comments using a wildcard
    strRule = "=SEARCH(""*Day*"",H2)"
    Range("tblCriteria").value = strRule
    tableApplyFilterFormula strTable, rngCriteria

    ' Comments AND closing price
    strRule = "=AND(E2>280, SEARCH(""*Day*"",H2))"
    Range("tblCriteria").value = strRule
    tableApplyFilterFormula strTable, rngCriteria
End Sub


Sub testTableLoopVisibleRows()
    Dim strTable As String
    Dim nNumRows As Long
    
    strTable = "tblMicrosoftStock"
    nNumRows = tableLoopVisibleRows(strTable, "printDateColumn")
End Sub

Sub testTableLoopAllRows()
    Dim strTable As String
    Dim nNumRows As Long
    
    strTable = "tblMicrosoftStock"
    nNumRows = tableLoopAllRows(strTable, "printCommentColumn", "printDateColumn")
End Sub


Sub printDateColumn(row As Range)
    Debug.Print (row.Cells(1, 1))
End Sub

Sub printCommentColumn(row As Range)
    Debug.Print (row.Cells(1, 8))
End Sub

Sub testTableClearFilters()
    Dim strTable As String
    Dim rngCriteria As Range
    Dim strRule As String
    
    '
    ' Test setup
    '
    
    ' The data range and criteria range both require headers
    strTable = TABLE_NAME
    Set rngCriteria = Range("tblCriteria")
    
    ' Filter for Close > 300 (April 5 and April 7)
    strRule = "=E2>300"
    Range("tblCriteria").value = strRule
    tableApplyFilterFormula strTable, rngCriteria
    
    '
    ' Test begin
    '
    
    ' Clear the existing filters
    tableClearFilters (TABLE_NAME)
    
    ' Call clear again to confirm no error w/o applied filters
    tableClearFilters (TABLE_NAME)
End Sub


Sub testTableCellValueFromRowColumn()
    Dim vCellValue As Variant
    Dim nRowIndex As Long
    Dim nColumnIndex As Long
    
    nRowIndex = 2
    nColumnIndex = 5
    
    ' Retrieve the value of the second data row, "Close" column
    vCellValue = tableCellValueFromRowColumn(TABLE_NAME, nRowIndex, nColumnIndex)
    Debug.Print ("The value in row " + CStr(nRowIndex) + ", column " + CStr(nColumnIndex) + " is " + CStr(vCellValue))
End Sub

Sub testtableCellValueFromColumnName()
    Dim vCellValue As Variant
    Dim nRowIndex As Long
    Dim strColumnName As String
    
    ' Test a known good value
    nRowIndex = 4
    strColumnName = "Adj Close"
    
    vCellValue = tableCellValueFromColumnName(TABLE_NAME, nRowIndex, strColumnName)
    Debug.Print ("The value in row " + CStr(nRowIndex) + ", in the " + strColumnName + " column is " + CStr(vCellValue))

    ' Check an error case
    strColumnName = "Flibberty"
    vCellValue = tableCellValueFromColumnName(TABLE_NAME, nRowIndex, strColumnName)
    If IsEmpty(vCellValue) Then
        Debug.Print ("Column " + strColumnName + " does not exist in table " + TABLE_NAME + ".")
    End If

End Sub

Sub testTableAddColumn()
    tableAddColumn TABLE_NAME, "TestColumn"
    tableAddColumn TABLE_NAME, "TestColumn", "Date"
    tableAddColumn TABLE_NAME, "TestColumn", "Wabasha"
End Sub


Sub testTableGetColumn()
    Dim nColumnIndex As Long
    
    nColumnIndex = tableGetColumn(TABLE_NAME, "Whats It", True)
    nColumnIndex = tableGetColumn(TABLE_NAME, "Whats It", False)
    nColumnIndex = tableGetColumn(TABLE_NAME, "No Column", False)
End Sub

Sub testTableGetParentFromCell()
    Dim strParentTable As String
    Dim rngCell As Range
    
    Set rngCell = Range(TABLE_NAME).Cells(5, 5)
    strParentTable = tableGetParentFromCell(rngCell)
    
    Set rngCell = Cells(1, Range(TABLE_NAME).CurrentRegion.Columns.Count + 3)
    strParentTable = tableGetParentFromCell(rngCell)
End Sub

''''''''''''''''''''''''''''''''''''''''''''''










Sub testGetTableColumn()
    Dim nNewColumnIndex As Integer
    
    nNewColumnIndex = GetTableColumn("xlaToolsTable1", "Flibbit", True)
End Sub


Sub buttonCycleRules()
    Dim arrTable
    Dim nFormulaIndex As Long
    Dim nIndex As Long
    Dim rngCriteria As Range
    
    arrTable = tableArrayFromName("tblRules")
    Set rngCriteria = Range("tblCriteria")
    nFormulaIndex = tableIndexFromColumnName("tblRules", "Formula")
    
    For nIndex = LBound(arrTable) To UBound(arrTable)
        rngCriteria.value = arrTable(nIndex, nFormulaIndex)
        Debug.Print (rngCriteria.Formula)
        tableApplyFilterFormula "tblMicrosoftStock", Range("tblCriteria")
    Next
End Sub

Sub buttonPopulateFormulaFromCurrentRow()
    Dim nCurrentRow As Long
    Dim loTable As ListObject
    Dim strRule As String
    
    Dim strTable As String
    Dim rngCriteria As Range
    
    ' The data range and criteria range both require headers
    strTable = TABLE_NAME + "[#All]"
    Set rngCriteria = Range("tblCriteria[#All]")
    
    ' Figure out the current row from cursor
    nCurrentRow = ActiveCell.row
    
    ' Retrieve the string in the Formula column
    Set loTable = tableListObjectFromName("tblRules")
    strRule = loTable.DataBodyRange.Cells(nCurrentRow - 1, loTable.ListColumns("Formula").Index)
    
    ' Paste the formula string into the Criteria value
    Range("tblCriteria").value = strRule
    
    ' Apply the filter
    tableApplyFilterFormula strTable, rngCriteria
End Sub

Sub testStringEndsWith()
    Dim bEndsWith As Boolean
    
    bEndsWith = stringEndsWith("day", "Happy Birthday")
    bEndsWith = stringEndsWith("day", "Happy Birthday!")
    bEndsWith = stringEndsWith("flib", "Fliberty jibbit")
    bEndsWith = stringEndsWith("I was on my way to school", "school")
    bEndsWith = stringEndsWith("school", "I was on my way to school")

End Sub


Sub buttonApplyCurrentFilter()
    Dim strTable As String
    Dim rngCriteria As Range
    
    ' The data range and criteria range both require headers
    strTable = TABLE_NAME + "[#All]"
    Set rngCriteria = Range("tblCriteria[#All]")
    
    tableApplyFilterFormula strTable, rngCriteria
End Sub

Sub buttonClearFilters()
    tableClearFilters (TABLE_NAME)
End Sub


