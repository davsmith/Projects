Attribute VB_Name = "modTables"
Option Explicit

' A collection of functions demonstrating and supporting working with
' tables in VBA
'
' Some of the functions are shorter or more explicit than necessary to
' demonstrate the syntax or concept
'
' See OneNote for more details: https://tinyurl.com/vmmzka2b
'
' Dave Smith
' 6/4/2023
'
'
Function tableRangeFromName(tableName) As Range
    ' Returns the Range object associated with a table
    ' Has tests
    Dim rng As Range
    
    On Error GoTo fail
        Set rng = Range(tableName)
        
        Set tableRangeFromName = rng
        Exit Function
fail:
End Function

Function tableListObjectFromName(tableName) As ListObject
    ' Returns the ListObject associated with a table
    ' Has tests
    Dim loTable As ListObject
    
    On Error GoTo fail
        Set loTable = Range(tableName).ListObject
        
        Set tableListObjectFromName = loTable
        Exit Function
fail:
End Function

Function tableArrayFromName(tableName) As Variant
    ' Returns the data range of a table as an array
    ' Has tests
    Dim loTable As ListObject
    Dim arTable As Variant
    
    On Error GoTo fail
        Set loTable = Range(tableName).ListObject
        arTable = loTable.DataBodyRange.value
        tableArrayFromName = arTable
        Exit Function
fail:
End Function

Sub arrayDimensions(arrTable As Variant, ByRef numRows As Long, ByRef numColumns As Long)
    ' Returns the number of rows and columns in a table
    '
    ' arrTable must be a two-dimensional array
    ' numRows/numColumns must be passed as variables, not constants
    '
    ' Has tests
    Dim nNumDimensions As Long
    
    numRows = -1
    numColumns = -1
    
    If (Not IsEmpty(arrTable)) And (arrayDimensionCount(arrTable) = 2) Then
        ' Assigns the row and column count of a table to the variables passed by reference
        numRows = UBound(arrTable) - LBound(arrTable) + 1
        numColumns = UBound(arrTable, 2) - LBound(arrTable, 2) + 1
    End If
End Sub

Sub tableDimensions(tableName As String, ByRef numRows As Long, ByRef numColumns As Long)
    Dim arrTable As Variant
    
    arrTable = tableArrayFromName(tableName)
    arrayDimensions arrTable, numRows, numColumns
End Sub


Function arrayDimensionCount(arr As Variant) As Long
    ' Determine the number of dimensions of an array by testing for the
    '  lower bound of the dimension until there is an error
    '
    ' Has tests
    Dim nDimensionCount As Long
    Dim nErrorCheck As Integer
    
    On Error GoTo last_dimension
        For nDimensionCount = 1 To 60
            nErrorCheck = LBound(arr, nDimensionCount)
        Next
    
last_dimension:
    arrayDimensionCount = nDimensionCount - 1
End Function

Sub tableArrayToName(tableName As String, arr As Variant)
    ' Populates the specified table with data in the array
    ' Has tests
    Dim loTable As ListObject
    
    Set loTable = tableListObjectFromName(tableName)
    If Not loTable Is Nothing Then
        loTable.DataBodyRange = arr
    End If
End Sub

Function tableIndexFromColumnName(tableName As String, columnName As String) As Long
    ' Returns the index of the specified column name within the specified table
    '   or -1 if the column name doesn't exist
    ' Has tests
    Dim nIndex As Long
    Dim loTable As ListObject
    
    tableIndexFromColumnName = -1
    
    On Error GoTo fail
        Set loTable = tableListObjectFromName(tableName)
        nIndex = loTable.ListColumns(columnName).Range.column
        tableIndexFromColumnName = nIndex
        Exit Function
fail:
End Function

Public Function tableColumnFromHeader(tableName As String, referenceColumn As String) As Range
    ' Returns a range object containing the data from the specified column
    Dim rng As Range
    Dim rngString As String
    
    Set rng = Range(tableName + "[" + referenceColumn + "]")
    Set tableColumnFromHeader = rng
End Function

Sub tableApplyFilterFormula(tableName As String, criteriaRange As Range)
' Applies an advanced filter using the formula in the specified criteria range
'
' Typically the criteria range is a 2-row, 1-column range with
'   - A heading/label in the first row
'   - A formula in the second row
'
' See <example sheet> for examples
'
' Has tests
    Dim rngTable As Range
    Dim rngCriteria As Range
    
    ' Ensure range includes headers and data
    Set rngTable = tableRangeFromName(tableName)
    Set rngTable = rngTable.ListObject.Range
    
    ' Ensure range includes headers and data
    Set rngCriteria = criteriaRange.ListObject.Range
    
    rngTable.AdvancedFilter Action:=xlFilterInPlace, criteriaRange:=rngCriteria, Unique:=False

End Sub

Function tableLoopVisibleRows(tableName As String, Optional functionName As String) As Long
    ' Traverses all visible (filtered) rows in the specified table
    ' If specified, functionName is called for each row
    '
    ' This is a demonstration of how to loop across visible cells, more
    ' than a useful function on its own
    ' Has tests
    Dim rngRow As Range
    Dim rngArea As Range
    Dim rngTable As Range
    Dim rngVisibleRows As Range
    Dim nRowCounter As Integer
    Dim nAreaTally As Integer
    
    Set rngTable = Range(tableName)
        
    ' Iterate only through the filtered rows
    nRowCounter = 0
    Set rngVisibleRows = rngTable.Rows.SpecialCells(xlCellTypeVisible)
    
    For Each rngRow In rngVisibleRows.Rows
        If functionName <> "" Then
            Run functionName, rngRow
        End If
        nRowCounter = nRowCounter + 1
    Next
    
    ' The .count property isn't correct so tally row by row,
    ' cross check against the sum of rows in each of the visible areas
    nAreaTally = 0
    For Each rngArea In rngVisibleRows.Areas
        nAreaTally = nAreaTally + rngArea.Rows.Count
    Next
    
    Debug.Print ("*** Iterated through " + CStr(nRowCounter) + " rows.")
    tableLoopVisibleRows = nRowCounter
End Function

Function tableLoopAllRows(tableName As String, Optional visibleFunction As String, Optional hiddenFunction As String) As Long
    ' Traverses all rows in the specified table values detecting whether it is visible or hidden (filtered out)
    ' If specified, visibleFunction is run on the rows that are visible
    ' If specified, hiddenFunction is run on the rows that are hidden
    '
    ' This is a demonstration of how to loop across visible cells, more
    ' than a useful function on its own
    ' Has tests
    Dim rngTable As Range
    Dim rngRow As Range
    Dim i As Integer
    Dim nRowCounter As Integer
    
    Set rngTable = Range(tableName)
        
    ' Iterate through every row detecting whether it is visible
    nRowCounter = 0
    For Each rngRow In rngTable.Rows
        If rngRow.Hidden Then
            If hiddenFunction <> "" Then
                Run hiddenFunction, rngRow
            End If
        Else
            nRowCounter = nRowCounter + 1
            If visibleFunction <> "" Then
                Run visibleFunction, rngRow
            End If
        End If
    Next
    
    tableLoopAllRows = nRowCounter
End Function

Sub tableClearFilters(tableName As String)
    ' Clears any filters from the specified table
    ' has tests
    Dim rngTable As Range
    
    ' Trying to clear filters when none are turned
    ' on results in a bogus error, so turn off
    ' error handling
    Set rngTable = tableRangeFromName(tableName)
    On Error Resume Next
    rngTable.Parent.ShowAllData
    On Error GoTo 0
End Sub

Function tableCellValueFromRowColumn(tableName As String, row As Long, column As Long) As Variant
    ' Example of retrieving an individual value from a table,
    '   providing the row and column index
    ' Has tests
    Dim loTable As ListObject
    Dim strValue As String
    
    Set loTable = tableListObjectFromName(tableName)
    tableCellValueFromRowColumn = loTable.DataBodyRange.Cells(row, column)
End Function

Function tableCellValueFromColumnName(tableName As String, row As Long, columnName As String) As Variant
    ' Example of retrieving an individual value from a table,
    '   providing the row and name of the column
    ' Has tests

    Dim nColumnIndex As Long
    Dim loTable As ListObject
    
    On Error GoTo fail
        Set loTable = tableListObjectFromName(tableName)
        nColumnIndex = loTable.ListColumns(columnName).Index
        tableCellValueFromColumnName = tableCellValueFromRowColumn(tableName, row, nColumnIndex)
        Exit Function
fail:
    
End Function


Function tableAddColumn(strTableName As String, strNewColumn As String, Optional strBefore As String = "") As Integer
    ' Adds a column to the specified table either before the column named in strBefore, or
    '  at the end of the table.
    ' Has tests
    Dim nNewColumnIndex As Integer
    Dim strExistingColumn As String
    Dim loTable As ListObject
       
    Set loTable = Range(strTableName).ListObject
    
    If strBefore = "" Then
        nNewColumnIndex = loTable.ListColumns.Count + 1
    Else
        nNewColumnIndex = tableColumnIndexFromColumnName(strTableName, strBefore)
        If nNewColumnIndex < 0 Then
            nNewColumnIndex = loTable.ListColumns.Count + 1
        End If
    End If
    
    loTable.ListColumns.Add(nNewColumnIndex).Name = strNewColumn
    
    tableAddColumn = nNewColumnIndex
End Function

Function tableGetColumn(strTableName As String, strColumnName As String, Optional bCreate As Boolean = False) As Integer
    ' Returns the index number of the specified column, with the option to create
    '    the column if it doesn't exist
    ' Has tests
    Dim nColumnIndex As Integer
    
    nColumnIndex = tableColumnIndexFromColumnName(strTableName, strColumnName)
    If (nColumnIndex < 0) And (bCreate) Then
        nColumnIndex = tableAddColumn(strTableName, strColumnName)
    End If
    
    tableGetColumn = nColumnIndex
End Function

Function tableGetParentFromCell(cell As Range) As String
    ' Returns the name of the table containing the specified cell
    '    or an empty string if the cell is not part of a table
    ' Has tests
    Dim tableName As String
    
    tableName = ""
    On Error Resume Next
    tableName = cell.ListObject.Name
    On Error GoTo 0
    
    tableGetParentFromCell = tableName
End Function


