Attribute VB_Name = "modUserForms"
Option Explicit

'
' Functions, contstants and types commonly used with UserForms and controls
'

'Public Function GetRange() As Range
'    Set GetRange = wksZipCodes.Range("A1").CurrentRegion
'    Set GetRange = GetRange.Offset(1).Resize(GetRange.Rows.Count - 1)
'End Function



Public Sub comboboxLoadFromTable(combobox As MSForms.combobox, tableName As String, columnName As String, _
                                    Optional viewLines As Long = 10, Optional selected As Long = 0)
    Dim rngTarget As Range
    
    ' Get the list of values to populate the combobox
    Set rngTarget = tableColumnFromHeader(tableName, columnName)
    
    ' Array containing the contents of the ComboBox
    combobox.List = rngTarget.value
    
    ' # of the the currently selected row (zero-based)
    combobox.ListIndex = selected
    
    ' # Rows to display under the selected row
    combobox.ListRows = viewLines
End Sub

Public Function comboboxValueFromColumnName(combobox As MSForms.combobox, tableName As String, columnName As String)
    ' Returns the value in the specified table at the row indicated by the combobox and
    '   the column with the specified name
    Dim nIndex As Long
    Dim nColumnIndex As Long
    Dim loTable As ListObject
    Dim value
    
    nIndex = combobox.ListIndex
    Set loTable = Range(tableName).ListObject
    nColumnIndex = loTable.ListColumns(columnName).Range.column
    comboboxValueFromColumnName = loTable.Range.Cells(nIndex + 2, nColumnIndex)
End Function

'Public Function GetValueFromTable(tableName As String, indexRow As Long, returnColumn As String) As String
'    Dim strValue As String
'
'    strValue =
'
'End Function

