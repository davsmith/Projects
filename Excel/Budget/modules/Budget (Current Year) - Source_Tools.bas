Attribute VB_Name = "Tools"
'****************************************
'* Excel Macro Tools
'*
'* 4/15/2023:
'*  - Fixed capitalization of property names
'*  - Made its and comments to pivot table macros
'*
'* 3/26/2023:
'*  - Checked in most up-to-date version of tools.bas
'*  - Mostly capitalization changes, with a few bullets added under 12/9/2015
'*  - Added OneDrive section
'*  - Added IsRemotePath function

'*
'* 12/9/2015:
'*  - Added FileExists
'*  - Added ReadFileToString
'*  - Modified Between to take a parameter to allow/reject partial matches.
'*  - Added PtrSafe to OS calls to make 64-bit compliant
'*  - Changed FSO to generic object instead of FileSystemObject
'*  - Added WriteStringToFile
'*
'* 12/28/2013:
'*  - Added GetLetterFromIndex
'*
'* 12/7/2013:
'*  - Added PurgeTempSheets
'*  - Added RenameSheet
'*  - Added GetTable
'*  - Added InsertAt parameter to GetColumnByHeading
'*
'* 11/19/2013:
'*  - Added AutoRange
'*  - Added DeleteSheets
'*
'* 11/9/2013:
'*  - Added docs on how to reference the Visual Basic IDE
'*  - Made sheetname parameter optional on GetDataSheet.
'*
'* 9/1/2013:
'*  - Added SuperCStr to convert a null object to ""
'*  - Added optional columns to SortOnColumnName
'*  - Modifed behavior of Between where one token (start or end) exists, but other does not.
'*  - Added ReplaceChars
'*  - Added ExtractNumberFromString
'*  - Modified SuperTrim
'*  - Added HTTPGet and URLEncode back into base tools
'*  - Added ExportGPX file (doesn't work yet)
'*  - Added IsPhoneNumber And IsURL
'*
'* 8/26/2013:
'*  - Modified SuperTrim to trim all non-alphanumeric characters from beginning and end.
'*
'* 8/24/2013:
'*  - Combined changes from other workbooks (Jubak, StatementParser, CarChargers)
'*  - Added Macro Helper section
'*  - Included changes to GetWebPage to make file dump optional
'*  - Added ExtractNumberFromString
'*  - Added IsPhoneNumber and IsURL to RegEx section
'*
'* 8/5/2013:
'*  - Modified FormatColumnWidths to take a range, rather than a worksheet, and made array parameter optional
'*
'* 7/14/2013:
'*  - Deprecated Tokenize function for Split
'*  - Overhauled documentation and formatting.
'*  - Added ShowNamedColumns and HideNamedColumns
'*  - Overhauled GeoCoding section to use Bing instead of Yahoo.
'*
'*
'* 6/17/2012:
'*  - *** Changed the return value of After function to vbNull if the delimiter is not found ***
'*  - Added IsNameValuePair
'*  - Updated the last row number in FindLast from 65536 to 1048576
'*  - Added TokenizeToCollection to accomodate the fact that the array returned
'*    by Tokenize cannot be redimensioned.
'*  - Added ReplaceCollectionElement
'*  - Added CollectionToString
'*  - Added Name/Value pair function definitions (NVPxxxx_)
'*  - Added section for working with GUIDs (IsGUID)
'*  - Added Append argument to GetColumnByHeading
'*  - Added FormatColumnWidths
'*
'* 3/7/2011:
'*  - Added external reference to Sleep
'*
'* 1/8/2011:
'*  - Added geocoding and regular expression function sets
'*
'* 1/3/2011:
'*  - Added comment about block commenting.
'*
'* 11/15/2010:
'*  - Added GroupByIndent
'*
'* 11/9/2010:
'*  - Added HexAnd
'*  - Added FixupCRLF
'*  - Added MakeDebugString
'*
'* 7/11/2010:
'*  - Added GetWorkbook
'*
'* 5/29/2010:
'*  - Added Path tools section, including:
'*      GetPath
'*      GetEntryName
'*      GetBaseName
'*      GetExtension
'*  - Added Ini section, including:
'*      ReadIniFileString
'*      WriteIniFileString
'*
'* 2/9/2010:
'*  - Added UNameWindows
'*
'* 12/04/09:
'*  - Added SuperTrim function
'*
'* 10/19/09:
'*  - Added Names section
'*  - Added PurgeTempNames function
'*  - Added NameExists function
'*
'* 10/18/09:
'*  - Changed MakePivot from a subroutine to a function
'*  - Added optional "After" parameter to FindValueInColumn, and
'*      changed the return type to long
'*  - Added InRange function to Set group.
'*
'* 1/1/09:
'*  - Added Collection section
'*  - Added ClearCollection function
'*  - Added SetPivotFilterValues function
'*  - Added PersistSettings function
'*
'* Updated 12/20/08
'*  - Added SortOnColumnName
'*  - Added MakePivot
'*  - Added comments to deprecate some of the "Find Column" functions
'*
'* Exported:
'*  - 6/2/08
'*
'* Updated 6/2/08
'*  - Added FormatRangeAsTable
'*
'* Updated 2/17/08
'*  - Added Between function
'*  - Added tips for objects/classes
'*  - Added section for class tools, but functionality is not yet clean
'*
'* Updated 2/10/08
'*  - Added Web Tools Section
'*  - Added GetWebPage function
'*  - Added external declarations
'*  - Added GetTempPath and GetTempFile functions
'*
'* Updated 11/22/07
'*  - Added Tokenize
'*  - Added Clipboard functions
'*
'* Updated 11/18/07
'*  - Combined standard and Pictures.xls Tools modules
'*
'* Updated 12/4/05
'*


' Some functions which would be interesting to write:
'   - parse a vba file and list all functions and subroutines with surrounding comments
'   - parse a function and list all variables which are/aren't used in the function
'

Option Explicit
Option Base 1

'
' Enumerated types
'
Public Enum etQuoteType
    enDoubleQuote = 0
    enSingleQuote = 1
End Enum

Type utNameValue
    szName As String
    varValue As Variant
End Type

Public Enum dsPersistType
    dsSave = 1
    dsRestore = 2
    dsFree = 3
End Enum

Public Enum dsFilterSettings
    dsOff = 0
    dsOn = 1
    dsToggle = 2
    dsOffExclusive = 3
    dsOnExclusive = 4
End Enum

'
' Constants
'
Public Const xlToggle = 2
Public Const xlOffExclusive = 3
Public Const xlOnExclusive = 4

Public Const MAX_ROW = 1048576

'
' Externally declared functions
'
Declare PtrSafe Function OS_GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal _
    nBufferLength As Long, _
    ByVal lpBuffer As String) As Long

Declare PtrSafe Function OS_GetTempFileName Lib "kernel32.dll" _
    Alias "GetTempFileNameA" (ByVal _
    lpszPath As String, _
    ByVal lpPrefixString As String, _
    ByVal wUnique As Long, _
    ByVal lpTempFileName As String) As Long


Private Declare PtrSafe Function GetPrivateProfileString Lib "Kernel32" Alias _
    "GetPrivateProfileStringA" ( _
        ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, _
        ByVal lpDefault As String, _
        ByVal lpReturnedString As String, _
        ByVal nSize As Long, _
        ByVal lpFileName As String) As Long

Private Declare PtrSafe Function WritePrivateProfileString Lib "Kernel32" Alias _
    "WritePrivateProfileStringA" ( _
        ByVal lpApplicationName As String, _
        ByVal lpKeyName As Any, _
        ByVal lpString As Any, _
        ByVal lpFileName As String) As Long

Declare PtrSafe Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

Declare PtrSafe Function QueryPerformanceCounter Lib "Kernel32" (X As Currency) As Boolean
Declare PtrSafe Function QueryPerformanceFrequency Lib "Kernel32" (X As Currency) As Boolean
Declare PtrSafe Function GetTickCount Lib "Kernel32" () As Long
Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long

Function UNameWindows() As String
    UNameWindows = Environ("USERNAME")
End Function



''*****************************
'       Workbook Tools
'.*****************************

Function RenameSheet(FromSheet As String, ToSheet As String, Optional Displace As Boolean = False)
    Dim nError As Long
    Dim bRetVal As Boolean
    Dim sError As String
    Dim nRenameCounter As Long
    Dim wks As Worksheet
    
    bRetVal = False

    If (Displace) Then
        On Error Resume Next
            Err.Clear
            nRenameCounter = 2
            While Err.Number = 0
                Set wks = Sheets(ToSheet + "(" + CStr(nRenameCounter) + ")")
                nRenameCounter = nRenameCounter + 1
            Wend
            Sheets(ToSheet).name = ToSheet + "(" + CStr(nRenameCounter - 1) + ")"
        On Error GoTo 0
    End If

    On Error Resume Next
        Sheets(FromSheet).name = ToSheet
        nError = Err.Number
        sError = Err.Description
    On Error GoTo 0
    
    If (nError = 0) Then
        bRetVal = True
    End If
    
    Select Case nError
        Case 9:
            sError = "Source sheet not found"
    End Select
    
    If (nError > 0) Then
        Debug.Print ("Error of sheet rename: " + sError)
    End If
    
    If (Displace) Then
    End If

EXIT_RENAMESHEET:
    RenameSheet = bRetVal
End Function


Sub PurgeTempSheets(Optional sPrefix As String = "TMP_")
    Dim n As Long

    For n = Sheets.Count To 1 Step -1
        If (Left(Sheets(n).name, Len(sPrefix)) = sPrefix) Then
            Debug.Print ("Purging temporary sheet " + Sheets(n).name)
            Application.DisplayAlerts = False
            Sheets(n).Delete
            Application.DisplayAlerts = True
        End If
    Next
End Sub


Sub DeleteSheets(SheetList As String)
    Dim sSheets() As String
    Dim n As Long
    Dim wks As Worksheet
    Dim sSheetName As String
    
    sSheets = Split(SheetList, ";")
    On Error Resume Next
    For n = 0 To UBound(sSheets)
        Application.DisplayAlerts = False
        Sheets(sSheets(n)).Delete
        Application.DisplayAlerts = True
    Next
    On Error GoTo 0
End Sub


'
' Saves the current session's settings so they can be restored when a macro is completed.
' The settings are stored in a static array, and returns an index so that if the function
' is called multiple times, the correct values can be restored.
'
'
Function PersistSettings(Optional dsAction As dsPersistType = dsSave, Optional nIndex As Long = 1) As Long
    Static colSettingsBagCollection As Collection
    Static nCount As Long
    Static colSettings As Collection
    Dim rng As Range
    
    Select Case dsAction
        Case dsSave:
            nCount = nCount + 1
            If colSettingsBagCollection Is Nothing Then
                Set colSettingsBagCollection = New Collection
            End If
            
            Set colSettings = New Collection
            colSettings.Add Selection.Address, "Selection"
            colSettings.Add ActiveSheet, "Worksheet"
            colSettings.Add ActiveWorkbook, "Workbook"
            colSettings.Add Application.ScreenUpdating, "ScreenUpdating"
            
            colSettingsBagCollection.Add colSettings
        Case dsRestore
            Set colSettings = colSettingsBagCollection(nIndex)
            colSettings("Worksheet").Activate
            Set rng = Range(colSettings("Selection"))
            rng.Select
            Application.ScreenUpdating = colSettings("ScreenUpdating")
    End Select
        
    PersistSettings = nCount
End Function


'
' Receives a cell as input and returns the range the input cell
' falls within
'
Function AutoRange(cell As Range, Optional RowCount As Long, Optional ColCount As Long) As Range
    Dim rng As Range
    Dim nLastRow As Long
    Dim nLastCol As Long

    Dim wks As Worksheet
    Dim rngLastRow As Range
    
    Set wks = cell.Worksheet
    
    nLastRow = cell.End(xlDown).Row
    nLastCol = cell.End(xlToRight).Column
    
    If (nLastRow = wks.Rows.Count) Then
        RowCount = 1
    Else
        RowCount = nLastRow - cell.Row + 1
    End If
    
    If (nLastCol = wks.Columns.Count) Then
        ColCount = 1
    Else
        ColCount = nLastCol - cell.Column + 1
    End If
    
    Set rng = cell.Resize(RowCount, ColCount)
    Set AutoRange = rng
End Function


'
' Clears out the range within the specified sheet, starting at the specified row
' The formatting is preserved.
'
' Made StartRow argument optional (7/20/2013)
'
Sub ResetData(strDataSheetName As String, Optional StartRow As Integer = 1)
    Dim datasheet As Worksheet
    Dim currentSheet As Worksheet
    Dim lastRow As Integer
    Dim currentScreenSetting As Boolean
    Dim nVisible As Long
    
    ' Turn off screen updates
    currentScreenSetting = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    ' Retrieve the data sheet
    Set datasheet = GetDataSheet(strDataSheetName, False)
    If Not (datasheet Is Nothing) Then
        Set currentSheet = ActiveSheet
        
        nVisible = datasheet.Visible
        datasheet.Visible = xlSheetVisible
        datasheet.Select
        lastRow = FindRow(datasheet, "*")
        
        If lastRow > StartRow Then
            Range(StartRow & ":" & lastRow).Delete
            Range(StartRow & ":" & lastRow).ClearContents
        Else
            Range(StartRow & ":" & StartRow).Delete
            Range(StartRow & ":" & StartRow).ClearContents
        End If
        
        datasheet.Visible = nVisible
        
        If (currentSheet.Visible = xlSheetVisible) Then
            currentSheet.Select
        End If
    End If
    
    Application.ScreenUpdating = currentScreenSetting
End Sub


'
' Protects the active worksheet
'
' UPDATED 7/20/2013:  Replaced hard coded range with row and column count
'
Sub ProtectMe()
    ActiveSheet.Protect Password:="Test"
    Cells(1, 1).Resize(Rows.Count, Columns.Count).Locked = True
End Sub


'
' Returns the worksheet with the specified name
'
' [in]  strSheetName    The name of the worksheet to retrieve
' [in]  bCreate         Create the worksheet if it doesn't exist
'
' If no sheet name is specified, with bCreate a temporary name is generated
'
Function GetDataSheet(Optional strSheetName As String, Optional bCreate As Boolean = False)
    Dim aSheet As Worksheet
    Dim bScreenUpdating As Boolean
    Dim shtCurrentWorksheet As Worksheet
    Dim nTickCount As Long
    
    If (strSheetName = "") Then
        strSheetName = "TMP_" + CStr(GetTickCount())
        bCreate = True
    End If
    
    ' Check if the specified worksheet name exists
    On Error Resume Next
    Set aSheet = Sheets(strSheetName)
    On Error GoTo 0
    
    ' If it doesn't, and the create argument is true
    If (aSheet Is Nothing) And bCreate Then
        ' Store the currently active sheet, and screen updating
        Set shtCurrentWorksheet = ActiveSheet
        
        ' Add a new sheet, which activates it
        Set aSheet = Sheets.Add
        aSheet.name = strSheetName
        
        ' Activate the previously active sheet if appropriate.
        If shtCurrentWorksheet.Visible = xlSheetVisible Then
            shtCurrentWorksheet.Select
        End If
    End If
    
    Set GetDataSheet = aSheet
End Function


'
' Returns the workbook with the specified name
'
' [in]  strWorkbook    The name of the worksheet to retrieve
' [in]  bCreate        Create the worksheet if it doesn't exist
'
'
Function GetWorkbook(strWorkbook As String, Optional bCreate As Boolean = False, Optional strPath As String = "")
    Dim aSheet As Worksheet
    Dim bScreenUpdating As Boolean
    Dim wkb As Workbook
    Dim wkbCurrent As Workbook
    Dim strDestFile As String
        
    On Error Resume Next
    Set wkb = Workbooks(strWorkbook)
    On Error GoTo 0
    
    If (wkb Is Nothing) And bCreate Then
        If (strPath = "") Then
            strDestFile = strWorkbook
        Else
            strDestFile = strPath + "\" + strWorkbook
        End If
        
        Set wkbCurrent = ActiveWorkbook
        bScreenUpdating = Application.ScreenUpdating
        Application.ScreenUpdating = False
        Set wkb = Workbooks.Add
        wkb.SaveAs (strDestFile)
        wkbCurrent.Activate
        
        Application.ScreenUpdating = bScreenUpdating
    End If
    
    Set GetWorkbook = wkb
End Function

'
'
'
'Sub AddNewRow()
'    Dim SelectedCol As Integer
'    Dim RowToCopy As Integer
'    Dim NewRow As Integer
'
'    Application.EnableEvents = False
'
'    ' Find the last row
'    RowToCopy = FindRow(ActiveSheet, "*")
'    Rows(RowToCopy & ":" & RowToCopy).Copy '.Select
'    'Selection.Insert Shift:=xlDown
'
'    ' select first cell in dest row and paste
'    Range("A" & (RowToCopy + 1)).Select
'    ActiveSheet.Paste
'
'    Application.EnableEvents = True
'End Sub


''******************
'  Search & Lookup
'.******************

Function GetTable(rngSearch As Range, sHeaders As Variant, Optional sOptional As Variant, Optional bIncludeBlankWithStar As Boolean = False) As Range
    Dim rngWorking As Range
    Dim ty As String
    Dim sFirstHeader As String
    Dim nNumHeadersFound As Long
    Dim n As Long
    Dim o As Long
    Dim bFoundRequiredHeaders As Boolean
    Dim rngFirstFind As Range
    Dim rngFinal As Range
    Dim xlSearchType As XlLookAt
    Dim xlSearchIn As XlFindLookIn
    Dim nNumCharsToCompare As Long
    
    bFoundRequiredHeaders = True
        
    ' Use find to locate the first element of the table
    sFirstHeader = sHeaders(0)
    xlSearchType = xlWhole
    xlSearchIn = xlFormulas
    
    If (InStr(sFirstHeader, "*")) Then
        xlSearchType = xlPart
    End If
    
    ' If we're searching for a number, a couple of parameters are diffent
    If (IsNumeric(sFirstHeader)) Then
        xlSearchType = xlPart
        xlSearchIn = xlValues
    End If
    
    Set rngWorking = rngSearch.Find(What:=sFirstHeader, After:=rngSearch.Cells(1, 1), LookIn:=xlSearchIn, _
        LookAt:=xlSearchType, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    
    ' Store the location of the first find, to make sure we haven't looped
    Set rngFirstFind = rngWorking
    bFoundRequiredHeaders = False
    
    ' If we find the first element of the range, check for the rest
    While ((Not rngWorking Is Nothing) And (Not bFoundRequiredHeaders))
        ' Check horizontally
        bFoundRequiredHeaders = True
        For n = 1 To UBound(sHeaders)
            If (InStr(sHeaders(n), "*") > 0) Then
                nNumCharsToCompare = Len(Before(CStr(sHeaders(n)), "*"))
            Else
                nNumCharsToCompare = Len(sHeaders(n))
            End If
            
            ' If the values we're comparing are numbers, convert them to float so we don't mess with formatting.
            If (IsNumeric(rngWorking.Offset(0, n)) And IsNumeric(sHeaders(n))) Then
                If (CDbl(rngWorking.Offset(0, n)) <> CDbl(sHeaders(n))) Then
                    bFoundRequiredHeaders = False
                    Exit For
                End If
            Else
                If (UCase(Left(rngWorking.Offset(0, n), nNumCharsToCompare)) <> UCase(Left(sHeaders(n), nNumCharsToCompare))) Then
                    bFoundRequiredHeaders = False
                    Exit For
                End If
            End If
        Next
        
        ' If we found the required headers, look for the optional ones too.
        If (bFoundRequiredHeaders) Then
            If (Not IsMissing(sOptional)) Then
                For o = 0 To UBound(sOptional)
                    If (InStr(sHeaders(o), "*") > 0) Then
                        nNumCharsToCompare = Len(Before(CStr(sHeaders(o)), "*"))
                    Else
                        nNumCharsToCompare = Len(sHeaders(o))
                    End If
                    
                    If IsNumeric(rngWorking.Offset(0, n + o)) And IsNumeric(sOptional(o)) Then
                        If (CLng(rngWorking.Offset(0, n + o)) <> CLng(sOptional(o))) Then
                            bFoundRequiredHeaders = False
                            Exit For
                        End If
                    Else
                        If (UCase(Left(rngWorking.Offset(0, n + o), nNumCharsToCompare)) <> UCase(Left(sOptional(o), nNumCharsToCompare))) Then
                            o = 0
                            Exit For
                        End If
                    End If
                Next
            End If
            
            Set rngFinal = rngWorking.Resize(1, n + o)
            GoTo EXIT_GETTABLE
        End If
        
        
        ' Check vertically
        bFoundRequiredHeaders = True
        For n = 1 To UBound(sHeaders)
            If (InStr(sHeaders(n), "*") > 0) Then
                nNumCharsToCompare = Len(Before(CStr(sHeaders(n)), "*"))
            Else
                nNumCharsToCompare = Len(sHeaders(n))
            End If
            
            ' If the values we're comparing are numbers, convert them to float so we don't mess with formatting.
            If (IsNumeric(rngWorking.Offset(n, 0)) And IsNumeric(sHeaders(n))) Then
                If (CLng(rngWorking.Offset(n, 0)) <> CLng(sHeaders(n))) Then
                    bFoundRequiredHeaders = False
                    Exit For
                End If
            Else
                If (UCase(Left(rngWorking.Offset(n, 0), nNumCharsToCompare)) <> UCase(Left(sHeaders(n), nNumCharsToCompare))) Then
                    bFoundRequiredHeaders = False
                    Exit For
                End If
            End If
        Next
        
        ' If we found the required headers, look for the optional ones too.
        If (bFoundRequiredHeaders) Then
            If (Not IsMissing(sOptional)) Then
                For o = 0 To UBound(sOptional)
                    If (InStr(sHeaders(o), "*") > 0) Then
                        nNumCharsToCompare = Len(Before(CStr(sHeaders(o)), "*"))
                    Else
                        nNumCharsToCompare = Len(sHeaders(o))
                    End If
                
                    If IsNumeric(rngWorking.Offset(n + o, 0)) And IsNumeric(sOptional(o)) Then
                        If (CLng(rngWorking.Offset(n + o, 0)) <> CLng(sOptional(o))) Then
                            bFoundRequiredHeaders = False
                            Exit For
                        End If
                    Else
                        If (UCase(Left(rngWorking.Offset(n + o, 0), nNumCharsToCompare)) <> UCase(Left(sOptional(o), nNumCharsToCompare))) Then
                            o = 0
                            Exit For
                        End If
                    End If
                Next
            End If
            Set rngFinal = rngWorking.Resize(n + o, 1)
            GoTo EXIT_GETTABLE
        End If
        
        ' Didn't find a matching row or column, so keep looking
        Set rngWorking = rngSearch.Find(What:=sFirstHeader, After:=rngWorking, LookIn:=xlSearchIn, _
            LookAt:=xlSearchType, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False)
        
        ' We've looped back to the beginning, so quit trying
        If (rngWorking.Address = rngFirstFind.Address) Then
            Set rngWorking = Nothing
        End If
    Wend
    
EXIT_GETTABLE:
    Set GetTable = rngFinal
End Function


'
' Returns the row number of the specified string within the specified worksheet.
' If * is specified as the string to find, the function returns the next open row.
'
' [in]  wks             The worksheet to search
' [in]  strToFind       The string to search for
'
'
Function FindRow(wks As Worksheet, strToFind As Variant)
    On Error Resume Next
    FindRow = wks.Cells.Find(What:=strToFind, _
                                After:=wks.Range("A1"), _
                                LookAt:=xlPart, _
                                LookIn:=xlFormulas, _
                                SearchOrder:=xlByRows, _
                                SearchDirection:=xlPrevious, _
                                MatchCase:=False).Row
    On Error GoTo 0
End Function


'
' Returns the column number of the specified string within the the specified worksheet.
' If * is specified as the string to find, the function returns the next open column.
'
' [in]  wks             The worksheet to search
' [in]  strToFind       The string to search for
'
' DEPRECATED:  Use GetColumnByHeading instead(?)
'
Function FindColumn(wks As Worksheet, strToFind As Variant)
    Dim rngCell As Range
    
    On Error Resume Next
    Set rngCell = wks.Cells.Find(What:=strToFind, _
                                    After:=wks.Range("A1"), _
                                    LookAt:=xlPart, _
                                    LookIn:=xlFormulas, _
                                    SearchOrder:=xlByColumns, _
                                    SearchDirection:=xlPrevious, _
                                    MatchCase:=False)
    FindColumn = rngCell.Column
    On Error GoTo 0
End Function


'
' Finds the last populated cell in the specified column
'
' [in]  rStartCell  A cell in the column which is to be searched
'
' UPDATED 7/20/2013: Made last row based on count property rather than fixed.
'
Function FindLastRow(rStartCell As Range) As Range
    Dim wks As Worksheet
    
    Set wks = rStartCell.Worksheet
    Set FindLastRow = wks.Cells(wks.Rows.Count, rStartCell.Column).End(xlUp)
End Function


'
' Finds the last populated cell in the specified row
'
' [in]  rStartCell  A cell in the row which is to be searched
'
' UPDATED 7/20/2013: Made last column based on count property rather than fixed.
'
Function FindLastCol(rStartCell As Range) As Range
    Dim wks As Worksheet
    
    Set wks = rStartCell.Worksheet
    Set FindLastCol = wks.Cells(rStartCell.Row, wks.Columns.Count).End(xlToLeft)
End Function


'
' Finds the specified string in the specified column number.  If found the function
' returns the row number where the search string was found.  If not found the function
' returns 0.
'
' [in]  szString    The string for which to search
' [in]  nColIndex   The column number in which to search for the string
'
Function FindValueInColumn(rngSearchRange As Range, szStringToFind As String, nColIndex As Long, Optional nAfter As Long = 1) As Long
    Dim nRetVal As Long
    Dim rngCell As Range
    Dim rngAfter As Range

    nRetVal = 0
    Set rngAfter = rngSearchRange.Cells(nAfter, nColIndex)
    Set rngCell = rngSearchRange.Columns(nColIndex).Find(What:=szStringToFind, _
                                                            LookIn:=xlFormulas, _
                                                            LookAt:=xlPart, _
                                                            SearchOrder:=xlByRows, _
                                                            SearchDirection:=xlNext, _
                                                            MatchCase:=False, _
                                                            SearchFormat:=False, _
                                                            After:=rngAfter)
        
    If (Not IsNothing(rngCell)) Then
        nRetVal = rngCell.Row
    End If
    
    FindValueInColumn = nRetVal
End Function


'
' Returns the column number which contains the specified heading
'
' [in]  ColHeading  The heading we're looking for
'
'   - Expects the column headings to be in the first row
'
' DEPRECATED:  Use GetColumnByHeading instead(?)
Function ColumnFromHeading(ColHeading As Variant)
    Dim rgHeader As Range
    Dim i As Integer
    
    ColumnFromHeading = -1
    
    Set rgHeader = Range("1:1")
    For i = 1 To ActiveSheet.UsedRange.Columns.Count
        If rgHeader.Cells(1, i).Value = ColHeading Then
            ColumnFromHeading = i
            Exit For
        End If
    Next i

End Function


'
' Returns the heading of the specified column in the specified worksheet
'
' [in]  aSheet  The worksheet in which we are to look for the heading
' [in]  Col     The column number of which to return the heading
'
'   - Expects the column headings to be in the first row
'
Function HeadingFromColumn(aSheet As Worksheet, col As Integer)
    HeadingFromColumn = aSheet.Range("1:1").Cells(1, col)
End Function


'
' Searches the specified row for the "heading" string, and returns the column number in
' which it finds it.  Returns 0 if the string is not in the row.
'
' Modified to search within a range, rather than a worksheet (1/6/08)
'
Function GetColumnByHeading(szHeading As String, Optional rngData As Range, Optional bAppend As Boolean = False, Optional nInsertAt = 0) As Long
    Dim nColumn As Long
    Dim rngEnd As Range
    Dim nRow As Long
    
    If IsNothing(rngData) Then
        Set rngData = ActiveSheet.Cells
    End If
    
    nColumn = 0
    On Error Resume Next
    nColumn = Application.WorksheetFunction.match(szHeading, rngData.Rows(1), 0)
    On Error GoTo 0
    
    If ((nColumn = 0) And (bAppend)) Then
        nRow = rngData.Rows(1).Row
        If (nInsertAt = 0) Then
            If rngData.Cells(1, 1) = "" Then
                Set rngEnd = rngData.Cells(1, rngData.Columns.Count).End(xlToLeft)
            Else
                Set rngEnd = rngData.Cells(1, rngData.Columns.Count).End(xlToLeft).Offset(0, 1)
            End If
        Else
            rngData.Cells(nRow, nInsertAt).EntireColumn.Insert (xlToRight)
            Set rngEnd = rngData.Cells(1, nInsertAt)
        End If
        rngEnd.Cells(1, 1).Value = szHeading
        nColumn = rngEnd.Column
    End If
    
    GetColumnByHeading = nColumn
End Function


'
' Looks up an entry in the first column within a range, and returns
' the value of the entry in the column specified by the column name
'
'
Function LookupByHeading(vValue As Variant, szHeading As String, rngData As Range, Optional nRow As Long)
    Dim nColumn As Long
    Dim vRetVal As Variant
    
    nColumn = GetColumnByHeading(szHeading, rngData)
    On Error Resume Next
    vRetVal = Application.WorksheetFunction.VLookup(vValue, rngData, nColumn, False)
    On Error GoTo 0
    
    LookupByHeading = vRetVal
End Function


'
' Finds a string in a list of headings and returns the value from the correct column of the data row.
'
' [in]  strLabel            The column name to search for
' [in]  rngHeadings         The range containing the headings for a table
' [in]  rngDataRow          The range containing the data from which a value is returned
Function FindValueFromHeading(strColLabel As String, rngHeadings As Range, rngDataRow As Range) As Variant
    Dim nResultCol As Long
    
    On Error Resume Next
    nResultCol = Application.WorksheetFunction.match(strColLabel, rngHeadings, 0)
    FindValueFromHeading = rngDataRow.Cells(, nResultCol).Value
    On Error GoTo 0

End Function



' ************************* End of Lookup and Search functions *************************************************




'
' Sorts a range on a specified column name.
'
' [in]  strColumnName   The name of the column to sort on
' [in]  rng             The range to be sorted.  If not specified, all cells in the current worksheet are sorted.
'
Sub SortOnColumnName(strColumnName As String, Optional rng As Range, Optional etSortOrder As XlSortOrder = xlAscending, _
                        Optional strColumnName2 As String, Optional etSortOrder2 As XlSortOrder = xlAscending, _
                        Optional strColumnName3 As String, Optional etSortOrder3 As XlSortOrder = xlAscending)

    Dim rngSortKey As Range
    Dim nSortCol As Integer
    Dim wks As Worksheet
    
    If IsNothing(rng) Then
        Set rng = ActiveSheet.Cells
    End If
    
    Set wks = rng.Worksheet
    wks.Sort.SortFields.Clear
    
    nSortCol = GetColumnByHeading(strColumnName, wks.Rows(1))
    wks.Sort.SortFields.Add Key:=rng.Cells(2, nSortCol).Resize(wks.Rows.Count - 1, 1) _
        , SortOn:=xlSortOnValues, Order:=etSortOrder, DataOption:=xlSortNormal
    
    If (strColumnName2 <> "") Then
        nSortCol = GetColumnByHeading(strColumnName2, wks.Rows(1))
        wks.Sort.SortFields.Add Key:=rng.Cells(2, nSortCol).Resize(wks.Rows.Count - 1, 1) _
            , SortOn:=xlSortOnValues, Order:=etSortOrder2, DataOption:=xlSortNormal
    End If
    
    If (strColumnName3 <> "") Then
        nSortCol = GetColumnByHeading(strColumnName3, wks.Rows(1))
        wks.Sort.SortFields.Add Key:=rng.Cells(2, nSortCol).Resize(wks.Rows.Count - 1, 1) _
            , SortOn:=xlSortOnValues, Order:=etSortOrder3, DataOption:=xlSortNormal
    End If
    
    
    With wks.Sort
        .SetRange rng
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


'
' Replaces blank cells with 0 for all cells in the specified range.
'
' [in]  rngFillRange    The range in which to replace blanks with zeros.
'
Sub FillZeros(rngFillRange As Range)
    Dim rCell As Range
    
    For Each rCell In rngFillRange
        If (IsEmpty(rCell)) Then
            rCell.Value = 0
        End If
    Next
End Sub


' BOOKMARK: END OF REVIEW 7/20/2013
'
' BUGBUG: Is this necessary?  Can't one just use 'is Nothing'?
'
Function IsNothing(obj) As Boolean
Dim t As Variant

    On Error Resume Next
    
    Err.Clear
    t = obj.Value
    If (Err <> 0) Then
        IsNothing = True
        Err.Clear
    Else
        IsNothing = False
    End If
End Function


''****************
'   Date Tools
'.****************

'
' Adjusts the dates in a range, by the specified number of hours.
' This function is mostly for adjusting Product Studio resolved and closed dates which
' are reported in GMT.
'
' [in]  rngRange            The range in which all cells containing dates will be adjusted
' [in]  fHours              The # of hours by which to adjust dates (negative is backward in time).
Sub AdjustDates(rngRange As Range, fHours As Double)
    Dim cell As Range
    Dim i, j As Integer
    
    For i = 1 To rngRange.Rows.Count
        For j = 1 To rngRange.Columns.Count
            Set cell = rngRange.Cells(i, j)
            If IsDate(cell) Then
                cell.Value = cell.Value + (fHours / 24)
            End If
        Next
    Next
End Sub


'
' Truncates the hours and minutes from a date in a range, by the specified number of hours.
' This function is mostly for adjusting Product Studio resolved and closed dates which
' include hours, minutes and seconds in the date.
'
' [in]  rngRange            The range in which all cells containing dates will be adjusted
Sub TruncateDates(rngRange As Range)
    Dim cell As Range
    Dim i, j As Integer
    
    For i = 1 To rngRange.Rows.Count
        For j = 1 To rngRange.Columns.Count
            Set cell = rngRange.Cells(i, j)
            If IsDate(cell) Then
                cell.Value = Int(cell.Value)
            End If
        Next
    Next
End Sub


'
''****************
'    File Tools
'.****************
'
' Some of the functions in File Tools require the Microsoft Scripting Runtime.
' To add it, select Tools|References|Microsoft Scripting Runtime
'

' For more documentation on using the FileSystemObject, see http://msdn2.microsoft.com/en-us/library/6kxy1a51.aspx
'


'
' Replaces the source path with a destination path and returns the new absolute filename
'
Function ReplacePath(strFullName As String, strSrcPath As String, strNewPath As String) As String
    Dim strRetVal As String
    
    strRetVal = Replace(strFullName, strSrcPath, strNewPath, 1, 1, vbTextCompare)
    ReplacePath = strRetVal
End Function


'
' Counts the number of lines in a text file
'
' UPDATED 7/21/2013: Defined constant for ForAppending, and made code more readable
'
Function NumberOfLines(sFile As String) As Double
    Dim f As Object
    Dim objScripting As Object
    
    Const ForAppending = 8
    
    
    Set objScripting = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
        Set f = objScripting.OpenTextFile(sFile, ForAppending)
        If Err.Number > 0 Then
            NumberOfLines = 0
        Else
            NumberOfLines = f.line
            f.Close
            Set f = Nothing
        End If
    On Error GoTo 0
End Function


' Creates the specified folder path, one node at a time.
' The fso.CreateFolder command fails if most of the path does not already exist.
'
' Returns:
'   True if the folder path was created or already existed.
'   False if the folder path could not be created.
'
Function CreateFolderPathEx(strPath As String) As Boolean
    Dim nDepth As Integer
    Dim i As Integer
    Dim strNode As String
    Dim strParent As String
    Dim fso As Object 'FileSystemObject
    Dim bRetVal As Boolean
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    strNode = strPath
    strParent = fso.GetParentFolderName(strNode)
    
    If (Not (fso.FolderExists(strParent))) Then
        bRetVal = CreateFolderPathEx(strParent)
    End If
    
    If (Not (fso.FolderExists(strNode))) Then
        fso.CreateFolder (strNode)
    End If
    
    CreateFolderPathEx = True
End Function


'
' Deletes the list of files in the specified range.
' If bDelDirs is false, directories are skipped from deletion
' If bRemoveEmpty is true, empty directories are deleted.
'
Sub DeleteFileList(rngFileList As Range, Optional bDelDirs As Boolean = False)
    Dim cell As Range
    Dim fso As FileSystemObject
    Dim strFile As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    For Each cell In rngFileList
        strFile = cell.Value
        If fso.FolderExists(strFile) Then
            If bDelDirs Then
                Debug.Print ("Deleting folder " + strFile)
                fso.DeleteFolder strFile
            Else
                Debug.Print ("Not deleting folder " + strFile)
            End If
        End If
        
        If fso.FileExists(strFile) Then
            Debug.Print ("Deleting " + strFile)
            fso.DeleteFile strFile
        End If
    Next
End Sub


Sub DeleteEmptyFolders(strSubFolder As String)
    Dim oFiles As FoundFiles
    Dim f As File
    Dim fso As FileSystemObject
    Dim cell As Range
    Dim i As Long
    Dim folderRoot As Folder
    Dim folderNext As Folder
        
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folderRoot = fso.GetFolder(strSubFolder)
    
    For Each folderNext In folderRoot.SubFolders
        DeleteEmptyFolders (folderNext.path)
    Next
    
    If ((folderRoot.SubFolders.Count = 0) And (folderRoot.Files.Count = 0)) Then
        Debug.Print ("Deleting " + folderRoot)
        fso.DeleteFolder folderRoot, True
    Else
        Debug.Print ("Didn't delete " + folderRoot)
    End If
        
End Sub


'
' The Application.FileSearch functionality appears to have been deprecated
' in Excel 2007 and beyond.  Use Dir and fso instead.
'
'Function GetFileList(strPath As String, strFileSpec As String, _
'                    Optional bSubFolders As Boolean = True, _
'                    Optional rngDestRange As Range) As FoundFiles
'    Dim fs As FileSearch
'    Dim i As Long
'
''    Set fs = Application.FileSearch
'    dir(
'
'    fs.NewSearch
'    fs.LookIn = strPath
'    fs.SearchSubFolders = bSubFolders
'    fs.Filename = strFileSpec
'    fs.Execute
'
''    msoFileTypeExcelWorkbooks
'
'    Set GetFileList = fs.FoundFiles
'
'    If Not rngDestRange Is Nothing Then
'        For i = 1 To fs.FoundFiles.Count
'            rngDestRange.Cells(i, 1) = fs.FoundFiles(i)
'        Next
'    End If
'End Function
'
'
'Function GetDirectoryList(strPath As String, strNameSpec As String, _
'                    Optional bSubFolders As Boolean = True, _
'                    Optional rngDestRange As Range) As FoundFiles
'    Dim fs As FileSearch
'    Dim i As Long
'
'    Set fs = Application.FileSearch
'
'    fs.NewSearch
'    fs.LookIn = strPath
'    fs.Filename = strFileSpec
'    fs.Execute
'
'    Set GetFileList = fs.FoundFiles
'
'    If Not IsEmpty(rngDestRange) Then
'        For i = 1 To fs.FoundFiles.Count
'            rngDestRange.Cells(i, 1) = fs.FoundFiles(i)
'        Next
'    End If
'End Function


'
' Returns the path to the OS temp directory for storing
' temporary files.  This path is different per user.
'
' Use GetTempFile to generate an empty temp file.
'
Function GetTempPath() As String
    Dim fso As Scripting.FileSystemObject
    Dim strPath As String
    Dim nBufferSize As Long
    Dim nLen As Long
    
    nBufferSize = 255
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    strPath = Space(nBufferSize)
    nLen = OS_GetTempPath(nBufferSize, strPath)
    
    GetTempPath = Left(strPath, nLen)
End Function


'
' Creates an empty temp file in the OS temp directory
' for the current user, and returns the filename
'
Function GetTempFile() As String
    Dim fso As Scripting.FileSystemObject
    Dim strPath As String
    Dim strFile As String
    Dim nBufferSize As Long
    Dim nLen As Long
    Dim strPrefix As String
    
    nBufferSize = 255
    
    strPath = GetTempPath()

    ' Get a uniquely assigned random file
    strPrefix = "tmp"
    strFile = Space(nBufferSize)                                ' Create a buffer to receive the filename
    nLen = OS_GetTempFileName(strPath, strPrefix, 0, strFile)   ' Call the Win32 function to create a temp file
    strFile = Left(strFile, nLen)                               ' Trim the null char from the end of the filename
    
    GetTempFile = strFile
End Function


Function FileExists(filename As String, Optional attributes As Long)
    Dim retVal As Boolean
    Dim path As String
    
    retVal = False
    
    ' Suppress errors from the dir function in case
    ' the input string is not a path (which errors out with File Not Found)
    On Error Resume Next
    path = Dir$(filename, attributes)
    On Error GoTo 0
    
    If Len(path) > 0 Then
        retVal = True
    End If
    
    FileExists = retVal
End Function

Sub WriteStringToFile(filename As String, line As String, Optional append As Boolean = False)
    Dim fileNum As Integer
 
    ' Open file for output
    fileNum = FreeFile()
    
    If (append) Then
        Open filename For Append As fileNum
    Else
        Open filename For Output As fileNum
    End If
 
    Write #fileNum, line
    Close #fileNum
 End Sub

Function ReadFileToString(filename As String) As String
    Dim fileNum As Integer
    Dim oneLine As String
    Dim wholeFile As String

    wholeFile = ""
    oneLine = ""
    
    ' Check if the file exists
    If Len(Dir$(filename)) = 0 Then
        GoTo EXIT_READFILETOSTRING
    End If

    fileNum = FreeFile()
    Open filename For Input As fileNum

    ' Concatenate all lines in the file to one string
    ' Line Input removes the CR/LF, so add it back.
    Do While Not EOF(fileNum)
        Line Input #fileNum, oneLine
        Debug.Print oneLine
        wholeFile = wholeFile + oneLine + Chr(13) + Chr(10)
    Loop

    ' Trim the extra CR/LF
    wholeFile = Left(wholeFile, Len(wholeFile) - 2)

    ' Close the file
    Close fileNum
    ReadFileToString = wholeFile

EXIT_READFILETOSTRING:
End Function


''************
'  Path Tools
'.************

'
' Returns the directory path of
' the specified absolute file name.
'
' That is, the absolute path minus the
' base filename and extension.
'
' example:
'   strFilename = c:\MyDirectory\test.txt
'   returns: c:\MyDirectory
'

Function GetPath(strFilename)
    Dim strPathTokens() As String
    Dim n As Integer
    Dim strPath As String
    
    strPathTokens = Split(strFilename, "\")
   
    For n = 0 To UBound(strPathTokens) - 1
      strPath = strPath + "\" + strPathTokens(n)
    Next
    
    GetPath = Mid(strPath, 2)
End Function


'
' Returns the filename and extension of
' the specified absolute file name.
'
'
' example:
'   strFilename = c:\MyDirectory\test.txt
'   returns: test.txt
'
Function GetEntryName(strFilename)
    Dim strPathTokens() As String
    Dim n As Integer
    Dim strEntry As String

    strPathTokens = Split(strFilename, "\")
    strEntry = strPathTokens(UBound(strPathTokens))

    GetEntryName = strEntry
End Function


'
' Returns the base filename without extension for
' the specified absolute file name.
'
' example:
'   strFilename = c:\MyDirectory\test.txt
'   returns: test
'
Function GetBaseName(strFilename)
    Dim strPathTokens() As String
    Dim n As Integer
    Dim strEntry As String
    Dim strBase As String

   strEntry = GetEntryName(strFilename)
   strBase = Before(strEntry, ".")

    GetBaseName = strBase
End Function


'
' Returns the base filename without extension for
' the specified absolute file name.
'
' example:
'   strFilename = c:\MyDirectory\test.txt
'   returns: TXT
'
Function GetExtension(strFilename)
    Dim strPathTokens() As String
    Dim n As Integer
    Dim strEntry As String
    Dim strExtension As String
    
    strEntry = GetEntryName(strFilename)
    strExtension = After(strEntry, ".")
    strExtension = UCase(strExtension)

    GetExtension = strExtension
End Function



''***************
'   Set Tools
'.***************

'
' Returns TRUE if the specified item is in the specified set (range)
'
' UPDATED 7/22/2013: Added optional arguments for MatchCase and LookAt
' BUGBUG: Should this use Match instead of Find so that an array can be searched as well?
'
Function IsInSet(varElement As Variant, rngSet As Range, _
                    Optional bMatchCase As Boolean = False, Optional xlSearchType As Integer = xlWhole) As Boolean
    Dim rngCell As Range
    
    Set rngCell = rngSet.Find(What:=varElement, MatchCase:=bMatchCase, LookAt:=xlSearchType)
    
    If rngCell Is Nothing Then
        IsInSet = False
    Else
        IsInSet = True
    End If
End Function


'
' Returns an array of strings which are unique from a range.
' If multiple columns are selected, the text from all columns
' in a row are combined into a single entry.
'
' NOTE: Requires a reference to the Microsoft Scripting Runtime
'
Sub GetUniqueItems(szUnique() As String, rngToSearch As Range, _
                    Optional strDelimiter As String = "_:_")
    Dim cell As Range
    Dim dic As Scripting.Dictionary
    Dim dicItem As Variant
    Dim strUnique() As String
    Dim nNumItems As Long
    Dim strSearchVal As String
    Dim i, j As Long
    Dim nNumColumns As Long
    Dim nNumRows As Long

    If rngToSearch Is Nothing Then
        Set rngToSearch = ActiveCell
    End If

    ' Create dictionay object
    Set dic = New Scripting.Dictionary

    ' Populate dictionary object with unique items
    nNumColumns = rngToSearch.Columns.Count
    nNumRows = rngToSearch.Rows.Count
    strSearchVal = ""
    
    For i = 1 To nNumRows
        Set cell = rngToSearch.Cells(i, 1)
        
        strSearchVal = ""
        For j = 1 To nNumColumns
            If Not IsEmpty(cell) Then
                strSearchVal = strSearchVal + strDelimiter + CStr(cell.Offset(0, j - 1).Value)
            End If
        Next
    
        ' Trim off the leading delimiter
        strSearchVal = After(strSearchVal, strDelimiter)
    
        If (Not dic.exists(strSearchVal)) Then
            dic.Add strSearchVal, CVar(strSearchVal)
        End If
    Next

    ' Pass back the list of unique values
    If Not dic Is Nothing Then
        ReDim szUnique(dic.Count)
        For i = 1 To dic.Count
            szUnique(i) = dic.Items(i - 1)
        Next i
        
        'Clean up objects
        Set dic = Nothing
    End If
End Sub


'
' Determines if a specified cell is within a range
'
Function InRange(Range1 As Range, Range2 As Range) As Boolean
   Dim InterSectRange As Range
   
   Set InterSectRange = Application.Intersect(Range1, Range2)
   InRange = Not InterSectRange Is Nothing
   Set InterSectRange = Nothing
End Function



''****************
'   Math Tools
'.****************


' Returns the incremental to be added to each cell in a range
' to interpolate between the first cell in the range to the last.
'
' BUGBUG: What's the nInterval argument for?  Seems like it should always be 1.
'
Function Interpolate(rngData As Range, nInterval As Integer) As Double
    Dim nIncrement As Double
    Dim nRangesize As Integer
    Dim nRetVal As Double
    
    ' Make sure that the range is either 1 row or 1 column.
    If ((rngData.Rows.Count > 1) And (rngData.Columns.Count > 1)) Then
        nRetVal = 0#
    Else
        nRangesize = Application.WorksheetFunction.max(rngData.Rows.Count, rngData.Columns.Count)               ' Figure out the number of cells in the range
        nIncrement = (rngData(rngData.Rows.Count, rngData.Columns.Count).Value - rngData(1, 1)) / nRangesize    ' Calculate the increment to provide a linear progression
        nRetVal = rngData(1, 1) + (nInterval * nIncrement)                                                      ' Return the next value in the seriesw
    End If
    
    Interpolate = nRetVal
End Function


'
' Provides a shortcut to the MAX worksheet function
'
Function Max_(x1 As Double, x2 As Double) As Double
    Max_ = Application.WorksheetFunction.max(x1, x2)
End Function


'
' Evaluates the first and last values in a column
' and populates the cells in between with linearly interpolated values
'
Sub InterpolateSelectedColumn()
    Dim nCurrentRow As Integer
    Dim nCurrentCol As Integer
    Dim rngCell As Range
    Dim nStartRow, nEndRow As Integer
    Dim nStartCol, nEndCol As Integer
    
    Dim i, j As Integer
    Dim rng As Range
    Dim nNumRows As Integer
    Dim nFirstVal, nLastVal As Integer
    Dim fIncrement As Double
    
    Set rng = Selection
    
    nStartRow = rng.Cells(1, 1).Row
    nEndRow = rng.Cells(rng.Rows.Count).Row
    
    ' Calculate the increment
    nNumRows = nEndRow - nStartRow
    
    nFirstVal = rng.Cells(1, 1).Value
    nLastVal = rng.Cells(rng.Rows.Count, 1).Value
    fIncrement = (nLastVal - nFirstVal) / (nNumRows)
    
    Set rngCell = rng.Cells(1, 1).Offset(1, 0)
    While rngCell.Row < nEndRow
        rngCell.Value = rngCell.Offset(-1, 0) + fIncrement
        Set rngCell = rngCell.Offset(1, 0)
    Wend
End Sub


'
' Accepts two hexadecimal numbers as strings, and returns
' the bitwise and of those numbers as a string.
'
' Example:  HexAnd("FFFF", "F0F0") returns "F0F0"
'
Function HexAnd(hex1 As String, hex2 As String) As String
    Dim bin1 As String
    Dim bin2 As String
    
    bin1 = CLng("&H" & hex1)
    bin2 = CLng("&H" & hex2)
 
    HexAnd = Hex$(bin1 And bin2)
End Function



''*******************
'   Clipboard Tools
'.*******************
'
' NOTE:  These require that you add a reference to the Microsoft Forms 2.0 (FM20.dll) object library
'        which may be in a different directory (try c:\windows\syswow64
'

' Puts the specified object on the clipboard
Public Sub PutOnClipboard(obj As Variant)
    Dim MyDataObj As New DataObject
    MyDataObj.SetText Format(obj)
    MyDataObj.PutInClipboard
End Sub

' Retrieves an object from the clipboard
Public Function GetOffClipboard() As Variant
    Dim MyDataObj As New DataObject
    MyDataObj.GetFromClipboard
    GetOffClipboard = MyDataObj.GetText()
End Function

' Clears the clipboard
Public Sub ClearClipboard()
    Dim MyDataObj As New DataObject
    MyDataObj.SetText ""
    MyDataObj.PutInClipboard
End Sub



''****************
'   String Tools
'.****************


' Takes a semicolon separated string, from the active cell, and dumps the contents
' NOTE:  The contents are dumped immediately below the active cell, stomping previous
'        contents
'
Sub SplitOnSemicolon()
Attribute SplitOnSemicolon.VB_ProcData.VB_Invoke_Func = "w\n14"
    Dim n As Long
    Dim i As Long
    Dim strTokens() As String
    
    n = Tokenize(ActiveCell.Value, ";", strTokens)
    For i = 1 To n
        ActiveCell.Offset(i, 0).Value = strTokens(i)
    Next
End Sub


' Returns the input string with quotation marks added at the beginning and end
Function WrapInQuotes(strString As String) As String
    WrapInQuotes = Chr(34) + CStr(strString) + Chr(34)
End Function


' Returns the input string with all blanks removed
' BUGBUG: Check if there are already functions that also remove tabs and newlines
Function RemoveWhiteSpace(strString As String) As String
    Dim i As Integer
    Dim strReduced As String
    Dim strOneChar As String
    
    i = 1
    While i <= Len(strString)
        If (Mid(strString, i, 1) <> " ") Then
            strReduced = strReduced + Mid(strString, i, 1)
        End If
        i = i + 1
    Wend
    
    RemoveWhiteSpace = strReduced
End Function


'
' Takes as input a string and a substring.  If the substring is found in
' the string, the function returns the remainder of the string after the substring.
' If the string is not found, the function returns an empty string.
'
Function After(szString As String, szSubStr As String) As String
    Dim nIndex As Integer
    Dim szReturnVal As String
    
    nIndex = InStr(1, szString, szSubStr, vbTextCompare)
    
    If (nIndex > 0) Then
        szReturnVal = Mid(szString, nIndex + Len(szSubStr))
    Else
        szReturnVal = vbNullString
    End If
    
    After = szReturnVal
End Function


'
' Takes as input a string and a substring.  If the substring is found in
' the string, the function returns the string before the first character of the substring.
' If the string is not found, the function returns szString.
'
Function Before(szString As String, szSubStr As String) As String
    Dim nIndex As Integer
    Dim szReturnVal As String
    
    nIndex = InStr(1, szString, szSubStr, vbTextCompare)
    If (nIndex = 0) Then
        szReturnVal = szString
    Else
        szReturnVal = Left(szString, nIndex - 1)
    End If
    
    Before = szReturnVal
End Function


'
' Returns the string between two other strings in a main string
'
' Modifed 12/9/15 to return the whole string if the start or end tokens
' are not found, and partialMatch is set to false.
'
Public Function Between(szString As String, szSubStrStart As String, szSubStrEnd As String, Optional partialMatch As Boolean = True) As String
Attribute Between.VB_Description = "Returns the substring of string between start and end"
    Dim nStartIndex As Long
    Dim nEndIndex As Long
    Dim szReturnVal As String
    
    nStartIndex = InStr(1, szString, szSubStrStart, vbTextCompare)
    nEndIndex = InStr(Application.max(1, nStartIndex), szString, szSubStrEnd, vbTextCompare)
    
    ' Handle different cases of if the start and end strings exist
    If (((nStartIndex = 0) Or (nEndIndex = 0)) And (Not partialMatch)) Then
        szReturnVal = szString
    Else
        If (nStartIndex > 0) And (nEndIndex > 0) Then
            szReturnVal = SuperTrim(Before(After(szString, szSubStrStart), szSubStrEnd))
        ElseIf (nStartIndex = 0) And (nEndIndex > 0) Then
            szReturnVal = SuperTrim(Before(szString, szSubStrEnd))
        ElseIf (nStartIndex > 0) And (nEndIndex = 0) Then
            szReturnVal = SuperTrim(After(szString, szSubStrStart))
        Else
            szReturnVal = ""
        End If
    End If
    
    Between = szReturnVal
End Function

Sub TestBetween()
    Debug.Print (Between("This is a test!", "This ", " a"))
    Debug.Print (Between("This is a test!", "Apple", "test"))
    Debug.Print (Between("This is a test!", "is a ", "Boat"))
End Sub


' Includes a null object in the elements to
' convert to string.  Normally this returns
' a run-time error 94.
Function SuperCStr(val As Variant) As String
    If (IsNull(val)) Then
        SuperCStr = ""
    Else
        SuperCStr = CStr(val)
    End If
End Function


Function RemoveChars(s As String, chars As String)
    Dim n As Long
    Dim t As String
    
    t = s
    
    For n = 1 To Len(chars)
        t = Replace(t, Mid(chars, n, 1), "")
    Next
    RemoveChars = t
End Function


Function ReplaceChars(s As String, sourcechars As String, replchars As String)
    Dim n As Long
    Dim t As String
    
    If (Len(sourcechars) <> Len(replchars)) Then
        GoTo EXIT_REPLACECHARS
    End If
    
    t = s
    
    For n = 1 To Len(sourcechars)
        t = Replace(t, Mid(sourcechars, n, 1), Mid(replchars, n, 1))
    Next

EXIT_REPLACECHARS:
    
    ReplaceChars = t
End Function


Sub TestReplaceChars()
    Dim s As String
    Dim t As String
    
    s = "ABCD EFG"
    t = ReplaceChars(s, " A", "!@")
End Sub


Sub DumpStringToASCII(s As String)
    Dim n As Long
    Dim nCode As Integer
    Dim sChar As String
    
    For n = 1 To Len(s)
        sChar = Mid(s, n, 1)
        nCode = Asc(sChar)
        Debug.Print (sChar + " : " + CStr(nCode))
    Next
End Sub


Function GetLetterFromIndex(nIndex As Long) As String
    If (nIndex < 1) Or (nIndex > 26) Then
        GetLetterFromIndex = "_"
    Else
        GetLetterFromIndex = Mid("ABCDEFGHIJKLMNOPQRSTUVWXYZ", nIndex, 1)
    End If
End Function


' Accepts a string and a token delimiter.  The function searches for
' the first instance of the token in the string, and returns the value of the string
' before the token.  The argument string is modified to contain the value after the
' token, so the argument string can be passed in multiple times to parse values.
'
Function StrTok(ByRef szString As String, szToken As String) As String
Dim szRetVal As String

szRetVal = Before(szString, szToken)

szString = After(szString, szToken)
StrTok = szRetVal
End Function


' Removes all non-numeric characters from a string, and
' returns only the complete number (as a string)
Function ExtractNumberFromString(str As String) As String
    Dim n As Long
    Dim sRetVal As String
    Dim nAscii As Integer
    
    For n = 1 To Len(str)
        nAscii = Asc(Mid(str, n, 1))
        If ((nAscii >= 48) And (nAscii <= 57)) Then
            sRetVal = sRetVal + Mid(str, n, 1)
        End If
    Next

    ExtractNumberFromString = sRetVal
End Function


'
' Splits a string into separate items divided on a delimiter.
' The function returns the number of items, and populates the Tokens
' array with the items.
'
' DEPRECATED:  Use Split function instead.
'
Function Tokenize(szString As String, szDelim As String, ByRef szStrTokens() As String) As Long
    Dim nCount As Long
    Dim szCopy As String
    
    ReDim szStrTokens(1)
    
    nCount = 1
    szCopy = szString
    
    If (Len(szCopy) = 0) Then
        nCount = 0
    Else
        szStrTokens(nCount) = StrTok(szCopy, szDelim)
        
        ' while we're not at the end of the line
        ' keep parsing for tokens
        While (szCopy <> "")
            nCount = nCount + 1
            ReDim Preserve szStrTokens(nCount)
            szStrTokens(nCount) = StrTok(szCopy, szDelim)
        Wend
    End If
    Tokenize = nCount
End Function


'
' Same as Tokenize function, but returns a collection rather than
' an array of strings.
'
Function TokenizeToCollection(szString As String, szDelim As String) As Collection
    Dim nCount As Long
    Dim szCopy As String
    Dim colTokens As New Collection
    Dim strToken As String
    
    nCount = 1
    szCopy = szString
    
    If (Len(szCopy) = 0) Then
        nCount = 0
    Else
        strToken = StrTok(szCopy, szDelim)
        colTokens.Add strToken
        While (szCopy <> "")
            strToken = StrTok(szCopy, szDelim)
            colTokens.Add strToken
        Wend
    End If
    Set TokenizeToCollection = colTokens
End Function


'
' Replaces an element in a collection with a new element
'
Function ReplaceCollectionElement(col As Collection, nIndex As Long, varValue As Variant, Optional strKey As String)
    col.Remove (nIndex)
    
    Select Case nIndex
        Case 0 To 1:
            col.Add varValue, strKey
        Case Else
            col.Add varValue, strKey, , nIndex - 1
    End Select
End Function


'
' Returns the index number of the specified item in a collection.
' For example, if a collection contains 5, 10, 15, and 20,  this
' function would return 3 when called with strValue = "15"
'
Function GetCollectionValueIndex(col As Collection, strValue As String) As Long
    Dim nRetVal As Long
    Dim n As Long
    Dim bFoundValue As Boolean
    
    bFoundValue = False
    
    n = 1
    While ((Not bFoundValue) And (n <= col.Count))
        If (col.Item(n) = strValue) Then
            bFoundValue = True
        End If
        
        n = n + 1
    Wend
    
    If (Not bFoundValue) Then
        n = 0
    Else
        n = n - 1
    End If
    
    GetCollectionValueIndex = n
End Function

'
' Counts the number of instances of a delimiter in a string
' NOTE: The delimiter can be only 1 character.
'
Function NumInstance(szString As String, szDelim As String) As Long
    Dim nCount As Long
    Dim strTokens() As String
    Dim n As Long
    
    If (Len(szString) = 0) Then
        nCount = 0
    Else
        For n = 1 To Len(szString)
            If (Mid(szString, n, 1) = szDelim) Then
                nCount = nCount + 1
            End If
        Next
    End If
     
    NumInstance = nCount
End Function


'
' Adds single or double quotes to the beginning and end of the specified string
'
' [in] szString     The string to be wrapped
' [in] nQuoteType   ucSingle or ucDouble [default]
'
Function QuoteString(szString As String, Optional lQuoteType As etQuoteType = enDoubleQuote) As String
    If lQuoteType = enDoubleQuote Then
        QuoteString = Chr(34) + szString + Chr(34)
    Else
        QuoteString = Chr(39) + szString + Chr(39)
    End If
End Function


'
' Accepts an array and sorts it in place
'
Sub BubbleSort(ByRef TempArray As Variant)
    Dim temp As Variant
    Dim i As Integer
    Dim NoExchanges As Integer

    ' Loop until no more "exchanges" are made.
    Do
        NoExchanges = True

        ' Loop through each element in the array.
        For i = LBound(TempArray) To UBound(TempArray) - 1

            ' If the element is greater than the element
            ' following it, exchange the two elements.
            If TempArray(i) > TempArray(i + 1) Then
                NoExchanges = False
                temp = TempArray(i)
                TempArray(i) = TempArray(i + 1)
                TempArray(i + 1) = temp
                                       End If
        Next i
    Loop While Not (NoExchanges)
End Sub


'
' Splits a string containing CRLF pairs into an array of separate strings.
' Returns the number of lines
'
Function StringToLines(str As String, strLines() As String) As Long
    Dim n As Long
    Dim strTemp As String
    Dim nPos As Long
    
    ReDim strLines(1)
    
    strTemp = str
    n = 0
    nPos = InStr(strTemp, Chr(13) + Chr(10))
    While (nPos > 0)
        n = n + 1
        ReDim Preserve strLines(n)
        strLines(n) = Left(strTemp, InStr(strTemp, Chr(13) + Chr(10)))
        strTemp = Mid(strTemp, nPos + 1)
        nPos = InStr(strTemp, Chr(13) + Chr(10))
    Wend
    
    n = n + 1
    ReDim Preserve strLines(n)
    
    strLines(n) = strTemp
    
    StringToLines = n
End Function


'
' Removes CRLF codes from a string
'
Function SuperTrim(strString As String) As String
    Dim strReturn As String
    Dim n As Long
    Dim strChar As String
    Dim nCode As Integer
    Dim bLoopAtBeginning As Boolean
    Dim bLoopAtEnd As Boolean
    
    strReturn = strString
    
    If (Len(strString) = 0) Then
        strReturn = ""
        GoTo EXIT_SUPERTRIM
    End If
    
    nCode = Asc(Mid(strReturn, 1, 1))

    While (nCode <= 32) Or (nCode >= 127)
        strReturn = Right(strReturn, Len(strReturn) - 1)
        nCode = Asc(Mid(strReturn, 1, 1))
    Wend
        
    nCode = Asc(Mid(strReturn, Len(strReturn), 1))
    While (nCode <= 32) Or (nCode >= 127)
        strReturn = Left(strReturn, Len(strReturn) - 1)
        nCode = Asc(Mid(strReturn, Len(strReturn), 1))
    Wend

EXIT_SUPERTRIM:
    SuperTrim = strReturn
End Function


Function FixupCRLF(strSource As String)
    Dim nIndex As Integer
    Dim nPrevChar As String
    Dim strRetString As String
    Dim strCRLF As String
    
    strCRLF = Chr(13) + Chr(10)
    
    strRetString = Replace(strSource, strCRLF, Chr(4))
    strRetString = Replace(strRetString, Chr(13), Chr(4))
    strRetString = Replace(strRetString, Chr(10), Chr(4))
    strRetString = Replace(strRetString, Chr(4), strCRLF)
    
    FixupCRLF = strRetString
End Function


Function MakeDebugString(strSource As String, Optional strName As String = "Debug") As String
    Dim strReturn As String
    Dim nIndex As Integer
    
    strReturn = strName + " (" + CStr(Len(strSource)) + "): "
    For nIndex = 1 To Len(strSource)
        strReturn = strReturn + " " + CStr(Asc(Mid(strSource, nIndex, 1)))
    Next
    
    MakeDebugString = strReturn
End Function


Function IsNameValuePair(strString As String, strDelimiter As String)
    IsNameValuePair = InStr(strString, strDelimiter) > 0
End Function


' ********************
'   Formatting Tools
' ********************


Sub FormatColumnWidths(rng As Range, Optional columnwidths As Variant = Nothing)
    Dim n As Long
    Dim col As Range
    
    If (columnwidths Is Nothing) Then
        For Each col In rng.Columns
            col.EntireColumn.AutoFit
        Next
    Else
        For n = 0 To UBound(columnwidths)
            If (columnwidths(n) = 0) Then
                rng.Columns(n + 1).EntireColumn.AutoFit
            Else
                rng.Columns(n + 1).ColumnWidth = columnwidths(n)
            End If
        Next
    End If
End Sub


' Formats the row height of a merged cell with wrapped text
' to fit all of the text within the row.
'
' That is, if the row height is not tall enough to accomodate
' the content of the cell, the text will appear clipped.
' This subroutine adjusts the row height.
Sub AutoFitMergedCellRowHeight(rngCells As Range)
    Dim nCurRowHeight As Single
    Dim nNewRowHeight As Single
    Dim nMergedCellWidth As Single
    Dim cellCurrent As Range
    Dim rngMerged As Range
    Dim nFirstCellWidth As Single
    
    ' Confirm the specified range holds exactly 1 row of
    ' merged cells, and that word wrap is on.
    If rngCells.MergeCells Then
        Set rngMerged = rngCells(1, 1).MergeArea
        If (rngMerged.Rows.Count = 1) And (rngMerged.WrapText = True) Then
            nCurRowHeight = rngMerged.RowHeight                 ' Store the current row height for comparison
            nFirstCellWidth = rngCells(1, 1).ColumnWidth        ' Store the width of the first column.  We're going to change it
            nMergedCellWidth = 0
            For Each cellCurrent In rngMerged                   ' Tally the total width of the merged area
                nMergedCellWidth = nMergedCellWidth + cellCurrent.ColumnWidth
            Next
            
            rngCells.MergeCells = False                         ' Unmerge the cells
            rngCells.Cells(1).ColumnWidth = nMergedCellWidth    ' Set the first column width to the total of the merged cell width
            rngCells.EntireRow.AutoFit                          ' Ask Excel to figure out the correct row height
            nNewRowHeight = rngCells.RowHeight
            rngCells.Cells(1).ColumnWidth = nFirstCellWidth
            rngMerged.MergeCells = True                         ' Re-merge the cells and set the correct row height
            rngCells.RowHeight = IIf(nCurRowHeight > nNewRowHeight, nCurRowHeight, nNewRowHeight)
        End If
    End If
End Sub


'
' Formats the the selected range using alternating background characteristics from two
' individual named cells "Format1", and "Format2"
'
Sub FormatRangeAsTable()
    Dim rng As Range
    Dim rngTemplate As Range
    
    For Each rng In Selection.Rows
        If rng.Row Mod 2 = 0 Then
            Set rngTemplate = Range("Format2")
            rng.Interior.Pattern = rngTemplate.Interior.Pattern
            rng.Interior.PatternColorIndex = rngTemplate.Interior.PatternColorIndex
            rng.Interior.ThemeColor = rngTemplate.Interior.ThemeColor
            rng.Interior.TintAndShade = rngTemplate.Interior.TintAndShade
        Else
            Set rngTemplate = Range("Format1")
            rng.Interior.Pattern = rngTemplate.Interior.Pattern
            rng.Interior.PatternColorIndex = rngTemplate.Interior.PatternColorIndex '
            rng.Interior.ThemeColor = rngTemplate.Interior.ThemeColor
            rng.Interior.TintAndShade = rngTemplate.Interior.TintAndShade
        End If
    Next
End Sub


'
' Groups cell rows by the level they've been indented
' for use in expand/contract
'
Function GroupAtIndent(rngStartCell As Range) As Integer
    Dim rngGroup As Range
    Dim n As Integer
    Dim nIndent As Integer
    
    n = 0
    nIndent = rngStartCell.IndentLevel
    While ((rngStartCell.Offset(n, 0).IndentLevel >= nIndent) And (rngStartCell.Offset(n, 0) <> ""))
        n = n + 1
        If (rngStartCell.Offset(n, 0).IndentLevel > nIndent) Then
            n = n + GroupAtIndent(rngStartCell.Offset(n, 0))
        End If
    Wend
    
    If (nIndent > 0) Then
        rngStartCell.Resize(n, 1).Rows.Group
    End If
    
    GroupAtIndent = n
End Function


Sub TestShow()
    Dim arrColumns() As Variant
    Dim n As Long
        
    
    arrColumns = Array("Blue", "Red")

    n = ShowNamedColumns(arrColumns)
End Sub

Sub TestHide()
    Dim arrColumns() As Variant
    Dim n As Long
        
    
    arrColumns = Array("Red")

    n = HideNamedColumns(arrColumns)
End Sub


'
' Makes the specified columns visible, if they exist.  Columns are specified by name.
' If the optional argument is true, all other columns are hidden.
' The function returns the number of columns made visible.
'
Function ShowNamedColumns(arrColumnNames() As Variant, Optional bHideOtherColumns = True) As Long
    Dim i As Long
    Dim nRetVal As Long
    Dim strColName As Variant
    
    If (bHideOtherColumns) Then
        Columns.Hidden = True
    End If
    
    For Each strColName In arrColumnNames
        i = GetColumnByHeading(CStr(strColName), Rows(1), False)
        Cells(1, i).EntireColumn.Hidden = False
    Next
    
    ShowNamedColumns = nRetVal
End Function


'
' Makes the specified columns hidden, if they exist.  Columns are specified by name.
' If the optional argument is true, all other columns are visible.
' The function returns the number of columns hidden.
'
Function HideNamedColumns(arrColumnNames() As Variant, Optional bShowOtherColumns = True) As Long
    Dim i As Long
    Dim nRetVal As Long
    Dim strColName As Variant
    
    If (bShowOtherColumns) Then
        Columns.Hidden = False
    End If
    
    For Each strColName In arrColumnNames
        i = GetColumnByHeading(CStr(strColName), Rows(1), False)
        Cells(1, i).EntireColumn.Hidden = True
    Next
    
    HideNamedColumns = nRetVal
End Function


' ***********
'  Web Tools
' ***********
'
' Dependent on:
'   Microsoft Internet Controls library (ieframe.dll)
'
'
Sub test()
    Dim strFile As String
    Dim strURL As String
    Dim strText As String
    Dim strText2 As String
    
    strURL = "http://www.microsoft.com"
    
    strFile = GetTempFile()
    
    strText = GetWebPage(strFile, strURL)
    strText2 = ScrapeHTML(strText)
End Sub


'
' Writes the HTML contents of the specified URL to the specified file.
'
'
' TODO: Look at HTTPGet function to determine if this is still the best
' TODO:   mechanism to retrieve web data.
'
'Function GetWebPage(strFilename As String, strQuery As String, Optional strMode As String = "overwrite") As String
'    Dim ie As SHDocVw.InternetExplorer
'    Dim nFile As Integer
'    Dim strRetVal As String
'
'    Set ie = CreateObject("InternetExplorer.Application")
'
'    With ie
'        .Visible = False
'        .Navigate2 strQuery
'        Do Until .ReadyState = READYSTATE_COMPLETE
'            DoEvents
'        Loop
'
'        nFile = FreeFile
'        If (LCase(strMode) = "append") Then
'            Open strFilename For Append Shared As #nFile
'        Else
'            Open strFilename For Output Shared As #nFile
'        End If
'
'        Print #nFile, .Document.DocumentElement.InnerHTML
'        strRetVal = .Document.DocumentElement.InnerHTML
'
'        Close #nFile
'        .Quit
'    End With
'
'    Set ie = Nothing
'    GetWebPage = strRetVal
'End Function
'
'
''
'' Execute an HTTPGet on the specified URL, and return
'' the results as a string.
''
'' Example: str = HTTPGet("http://www.microsoft.com")
''
'Function HTTPGet(URL As String, Optional filename As String = "") As String
'    Dim http As Object
'    Dim script As String
'    Dim nFile As Integer
'    Dim strFilename As String
'
'    strFilename = filename
'
'    'Create Http object
'    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
'
'    'Send request To URL
'    http.Open "GET", URL
'    http.send
'
'
'    If (filename <> "") Then
'        nFile = FreeFile
'        Open strFilename For Output Shared As #nFile
'        Print #nFile, http.responseText
'        Close #nFile
'    End If
'
'    'Get response data As a string
'    HTTPGet = http.responseText
'End Function
'
'
''
'' Encodes a string into proper format to be included in an URL
''
'' From http://stackoverflow.com/questions/218181/how-can-i-url-encode-a-string-in-excel-vba
''
'Function URLEncode(StringToEncode As String, Optional SpaceAsPlus As Boolean = False) As String
'    Dim nStrLen As Long: nStrLen = Len(StringToEncode)
'    Dim i As Long
'    Dim nCharCode As Integer
'    Dim strChar As String
'    Dim strSpace As String
'
'
'    On Error GoTo Catch
'    If nStrLen > 0 Then
'        ReDim strResult(nStrLen) As String
'
'        If SpaceAsPlus Then
'            strSpace = "+"
'        Else
'            strSpace = "%20"
'        End If
'
'        For i = 1 To nStrLen
'            strChar = Mid(StringToEncode, i, 1)
'            nCharCode = Asc(strChar)
'            Select Case nCharCode
'                Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
'                    strResult(i) = strChar
'                Case 32
'                    strResult(i) = strSpace
'                Case 0 To 15
'                    strResult(i) = "%0" & Hex(nCharCode)
'                Case Else
'                    strResult(i) = "%" & Hex(nCharCode)
'                End Select
'        Next i
'
'        URLEncode = Join(strResult, "")
'    End If
'
'Finally:
'    Exit Function
'Catch:
'    URLEncode = ""
'    Resume Finally
'End Function
'
'
''
'' BUGBUG: Not sure what this function does
''
'Function ScrapeHTML(szString As String) As String
'    Dim str As String
'    Dim strHTML As String
'
'    str = szString
'    strHTML = Between(str, "<", ">")
'
'    While strHTML <> ""
'        str = Replace(str, "<" + strHTML + ">", "")
'        strHTML = Between(str, "<", ">")
'    Wend
'
'    If (str = "") Then
'        str = szString
'    End If
'
'    ScrapeHTML = str
'End Function
'


' *************************
'  Class and Object Tools
' *************************
'

Sub ListProps2()
    Dim StartLine As Long
    Dim EndLine As Long
    Dim StartCol As Long
    Dim EndCol As Long
    Dim DestCell As Range
    Dim PropName As String
    Dim Pos As Integer
    Dim TheObj As CSomeClass
    
    Set DestCell = Range("A1")
    
    Set TheObj = New CSomeClass
    TheObj.Name1 = "First Test Value"
    TheObj.Name2 = "Second Test Value"
    TheObj.Name3 = "Third Test Value"
    
    StartLine = 1
    EndLine = 99999
    StartCol = 1
    EndCol = 9999
        
    With ThisWorkbook.VBProject.VBComponents("CSomeClass").CodeModule
        Do While .Find("Property Get", StartLine, StartCol, EndLine, EndCol, True, True)
            Pos = InStr(1, .Lines(StartLine, 1), "(")
            PropName = Mid(.Lines(StartLine, 1), 14, Pos - 14)
            DestCell = PropName
            DestCell(1, 2) = CallByName(TheObj, PropName, VbGet)
            Set DestCell = DestCell(2, 1)
    
            StartLine = StartLine + 1
            EndLine = 99999
            StartCol = 1
            EndCol = 9999
        Loop
    End With
End Sub



'****************
'  Pivot Tables
'****************

'
Function MakePivot(rngSrcData As Range, rngPivotDest As Range, strPivotName As String, _
                Optional colColFields As Collection, Optional colRowFields As Collection, _
                Optional colFilterFields As Collection, Optional colSumFields As Collection) As PivotTable
'
' Creates a new pivot table from the specified source in the destination
'
' Row, column, and filter fields can be specified by adding the field names to the respective collection
'  and passing to the MakePivot function
'
' The value fields are also passed as a collection, but the operator must
'  be part of the string (e.g. Sum or amount)
'
' The function returns a reference to the new pivot table
    
    Dim wksPivot As Worksheet
    Dim pvt As PivotTable
    Dim n As Integer
    Dim strOperation As String
    Dim strField As String
    Dim func As XlConsolidationFunction
    
    Set wksPivot = rngPivotDest.Worksheet
    ' Turn off error checking in case the pivot already exists
    On Error Resume Next
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rngSrcData, _
        version:=xlPivotTableVersion12).CreatePivotTable _
        TableDestination:=rngPivotDest, TableName:=strPivotName, DefaultVersion _
        :=xlPivotTableVersion12
    On Error GoTo 0
    
    ' Clear out all fields, filters, etc (only relevant if pivot already existed)
    Set pvt = wksPivot.PivotTables(strPivotName)
    pvt.ClearTable
    
    ' Add each of the rows specified in the colRowFields argument
    If (Not colRowFields Is Nothing) Then
        For n = 1 To colRowFields.Count
            pvt.PivotFields(colRowFields.Item(n)).Orientation = xlRowField
        Next
    End If

    ' Add each of the columns specified in the colColFields argument
    If (Not colColFields Is Nothing) Then
        For n = 1 To colColFields.Count
            pvt.PivotFields(colColFields.Item(n)).Orientation = xlColumnField
        Next
    End If

    ' Add each of the filters specified in the colFilterFields argument
    If (Not colFilterFields Is Nothing) Then
        For n = 1 To colFilterFields.Count
            pvt.PivotFields(colFilterFields.Item(n)).Orientation = xlPageField
        Next
    End If

    ' Add each of the summary fields
    If (Not colSumFields Is Nothing) Then
        For n = 1 To colSumFields.Count
            strOperation = Trim(UCase(Before(colSumFields.Item(n), "of")))
            strField = Trim(After(colSumFields.Item(n), "of"))

            Select Case strOperation
                Case "SUM":
                    func = xlSum
                Case "COUNT"
                    func = xlCount
                Case "AVERAGE"
                    func = xlAverage
                Case Else
                    func = xlUnknown
            End Select

            pvt.AddDataField pvt.PivotFields(strField), colSumFields.Item(n), func
        Next

        ' Re-add fields as they get deleted for some reason after the AddDataField
'        If (Not colRowFields Is Nothing) Then
'            For n = 1 To colRowFields.Count
'                pvt.PivotFields(colRowFields.Item(n)).Orientation = xlRowField
'            Next
'        End If
'
'        If (Not colColFields Is Nothing) Then
'            For n = 1 To colColFields.Count
'                pvt.PivotFields(colColFields.Item(n)).Orientation = xlColumnField
'            Next
'        End If
    End If
    
    Set MakePivot = pvt
End Function


'Sub testSetPivotFilterFields()
'    Dim test
'    test = SetPivotFilterFields("pvtSummary", , True)
'End Sub

'Function SetPivotFilterFields(strPivotName As String, Optional colFilterFields As Collection, Optional bReplace As Boolean = False) As PivotTable
'
'    Dim wksPivot As Worksheet
'    Dim pvt As PivotTable
'    Dim fields As PivotFields
'
'    Set wksPivot = ActiveSheet
'
'    Set pvt = wksPivot.PivotTables(strPivotName)
'    wksPivot.Activate
'
'    If (bReplace <> True) Then
'        pvt.PivotFields.Clear
'    End If
'
'    Set fields = pvt.PivotFields
'
'    If (Not colFilterFields Is Nothing) Then
'        For n = 1 To colFilterFields.Count
'            pvt.PivotFields(colFilterFields.Item(n)).Orientation = xlPageField
'        Next
'    End If
'
'End Function

Function SetPivotFilterValues(strPivotName As String, strFilterField As String, Optional wksPivot As Worksheet, Optional colSet As Collection, Optional etStatus As dsFilterSettings = dsOn)
    '
    ' Sets/Clears the values of the filter fields in the specified pivot table via a collection (colSet)
    '
    ' [in] strPivotName     The string name of pivot table to set field values
    ' [in] strFilterField   The field to set filter values
    ' [in] wksPivot         The worksheet containing the pivot table to filter.  If not specified the active sheet is used.
    ' [in] colSet           Collection of field values to set in the filter
    ' [in] etStatus         Sets the values in colSet to On, Off, Toggle, Exclusive On or Exclusive Off
    '
    
    Dim pvt As PivotTable
    Dim varValue As Variant
    Dim bVisible As Boolean
    Dim colRest As New Collection
    Dim pivotItem As pivotItem
    Dim strValue As String
    Dim n As Integer
    
    On Error GoTo CleanFail
    
    ' If the target worksheet isn't specified, use the active sheet
    If wksPivot Is Nothing Then
        Set wksPivot = ActiveSheet
    End If
    
    ' Get the pivot table from the collection on the worksheet
    Set pvt = wksPivot.PivotTables(strPivotName)
    wksPivot.Activate
    
    ' If no values are specified to set, use all values in the pivot
    If colSet Is Nothing Then
        Set colSet = New Collection
        For Each varValue In pvt.PivotFields(strFilterField).PivotItems
            colSet.Add varValue.name
        Next
    End If
    
    
    ' TEMP: Dump the list of PivotField names
'    For Each varValue In pvt.PivotFields
'        Debug.Print (varValue.name)
'    Next
    
    ' Build a collection of all values that are not included in colSet (rest)
    '   by building a collection of all values in the pivot and removing the values
    '   in colSet
    '
    For Each varValue In pvt.PivotFields(strFilterField).PivotItems
        ' The if statement filters out ghost items which are in the pivot items but have no records
        ' To get rid of the ghost items, use PivotTable Analyze | Options | Data
        '    Set Number of items to retain to None
        '
'        If varValue.RecordCount > 0 Then
        colRest.Add varValue.name, varValue.name
'        Else
'            Debug.Print ("PivotItem " + varValue.name + " had no records.  Didn't add.")
'        End If
    Next
    
    For Each varValue In colSet
        colRest.Remove varValue
    Next
    
    ' Check/uncheck the values in the filter
    Select Case etStatus
        ' Turn on the specified columns, and turn off the rest (Exclusive On)
        Case dsOnExclusive:
            SetPivotFilterValues strPivotName, strFilterField, wksPivot, colSet, dsOn
            SetPivotFilterValues strPivotName, strFilterField, wksPivot, colRest, dsOff
            
        ' Turn off the specified columns, and turn on the rest (Exclusive Off)
        Case dsOffExclusive
            SetPivotFilterValues strPivotName, strFilterField, wksPivot, colRest, dsOn
            SetPivotFilterValues strPivotName, strFilterField, wksPivot, colSet, dsOff
        
        ' On, Off, and Toggle (called recursively by exclusive on and exclusive off)
        Case Else:
            pvt.PivotFields(strFilterField).EnableMultiplePageItems = True
            With pvt.PivotFields(strFilterField)
                For Each varValue In colSet
                    strValue = CStr(varValue)
                    Select Case (etStatus)
                        Case dsOn:
                            bVisible = True
                        Case dsOff:
                            bVisible = False
                        Case dsToggle:
                            bVisible = Not .PivotItems(strValue).Visible
                    End Select
                    
                    If (strValue = "(blank)") Then
                        Set pivotItem = .PivotItems.Item(.PivotItems.Count)
                    Else
                        Set pivotItem = .PivotItems(strValue)
                        pivotItem.Visible = bVisible
                    End If
                    
                Next
            End With
    End Select
CleanExit:
    Set SetPivotFilterValues = pvt
    Exit Function
    
CleanFail:
    Select Case Err.Number
        Case 1004:
            Debug.Print ("Pivot field '" + strFilterField + "' does not exist.  Exiting function.")
        Case Else:
    End Select
    Resume CleanExit
End Function



' *****************
'    Collections
' *****************

Sub ClearCollection(col As Collection)
'
' Iterates through a collection removing all values
    ' TODO: Consider returning a new collection rather than deleting each member
    
    Dim n As Long
    
    For n = 1 To col.Count
        col.Remove 1
    Next
'    Set col = New Collection
End Sub


'
' Returns the specified element from a collection, or nothing
' if the element doesn't exist in the collection
'
Function GetCollectionElement(col As Variant, element As String) As Object
    Dim objRetVal As Object
    
    On Error Resume Next
        Set objRetVal = col.Item(element)
    On Error GoTo 0
    
    Set GetCollectionElement = objRetVal
End Function


'
' Returns a delimited string containing the contents of a collection
'
Function CollectionToString(col As Collection, Optional strDelimiter = ";") As String
    
    Dim strWholeString As String
    Dim n As Long
    
    strWholeString = ""
    
    If (strDelimiter = "") Then
        strDelimiter = ";"
    Else
        If (Len(strDelimiter) > 1) Then
            strDelimiter = Left(strDelimiter, 1)
        End If
    End If
    
    For n = 1 To col.Count
        strWholeString = strWholeString + col.Item(n) + strDelimiter
    Next
    
    If (Len(strWholeString) > 0) Then
        strWholeString = Left(strWholeString, Len(strWholeString) - 1)
    End If
    
    CollectionToString = strWholeString
End Function


' Converts a collection to an array
Function CollectionToArray(col As Collection) As Variant()
    Dim varReturns() As Variant
    ReDim varReturns(col.Count)
    Dim n As Long
    
    For n = 1 To col.Count
        varReturns(n) = col(n)
    Next
    
    CollectionToArray = varReturns
End Function


' Converts an array to a collection
Function ArrayToCollection(arr() As Variant) As Collection
    Dim colReturns As Collection
    Dim n As Long
    
    Set colReturns = New Collection
    
    For n = 1 To UBound(arr)
        colReturns.Add (arr(n))
    Next
    
    Set ArrayToCollection = colReturns
End Function



' ***********
'    Names
' ***********

' Iterates through the list of names, purging those starting with a tmp prefix
Sub PurgeTempNames()
   For Each nmName In Names
      If (Left(nmName.name, 3) = "tmp") Then
         nmName.Delete
      End If
   Next
End Sub


' Checks if the specified name is defined
Function NameExists(strName As String) As Boolean
   Dim bRetVal As Boolean
   Dim nm As name
   
   On Error Resume Next
   Set nm = Names(strName)
   If (Err.Number = 0) Then
      bRetVal = True
   Else
      bRetVal = False
   End If
   On Error GoTo 0

   NameExists = bRetVal
End Function


'*************
'  INI Files
'*************

'
' Reads the specified key from an INI file
'
Function ReadIniFileString(strSection As String, strKeyname As String, Optional strIniFile As String = "") As String
    Dim nSuccess As Long
    Dim strBuffer As String * 128
    Dim nStrSize As Long
    Dim nNumChars As Long
    Dim strReturn As String
    
    nNumChars = 0
    If (strIniFile = "") Then
        strIniFile = ThisWorkbook.path + "\" + ThisWorkbook.name + ".ini"
    End If
        
    If (strSection = "" Or strKeyname = "") Then
        Debug.Print "Not enough arguments provided"
        strReturn = "ERROR"
    Else
        strBuffer = Space(128)
        nStrSize = Len(strBuffer)
        
        nNumChars = GetPrivateProfileString(strSection, strKeyname, "", strBuffer, nStrSize, strIniFile)
        If (nNumChars > 0) Then
            strReturn = Left(strBuffer, nNumChars)
        Else
            strReturn = "ERROR"
        End If
    End If
    
    ReadIniFileString = strReturn
End Function

'
' Writes the specified key to an INI file
'
Function WriteIniFileString(strSection As String, strKeyname As String, strValue As String, Optional strIniFile = "") As String
    Dim nNumChars As Long
    Dim strReturn As String
    
    nNumChars = 0
    If (strIniFile = "") Then
        strIniFile = ThisWorkbook.path + "\" + ThisWorkbook.name + ".ini"
    End If
        
    If (strSection = "" Or strKeyname = "") Then
        Debug.Print "Not enough arguments provided"
        strReturn = "ERROR"
    Else
        nNumChars = WritePrivateProfileString(strSection, strKeyname, strValue, strIniFile)
        If (nNumChars > 0) Then
            strReturn = strValue
        End If
    End If
    
    WriteIniFileString = strReturn
End Function



' ***************
'    Geocoding
' ***************

'
' Make RESP call to Bing with an address, receiving back location data in lat/long.
'
' Parameters for URL on bing are here:
'   http://alastaira.wordpress.com/2012/06/19/url-parameters-for-the-bing-maps-website/
'
Function bingAddressLookup(location As String, Optional urlBing As String) As String
    Dim strBingMapsKey As String
    Dim strResponse As String
    Dim URL As String
    Dim lat As String
    Dim lng As String
    Dim objJSON As Object
    Dim strLat As String
    Dim strLon As String
    Dim strCon As String
    
    ' Unique Bing key (https://www.bingmapsportal.com/application/index/1053340?status=NoStatus)
    strBingMapsKey = "AuQU1d84BgN9ogcYZd9GxnKHI6Pl9O-lyalqoh0bCGF5YyHlHhlk3BtkH7nZ-tSB"

    ' Create the URL for the RESP call
    URL = "http://dev.virtualearth.net/REST/v1/Locations?query=" & URLEncode(location, True) & "&maxResults=1&key=" & strBingMapsKey
    
    ' Get the response via HTTP GET
    strResponse = HTTPGet(URL)

    ' Parse the JSON response using JavaScript
    Set objJSON = ParseJSONGeoData(CStr(strResponse), strLat, strLon, strCon)
    
    ' Create the URL for the location on Bing Maps
    urlBing = "http://www.bing.com/maps/default.aspx?cp=" + strLat + "~" + strLon + "&lvl=16"
    
    ' Return the latitude, longitude and precision of the location
    bingAddressLookup = strLat & "," & strLon & "," & strCon
End Function


'
' Receives string data containing a JSON object from a Bing RESP call.  The string is converted
' to an excel object using a call to JavaScript Eval.  There are potential security ramifications
' to this method, but easier and faster than parsing the string directly.
'
' The function returns the object itself, as well as extracts relevant properties, and creates
' an URL which can be used to show the location in Bing Maps.
'
' Requires reference to Microsoft Script Control 1.0 library
'Function ParseJSONGeoData(strRawJSON As String, strLatitude As String, strLongitude As String, strPrecision As String) As Object
'    Dim objJSON As Object
'    Static objScriptEngine As Object
'
'    If (objScriptEngine Is Nothing) Then
'        Set objScriptEngine = New ScriptControl
'        objScriptEngine.Language = "JScript"
'        objScriptEngine.AddCode "function getLatitude(jsonObj) { return jsonObj.resourceSets[0].resources[0].geocodePoints[0].coordinates[0]; } "
'        objScriptEngine.AddCode "function getLongitude(jsonObj) { return jsonObj.resourceSets[0].resources[0].geocodePoints[0].coordinates[1]; } "
'        objScriptEngine.AddCode "function getPrecision(jsonObj) { return jsonObj.resourceSets[0].resources[0].confidence; } "
'    End If
'
'    Set objJSON = objScriptEngine.Eval("(" + strRawJSON + ")")
'    strLatitude = objScriptEngine.Run("getLatitude", objJSON)
'    strLongitude = objScriptEngine.Run("getLongitude", objJSON)
'    strPrecision = objScriptEngine.Run("getPrecision", objJSON)
'
'    Set ParseJSONGeoData = objJSON
'End Function


'
' Creates a GPX file from data in the specified range.
' The GPX file can be imported into a collection in Bing Maps.
'
' NOTE:  This function doesn't work yet.  It was just a copy/paste on 9/7/13.
'
' Assumptions:
' - All waypoints go into the same collection, named after the worksheet
' - Following columns exist with relevant data:
'      Latitude, Longitude, Name, Notes
'
Sub TestExportGPXFile()
    ExportGPXFile Sheets("Temp").Cells(1, 1).Resize(20, 4), "c:\temp\test.gpx"
End Sub

Sub ExportGPXFile(rng As Range, strFilename As String)
    Dim nFile As Integer
    Dim strOutput As String
    Dim strListName As String
    Dim nLatCol As Long
    Dim nLonCol As Long
    Dim nNameCol As Long
    Dim nNotesCol As Long
    Dim nRowOffset As Long
    
    strListName = rng.Worksheet.name
    nLatCol = GetColumnByHeading("Latitude", rng)
    nLonCol = GetColumnByHeading("Longitude", rng)
    nNameCol = GetColumnByHeading("Name", rng)
    nNotesCol = GetColumnByHeading("Notes", rng)
    nRowOffset = 2
    
    If (nLatCol = 0) Or (nLonCol = 0) Or (nNameCol = 0) Or (nNotesCol = 0) Then
        GoTo EXIT_EXPORTGPXFILE
    End If
    
    strOutput = "<?xml version=" + WrapInQuotes("1.0") + " encoding=" + WrapInQuotes("utf-8") + "?>" + vbNewLine
    strOutput = strOutput + "<gpx xmlns:xsi=" + WrapInQuotes("http://www.w3.org/2001/XMLSchema-instance") + _
                    " xmlns:xsd=" + WrapInQuotes("http://www.w3.org/2001/XMLSchema") + " version=" + WrapInQuotes("1.1") + " creator=" + _
                    WrapInQuotes("Bing Maps") + " xmlns=" + WrapInQuotes("http://www.topografix.com/GPX/1/1") + ">" + vbNewLine
    strOutput = strOutput + "  <metadata>" + vbNewLine
    strOutput = strOutput + "    <name>" + strListName + "</name>" + vbNewLine
    strOutput = strOutput + "    <desc />" + vbNewLine
    strOutput = strOutput + "  </metadata>" + vbNewLine
    
    While rng.Cells(nRowOffset, nLatCol) <> ""
        strOutput = strOutput + "  <wpt lat=" + WrapInQuotes(CStr(rng.Cells(nRowOffset, nLatCol))) + " lon=" + WrapInQuotes(CStr(rng.Cells(nRowOffset, nLonCol))) + ">" + vbNewLine
        strOutput = strOutput + "    <name>" + rng.Cells(nRowOffset, nNameCol) + "</name>" + vbNewLine
        strOutput = strOutput + "    <desc>" + rng.Cells(nRowOffset, nNotesCol) + "</desc>" + vbNewLine
        strOutput = strOutput + "  </wpt>" + vbNewLine
        nRowOffset = nRowOffset + 1
    Wend
    
    strOutput = strOutput + "</gpx>" + vbNewLine
    nFile = FreeFile
    Open strFilename For Output Shared As #nFile
    Print #nFile, strOutput
    Close #nFile
    
EXIT_EXPORTGPXFILE:
End Sub



' ********************
'  Regular Expression
' ********************
' boolean functin tests if regular expression test against string souce
'
' Example: RegExTest("this is a string","[A-Z]") returns False
'          This searches for capital letters in a string
Public Function RegExTest(ByRef source As String, _
                          ByRef test As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("vbscript.regexp")
        
    regex.Pattern = test
    RegExTest = regex.test(source)
End Function


' Counts the number of matches of regular expression test in source
'
' Example: RegExNumMatches("this is a string","\w+") returns 4
'          The regular expression "\w+" counts words (one or more strings of consecutive letters)
Public Function RegExNumMatches(ByRef source As String, _
                                ByRef test As String) As Integer
    Dim regex As Object
    
    Set regex = CreateObject("vbscript.regexp")
    
    regex.Pattern = test
    regex.Global = True
    
    Dim match As Object
    Set match = regex.Execute(source)
    
    RegExNumMatches = match.Count
End Function


' Returns a collection object containing all matches of regular expression test against source
Function RegExSubmatches(ByRef source As String, _
                         ByRef test As String) As Object
    Dim regex As Object
    Set regex = CreateObject("vbscript.regexp")
    
    With regex
        .Pattern = test
        .Global = True
    End With
    
    Dim match As Object
    Set match = regex.Execute(source)
    
    If match.Count > 0 Then
        Set RegExSubmatches = match(0).SubMatches
    Else
        Set RegExSubmatches = Nothing
    End If
End Function

'
' returns a regular expression object after comparing
' test to source
Function RegExMatches(ByRef source As String, _
                      ByRef test As String) As Object
    Dim regex As Object
    Dim match As Object
    
    Set regex = CreateObject("vbscript.regexp")
       
    With regex
        .Pattern = test
        .Global = True
    End With
    
    Set match = regex.Execute(source)
    Set RegExMatches = match
End Function


' Returns the first regular expression match object of comparing regular express test to source
Function RegExMatch(ByRef source As String, ByRef test As String) As String
    Dim regex As Object
    Dim match As Object
    
    Set regex = CreateObject("vbscript.regexp")
        
    With regex
        .Pattern = test
        .Global = True
    End With
    
    Set match = regex.Execute(source)
    If match.Count > 0 Then
        If match(0).SubMatches.Count > 0 Then
            RegExMatch = match(0).SubMatches(0)
        Else
            RegExMatch = ""
        End If
    Else
        RegExMatch = ""
    End If
End Function


'
' Returns a string containing only characters in source that match elements in test
'
' Example: RegExValidate("chris gemignani","aeiou") returns "ieiai"
Public Function RegExValidate(ByRef source As String, _
                              ByRef test As String) As String

    Dim s As String
    Dim regex As Object
    
    Set regex = CreateObject("vbscript.regexp")
    
    With regex
        .Pattern = test
        .Global = True
    End With
    
    Dim matches As Object
    Set matches = regex.Execute(source)
    
    s = ""
    For Each m In matches
        s = s & m
    Next m
    RegExValidate = s
End Function


Function IsPhoneNumber(str As String) As Boolean
    IsPhoneNumber = RegExTest(str, "(1[. -])?((\(\d{3}\)|\d{3})[. -]){1,2}\d{4}")
End Function


Function IsURL(str As String) As Boolean
    IsURL = RegExTest(str, "(http(s?)\:\/\/|~/|/)?([a-zA-Z]{1}([\w\-]+\.)+([\w]{2,5}))(:[\d]{1,5})?/?(\w+\.[\w]{3,4})?((\?\w+=\w+)?(&\w+=\w+)*)?")
End Function


' ***********
'   GUIDS
' ***********

Function IsGuid(strGUID As String) As Boolean
    Dim bRetVal As Boolean
    Dim strMatch As String
    
    strMatch = RegExMatch(strGUID, "^(\{){0,1}[0-9a-fA-F]{8}\-" + _
                     "[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-" + _
                     "[0-9a-fA-F]{12}(\}){0,1}$")
                     
    If (strMatch <> "") Then
        bRetVal = True
    Else
        bRetVal = False
    End If
    
    IsGuid = bRetVal
End Function



'******************
' Name/Value Pairs
'******************

Sub Test3()
    Dim nIndex As Long
    Dim str As String
    
    str = ""
    
    NVPAdd str, "Test", "Value"
    NVPAdd str, "First Track", "Value"
    NVPAdd str, "Second Track", "Value"
    NVPAdd str, "Third Track", "Value"
    NVPAdd str, "Second Track", "Mub"
    nIndex = NVPCount("Hello")
    nIndex = NVPCount("")
End Sub


Sub NVPAdd(str As String, strName As String, strValue As String, Optional bReplace As Boolean = True, Optional DelimSep As String = ";", Optional DelimAssign = ":")
    Dim nIndex As Long
    Dim nNVPairEnd As Long
    Dim strCurrent As String
    Dim bAppend As Boolean
    
    bAppend = False
    
    If (Not bReplace) Then
        bAppend = True
    Else
        nIndex = InStr(1, str, strName + DelimAssign)
        If (nIndex > 0) Then
            nNVPairEnd = InStr(nIndex + 1, str, DelimSep)
            If (nNVPairEnd = 0) Then
                nNVPairEnd = Len(str)
            End If
        Else
            bAppend = True
        End If
    End If
    
    If (bAppend) Then
        str = str + DelimSep + strName + DelimAssign + CStr(strValue)
    Else
        str = Left(str, nIndex - 1) + strName + DelimAssign + CStr(strValue) + Mid(str, nNVPairEnd)
    End If

    If (Left(str, Len(DelimSep)) = DelimSep) Then
        str = Right(str, Len(str) - Len(DelimSep))
    End If

FN_EXIT:

End Sub

Sub NVPRemove(str As String, strName As String)

End Sub

Function NVPExists(str As String, strName As String, Optional ByRef strValue As String, Optional ByRef nIndex As Long) As Boolean

End Function

Sub NVPReplace(str As String, strName As String, strValue As String, Optional bAppend As Boolean = True)

End Sub

Function NVPItem(nIndex As Long, Optional strName As String) As String

End Function


Function NVPCount(str As String, Optional DelimSep As String = ";", Optional DelimAssign = ":") As Long
    Dim nIndex As Long
    Dim nCount As Long
    
    If (Len(str) = 0) Then
        nCount = 0
        GoTo FN_EXIT
    End If
    
    nCount = 1
    nIndex = InStr(1, str, DelimSep)
    While (nIndex > 0)
        nIndex = InStr(nIndex + 1, str, DelimSep)
        nCount = nCount + 1
    Wend
    
FN_EXIT:
    NVPCount = nCount
End Function


'********************
' OneDrive Functions
'********************
' Rudimentary check if a path is the remote path to a OneDrive location
Function IsOneDrivePath(path As String)
    IsOneDrivePath = False
    
    If (Left(path, 6) = "https:") Then
        IsOneDrivePath = True
    End If
End Function


'********************
' Performance Timing
'********************
' http://en.allexperts.com/q/Excel-1059/time-milliseconds.htm
' http://support.microsoft.com/?kbid=172338

Function MicroTimer() As Currency
    Static curCounterStart As Currency, curCounterEnd As Currency, curFreq As Currency
    Static curOverhead As Currency
    Dim a As Long, i As Long

    QueryPerformanceFrequency curFreq
    
    If (curCounterStart > 0) Then
        QueryPerformanceCounter curCounterEnd
        MicroTimer = (curCounterEnd - curCounterStart) / curFreq
        curCounterStart = curCounterEnd
    Else
        QueryPerformanceCounter curCounterStart
        MicroTimer = 0
    End If
End Function


Function MicroTimerEx(Optional strAction As String = "DELTA") As Currency
    Static Ctr1 As Currency, Ctr2 As Currency, Freq As Currency
    Static Overhead As Currency, a As Long, i As Long
    

End Function


Function PerfTimer_GetTimerFrequency() As Currency
    Static curFrequency As Currency
    
    If (curFrequency <> 0) Then
        PerfTimer_GetTimerFrequency = curFrequency
    Else
        QueryPerformanceFrequency curFrequency
    End If
    
    PerfTimer_GetTimerFrequency = curFrequency
End Function


Sub Time_Addition()
    Dim Ctr1 As Currency, Ctr2 As Currency, Freq As Currency
    Dim Overhead As Currency, a As Long, i As Long
    
    Dim curDelta As Currency
    
    curDelta = MicroTimer()
    
    For i = 1 To 10000
      a = a + i
    Next i
    
    Debug.Print "Operation took "; MicroTimer(); " seconds."
    
'    QueryPerformanceFrequency Freq
'    QueryPerformanceCounter Ctr1
'    QueryPerformanceCounter Ctr2
'    Overhead = Ctr2 - Ctr1        ' determine API overhead
'    QueryPerformanceCounter Ctr1  ' time loop
'    For i = 1 To 100
'      a = a + i
'    Next i
'    QueryPerformanceCounter Ctr2
'    Debug.Print "("; Ctr1; "-"; Ctr2; "-"; Overhead; ") /"; Freq
'    Debug.Print "100 additions took";
'    Debug.Print (Ctr2 - Ctr1 - Overhead) / Freq; "seconds"
    
'    Sleep 20000
    curDelta = MicroTimer()
    Debug.Print "MicroTimer Delta: "; curDelta
End Sub


Sub Test_Timers()
    Dim Ctr1 As Currency
    Dim Ctr2 As Currency
    Dim Freq As Currency
    Dim Count1 As Long
    Dim Count2 As Long
    Dim Loops As Long

'
' Time QueryPerformanceCounter
'
    If QueryPerformanceCounter(Ctr1) Then
        QueryPerformanceCounter Ctr2
        Debug.Print "Start Value: "; Format$(Ctr1, "0.0000")
        Debug.Print "End Value: "; Format$(Ctr2, "0.0000")
        QueryPerformanceFrequency Freq
        Debug.Print "QueryPerformanceCounter minimum resolution: 1/" & Freq * 10000; " sec ("; 1 / Freq; ")"
        Debug.Print "API Overhead: "; (Ctr2 - Ctr1) / Freq; "seconds"
    Else
        Debug.Print "High-resolution counter not supported."
    End If

    '
    ' Time GetTickCount
    '
    Debug.Print
    Loops = 0
    Count1 = GetTickCount()
    Do
        Count2 = GetTickCount()
        Loops = Loops + 1
    Loop Until Count1 <> Count2
    Debug.Print "GetTickCount minimum resolution: "; (Count2 - Count1); "ms"
    Debug.Print "Took"; Loops; "loops"

    '
    ' Time timeGetTime
    '
    Debug.Print
    Loops = 0
    Count1 = timeGetTime()
    Do
        Count2 = timeGetTime()
        Loops = Loops + 1
    Loop Until Count1 <> Count2
    Debug.Print "timeGetTime minimum resolution: "; (Count2 - Count1); "ms"
    Debug.Print "Took"; Loops; "loops"
End Sub



' *****************
'   Macro Helpers
' *****************
'
' Need to reference Microsoft Visual Basic for Applications Extensibility 5.3, and enable the checkbox
' under Developer|Macro Security|Trust Access to the VBA Object model
'
'
'    Information about programmatic access to the VB IDE can be found here: http://www.cpearson.com/excel/vbe.aspx
'    That page explains how to add a reference programatically.
'
'    ThisWorkbook.VBProject.References.AddFromGuid GUID:="{0002E157-0000-0000-C000-000000000046}", Major:=5, Minor:=3
'

'
' Creates a class module from a list of variables/members specified in the current spreadsheet.
' The class name should be specified in the first row/column of the active spreadsheet.
' Property member variables are specified in row 2 and beyond.
'
' Variable types are inferred from the prefix of the variable name (sFilename implies a string)
'
Sub CreateClassModule()
    Dim colVarNames As Collection
    Dim colVarTypes As Collection
    Dim colVarFriendlyNames As Collection
    Dim colCollections As Collection
    Dim strClassName As String
    Dim strVarName As String
    Dim i As Long
    Dim j As Long
    Dim n As Long
    
    Dim strFullVarName As String
    Dim strVarType As String
    Dim strFullVarType As String
    Dim compClass As VBIDE.VBComponent
    Dim sCode As String
    
    Set colVarNames = New Collection
    Set colVarTypes = New Collection
    Set colVarFriendlyNames = New Collection
    Set colCollections = New Collection
    
    strClassName = ActiveCell
    
    i = 1
    While (ActiveCell.Offset(i) <> "")
        strFullVarName = ActiveCell.Offset(i)
        j = 1
        strVarType = ""
        While (j < Len(strFullVarName)) And (Mid(strFullVarName, j, 1) = LCase(Mid(strFullVarName, j, 1)))
            strVarType = Left(strFullVarName, j)
            strVarName = After(strFullVarName, strVarType)
            j = j + 1
        Wend
        
        Select Case strVarType
            Case "s":
                strFullVarType = "string"
            Case "dt":
                strFullVarType = "date"
            Case "n":
                strFullVarType = "long"
            Case "int":
                strFullVarType = "int"
            Case "f":
                strFullVarType = "double"
            Case "col":
                strFullVarType = "collection"
                colCollections.Add strFullVarName
            Case Else
                strFullVarType = "variant"
        End Select
    
        colVarNames.Add strFullVarName
        colVarFriendlyNames.Add strVarName
        colVarTypes.Add strFullVarType
        
        i = i + 1
    Wend
        
    DeleteModule (strClassName)
    Set compClass = AddClassToProject(strClassName)
    
    For n = 1 To colVarNames.Count
        sCode = sCode + "Private " + colVarNames.Item(n) + " as " + colVarTypes.Item(n) + vbNewLine
    Next
    
    sCode = sCode + vbNewLine + vbNewLine
    
    sCode = sCode + "Private Sub class_initialize()" + vbNewLine
    sCode = sCode + vbTab + "Debug.Print (" + Chr(34) + "Initializing " + strClassName + Chr(34) + ")" + vbNewLine
    For i = 1 To colCollections.Count
        strFullVarName = colCollections(i)
        sCode = sCode + vbTab + "set " + strFullVarName + " = new collection" + vbNewLine
    Next
    sCode = sCode + "End Sub" + vbNewLine
    sCode = sCode + vbNewLine
    sCode = sCode + "Private Sub class_terminate()" + vbNewLine
    sCode = sCode + vbTab + "Debug.Print (" + Chr(34) + "Destroying " + strClassName + Chr(34) + ")" + vbNewLine
    For i = 1 To colCollections.Count
        strFullVarName = colCollections(i)
        sCode = sCode + vbTab + "set " + strFullVarName + " = Nothing" + vbNewLine
    Next
    sCode = sCode + "End Sub" + vbNewLine + vbNewLine
    sCode = sCode + vbNewLine
    For n = 1 To colVarNames.Count
        sCode = sCode + "Public Property Get " + colVarFriendlyNames.Item(n) + " As " + colVarTypes.Item(n) + vbNewLine
        sCode = sCode + "   " + colVarFriendlyNames.Item(n) + " = " + colVarNames.Item(n) + vbNewLine
        sCode = sCode + "End Property" + vbNewLine
        sCode = sCode + vbNewLine
        sCode = sCode + "Public Property Let " + colVarFriendlyNames.Item(n) + "(ByVal " + colVarNames.Item(n) + "In as " + colVarTypes.Item(n) + ")" + vbNewLine
        sCode = sCode + "   " + colVarNames.Item(n) + " = " + colVarNames.Item(n) + "In" + vbNewLine
        sCode = sCode + "End Property" + vbNewLine
        sCode = sCode + vbNewLine
    Next

    ' Close #nFile
    compClass.CodeModule.InsertLines 1, sCode
    
    Set colVarNames = Nothing
    Set colVarTypes = Nothing
End Sub


'Function AddClassToProject(sModuleName As String) As VBIDE.VBComponent
'    Dim VBProj As VBIDE.VBProject
'    Dim VBComp As VBIDE.VBComponent
'
'    Set VBProj = ThisWorkbook.VBProject
'    Set VBComp = VBProj.VBComponents.Add(vbext_ct_ClassModule)
'    VBComp.Name = sModuleName
'
'    Set AddClassToProject = VBComp
'End Function
'
'
'Function DeleteModule(sModuleName As String) As Boolean
'    Dim VBProj As VBIDE.VBProject
'    Dim VBComp As VBIDE.VBComponent
'    Dim bRetVal As Boolean
'
'    bRetVal = False
'
'    Set VBProj = ThisWorkbook.VBProject
'
'    On Error Resume Next
'    Set VBComp = VBProj.VBComponents(sModuleName)
'    If Err.Number = 0 Then
'        bRetVal = True
'    End If
'    On Error GoTo 0
'
'    VBProj.VBComponents.Remove VBComp
'
'    DeleteModule = bRetVal
'End Function




'
' Tips
'
' R1C1 References:
' - Use the .Address property of a cell to convert between R1C1 and $A$1
' - Use RefersTo to define a name as $A$1, use RefersToR1C1 to define a name as r1c1
' - Setting a formula using R1C1:      ActiveCell.FormulaR1C1 = "=AVERAGE(RC[-7]:RC[-3])"
'
'
' Dialog Tips
'
' - Set radio buttons to a height of 16
' - Set spacing between to 12
'
'
' Class/Object Tips
'
' - Use CallByName(object, "Test", VbMethod) to call the "Test" method on the specified object
' - Add a reference to "Microsoft Visual Basic for Applications Extensibility 5.3"
'   to get at the codemodule for a VB project
' - Use the "type" keyword to define a structure within VBA
' - Reference on classes is at http://msdn2.microsoft.com/en-us/library/aa164936(office.10).aspx
'
'
' Variable Tips
'
' - Compare static variables to Nothing to determine if they are initialized
'       static rngSelection as range
'       if (rngSelection is Nothing) then
'           set rngSelection = Selection
'       endif
'
'
' Documentation/Comments
' - Use the Comment Block/Uncomment Block toolbar items (added from View/Toolbars/Customize|Edit menu to
'   comment and uncomment blocks of code at once.
'
