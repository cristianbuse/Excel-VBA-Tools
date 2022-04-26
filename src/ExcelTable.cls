VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''=============================================================================
''' Excel VBA Tools
''' -----------------------------------------------
''' https://github.com/cristianbuse/Excel-VBA-Tools
''' -----------------------------------------------
''' MIT License
'''
''' Copyright (c) 2022 Ion Cristian Buse
'''
''' Permission is hereby granted, free of charge, to any person obtaining a copy
''' of this software and associated documentation files (the "Software"), to
''' deal in the Software without restriction, including without limitation the
''' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
''' sell copies of the Software, and to permit persons to whom the Software is
''' furnished to do so, subject to the following conditions:
'''
''' The above copyright notice and this permission notice shall be included in
''' all copies or substantial portions of the Software.
'''
''' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
''' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
''' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
''' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
''' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
''' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
''' IN THE SOFTWARE.
'''=============================================================================

''==============================================================================
'' Description:
'' Allows easy manipulation of tabular data in Excel within a single class
'' Can be initialized via:
''    - InitFromListObject - 1 row headers always non-blank and unique
''    - InitFromRange      - joins multi header rows and makes them unique
'' Can return the following arrays:
''    - ColumnFormats      - a copy of the internal formats array
''    - DataByVal          - a copy of the internal data array
''    - DataByRef          - data array wrapped in a ByRef Variant to avoid copy
''    - HeadersArray       - a copy of the internal headers array
'' Has the following utilities:
''    - ColumnsCount       - returns the nummber of headers/columns
''    - HeaderAtIndex      - returns the header string at a given index
''    - HeaderExists       - checks if a header string exists
''    - IndexForHeader     - returns the index for a header string
''    - RowsCount          - returns the number of data rows
''    - Self               - returns the instance
''==============================================================================

Option Explicit

Private m_data() As Variant
Private m_columnFormats() As Variant
Private m_headers() As Variant
Private m_headersColl As Collection

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Class Events
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Class_Initialize()
    Set m_headersColl = New Collection
End Sub
Private Sub Class_Terminate()
    Set m_headersColl = Nothing
End Sub

'*******************************************************************************
'Init from already built arrays of headers and data
'*******************************************************************************
Public Sub InitFromArrays(ByRef arrHeaders As Variant _
                        , ByRef arrData As Variant)
    Const methodName As String = "InitFromArrays"
    If GetArrayDimsCount(arrHeaders) <> 2 Then
        Err.Raise 91, TypeName(Me) & "." & methodName, "Expected 2D headers"
    ElseIf GetArrayDimsCount(arrData) <> 2 Then
        Err.Raise 91, TypeName(Me) & "." & methodName, "Expected 2D data"
    ElseIf UBound(arrHeaders, 1) <> LBound(arrHeaders, 1) Then
        Err.Raise 91, TypeName(Me) & "." & methodName, "Expected 1 row headers"
    ElseIf LBound(arrHeaders, 2) <> LBound(arrData, 2) _
        Or UBound(arrHeaders, 2) <> UBound(arrData, 2) Then
        Err.Raise 91, TypeName(Me) & "." & methodName, "Invalid column bounds"
    End If
    '
    m_headers = arrHeaders
    FixHeaderDuplicates m_headers
    InitHeaderCollection m_headers
    '
    UnlockDataSizeIfNeeded
    m_data = arrData
    LockDataSize
End Sub
Private Function GetArrayDimsCount(ByRef arr As Variant) As Long
    Const MAX_DIMENSION As Long = 60
    Dim dimension As Long
    Dim tempBound As Long
    '
    On Error GoTo FinalDimension
    For dimension = 1 To MAX_DIMENSION
        tempBound = LBound(arr, dimension)
    Next dimension
Exit Function
FinalDimension:
    GetArrayDimsCount = dimension - 1
End Function

'*******************************************************************************
'A ListObject header/column names are never blank and are always unique
'*******************************************************************************
Public Sub InitFromListObject(ByVal tbl As ListObject)
    If tbl Is Nothing Then
        Err.Raise 91, TypeName(Me) & ".InitFromListObject", "Table not set"
    End If
    RangeValues(m_headers) = tbl.HeaderRowRange
    InitHeaderCollection m_headers
    UnlockDataSizeIfNeeded
    RangeValues(m_data) = tbl.DataBodyRange
    LockDataSize
    Formats(m_columnFormats) = tbl.DataBodyRange
End Sub

'*******************************************************************************
'Reads a 2D array of values from a range even if range has one cell
'*******************************************************************************
Private Property Let RangeValues(ByRef arr() As Variant, ByVal rng As Range)
    If rng Is Nothing Then
        On Error Resume Next
        Erase arr
        On Error GoTo 0
    ElseIf rng.Cells.Count = 1 Then
        ReDim arr(1 To 1, 1 To 1) As Variant
        arr(1, 1) = rng.Value2
    Else
        arr = rng.Value2
    End If
End Property

'*******************************************************************************
'Mimics a static array - data cannot be resized when retrieved via 'DataByRef'
'https://github.com/cristianbuse/VBA-MemoryTools/blob/master/src/LibMemory.bas
'*******************************************************************************
Private Sub LockDataSize()
    Const FADF_STATIC As Long = &H2
    Const FADF_FIXEDSIZE As Long = &H10
    Const FADF_HAVEVARTYPE As Long = &H80
    Const F_STATIC As Long = FADF_HAVEVARTYPE Or FADF_FIXEDSIZE Or FADF_STATIC
    '
    If Me.RowsCount > 0 Then
        LibMemory.MemInt(LibMemory.ArrPtr(m_data) + 2) = F_STATIC
    End If
End Sub
Private Sub UnlockDataSizeIfNeeded()
    If Me.RowsCount > 0 Then
        Const FADF_HAVEVARTYPE As Long = &H80
        LibMemory.MemInt(LibMemory.ArrPtr(m_data) + 2) = FADF_HAVEVARTYPE
    End If
End Sub

'*******************************************************************************
'Init headers collection for easy/fast name to index mapping
'*******************************************************************************
Private Sub InitHeaderCollection(ByRef arr() As Variant)
    Dim headerName As Variant
    Dim i As Long
    '
    If m_headersColl.Count > 0 Then Set m_headersColl = New Collection
    i = LBound(arr, 1)
    For Each headerName In arr
        m_headersColl.Add i, headerName
        i = i + 1
    Next headerName
End Sub

'*******************************************************************************
'Reads a 1D array of column formats
'*******************************************************************************
Private Property Let Formats(ByRef arr() As Variant, ByVal rng As Range)
    If rng Is Nothing Then
        On Error Resume Next
        Erase arr
        On Error GoTo 0
    Else
        Dim i As Long
        ReDim arr(1 To rng.Columns.Count) As Variant
        For i = LBound(arr, 1) To UBound(arr, 1)
            arr(i) = GetRangeFormat(rng.Columns(i))
        Next i
    End If
End Property

'*******************************************************************************
'Returns the format of a given range
'*******************************************************************************
Private Function GetRangeFormat(ByRef rng As Range) As String
    Const generalFormat As String = "General"
    With rng
        If IsNull(.NumberFormat) Then
            GetRangeFormat = generalFormat
        Else
            GetRangeFormat = .NumberFormat
        End If
    End With
End Function

'*******************************************************************************
'When initializing headers from a Range:
'   - headers can be blank and are not always unique
'   - headers could be spanning multiple rows
'   - headers can be centered across selection or can be part of merged cells
'This method will clean headers (wrong data types and merged/centereed) and join
'   them into a single row of unique values
'Also merged data (not headers) will be duplicated across the merged area
'*******************************************************************************
Public Sub InitFromRange(ByVal rng As Range _
                       , Optional ByVal headerRowsCount As Long = 1 _
                       , Optional ByVal dataRowsToIgnore As Long = 0)
    Const methodName As String = "InitFromRange"
    If rng Is Nothing Then
        Err.Raise 91, TypeName(Me) & "." & methodName, "Range not set"
    ElseIf rng.Areas.Count > 1 Then
        Err.Raise 5, TypeName(Me) & "." & methodName, "Range not contiguous"
    ElseIf headerRowsCount < 1 Or headerRowsCount > rng.Rows.Count Then
        Err.Raise 5, TypeName(Me) & "." & methodName, "Not enough header rows"
    ElseIf headerRowsCount + dataRowsToIgnore > rng.Rows.Count Then
        Err.Raise 5, TypeName(Me) & "." & methodName, "Not enough rows"
    End If
    '
    Dim tempHeaders() As Variant
    Dim nonDataRows As Long
    Dim rCount As Long
    Dim rngData As Range
    '
    CleanHeaders(tempHeaders) = rng.Resize(RowSize:=headerRowsCount)
    JoinHeaderRows(m_headers) = tempHeaders
    FixHeaderDuplicates m_headers
    InitHeaderCollection m_headers
    '
    nonDataRows = headerRowsCount + dataRowsToIgnore
    rCount = rng.Rows.Count - nonDataRows
    If rCount > 0 Then
        Set rngData = rng.Resize(RowSize:=rCount).Offset(nonDataRows)
    End If
    UnlockDataSizeIfNeeded
    CleanData(m_data) = rngData
    LockDataSize
    '
    Formats(m_columnFormats) = rngData
End Sub

'*******************************************************************************
'Cleans the headers by:
'   - replacing errors with the text counterpart
'   - replacing dates with formatted text
'   - converting all other data types to text
'   - repeating merge cells values for the first row of the merge area
'   - repeating values for cells formatted as 'centered across selection'
'*******************************************************************************
Private Property Let CleanHeaders(ByRef arr() As Variant, ByVal rng As Range)
    Dim topRow As Long: topRow = rng.Row - 1
    Dim leftCol As Long: leftCol = rng.Column - 1
    Dim i As Long
    Dim j As Long
    Dim rowHasMerge As Boolean
    Dim v As Variant
    Dim cellHasMerge As Boolean
    Dim arrDates() As Variant
    Dim c As Range
    '
    RangeValues(arr) = rng
    arrDates = rng.Value
    For i = LBound(arr, 1) To UBound(arr, 1)
        rowHasMerge = RangeHasMerge(rng.Rows(i))
        For j = LBound(arr, 2) To UBound(arr, 2)
            v = arr(i, j)
            Select Case VarType(v)
            Case vbBoolean
                arr(i, j) = UCase$(v)
            Case vbEmpty
                If j > 1 Then
                    cellHasMerge = False
                    If rowHasMerge Then
                        Set c = rng.Cells(i, j)
                        If c.MergeCells Then
                            cellHasMerge = True
                            If c.MergeArea.Row - topRow = i Then
                                If c.MergeArea.Column - leftCol < j Then
                                    arr(i, j) = arr(i, j - 1)
                                End If
                            End If
                        End If
                    End If
                    If Not cellHasMerge Then
                        If CanCopyLeftForCentered(rng, i, j) Then
                            arr(i, j) = arr(i, j - 1)
                        End If
                    End If
                End If
            Case vbError
                arr(i, j) = ExcelErrToString(v)
            Case vbString 'Do nothing
            Case Else
                If IsDate(arrDates(i, j)) Then
                    Const fDate As String = "dd-mmm-yyyy"
                    Const fTime As String = "dd-mmm-yyyy hh:mm:ss"
                    arr(i, j) = Format$(v, IIf(Int(v) = v, fDate, fTime))
                Else
                    arr(i, j) = CStr(v)
                End If
            End Select
        Next j
    Next i
End Property
Private Function CanCopyLeftForCentered(ByRef rng As Range _
                                      , ByRef rIndex As Long _
                                      , ByRef cIndex As Long)
    Dim ha As Long
    '
    ha = rng.Cells(rIndex, cIndex).HorizontalAlignment
    If ha = xlCenterAcrossSelection Then
        ha = rng.Cells(rIndex, cIndex - 1).HorizontalAlignment
        CanCopyLeftForCentered = (ha = xlCenterAcrossSelection)
    End If
End Function

'*******************************************************************************
'Checks if a range has any merged cells
'*******************************************************************************
Private Function RangeHasMerge(ByRef rng As Range) As Boolean
    If rng Is Nothing Then Exit Function
    If IsNull(rng.MergeCells) Then
        RangeHasMerge = True
    Else
        RangeHasMerge = rng.MergeCells
    End If
End Function

'*******************************************************************************
'Converts an Excel Error to its String counterpart
'*******************************************************************************
Private Function ExcelErrToString(ByRef excelError As Variant) As String
    Select Case excelError
    Case CVErr(xlErrBlocked):     ExcelErrToString = "#BLOCKED!"
    Case CVErr(xlErrCalc):        ExcelErrToString = "#CALC!"
    Case CVErr(xlErrConnect):     ExcelErrToString = "#CONNECT!"
    Case CVErr(xlErrDiv0):        ExcelErrToString = "#DIV/0"
    Case CVErr(xlErrField):       ExcelErrToString = "#FIELD!"
    Case CVErr(xlErrGettingData): ExcelErrToString = "#GETTING_DATA"
    Case CVErr(xlErrNA):          ExcelErrToString = "#N/A"
    Case CVErr(xlErrName):        ExcelErrToString = "#NAME?"
    Case CVErr(xlErrNull):        ExcelErrToString = "#NULL!"
    Case CVErr(xlErrNum):         ExcelErrToString = "#NUM!"
    Case CVErr(xlErrRef):         ExcelErrToString = "#REF!"
    Case CVErr(xlErrSpill):       ExcelErrToString = "#SPILL!"
    Case CVErr(xlErrUnknown):     ExcelErrToString = "#UNKNOWN!"
    Case CVErr(xlErrValue):       ExcelErrToString = "#VALUE!"
    Case Else: Err.Raise 5, TypeName(Me) & ".ExcelErrToString", "Invalid error"
    End Select
End Function

'*******************************************************************************
'Joins multi-row headers into one row
'*******************************************************************************
Private Property Let JoinHeaderRows(ByRef outArr() As Variant _
                                  , ByRef inArr() As Variant)
    If UBound(inArr, 1) - LBound(inArr, 1) = 0 Then
        outArr = inArr
        Exit Property
    End If
    '
    Const delimiter As String = " "
    Dim i As Long
    Dim j As Long
    Dim v As Variant
    Dim tempHeader As String
    Dim hasStartValue As Boolean
    Dim hasValue As Boolean
    '
    ReDim outArr(1 To 1, LBound(inArr, 2) To UBound(inArr, 2))
    For j = LBound(inArr, 2) To UBound(inArr, 2)
        tempHeader = vbNullString
        hasStartValue = False
        For i = LBound(inArr, 1) To UBound(inArr, 1)
            v = inArr(i, j)
            Select Case VarType(v)
                Case vbEmpty:  hasValue = False
                Case vbString: hasValue = (LenB(v) > 0)
                Case Else:     hasValue = True
            End Select
            If hasValue Then
                If hasStartValue Then
                    tempHeader = tempHeader & delimiter & v
                Else
                    tempHeader = v
                End If
                hasStartValue = True
            End If
        Next i
        outArr(1, j) = tempHeader
    Next j
End Property

'*******************************************************************************
'To make columns unique and non-blank, same strategy is applied as in Excel's
'   ListObjects. This strategy mimics the conversion of the header range to a
'   'ListObject.HeaderRowRange'. Blank headers are replaced with 'Column' and
'   then an integer suffix is used to make all headers unique. For each
'   duplicated header the suffix is incremented until the header + suffix is
'   unique. The suffix does not get decremented for the next duplicate!
'*******************************************************************************
Private Sub FixHeaderDuplicates(ByRef arr() As Variant)
    Const blankHeader As String = "Column"
    Dim j As Long
    Dim v As Variant
    Dim coll As New Collection
    Dim hasBlank As Boolean
    Dim suffix As Long
    Dim tempHeader As String
    '
    On Error Resume Next
    j = LBound(arr, 2)
    For Each v In arr
        If LenB(v) = 0 Then
            hasBlank = True
            arr(1, j) = blankHeader 'Do NOT add blankHeader to Collection!
        End If
        coll.Add j, v
        j = j + 1
    Next v
    If Err.Number = 0 And Not hasBlank Then Exit Sub
    '
    If arr(1, 1) = blankHeader Then suffix = 1 Else suffix = 2
    j = LBound(arr, 2)
    For Each v In arr
        If coll.Item(v) <> j Then
            Do
                Err.Clear
                tempHeader = v & suffix
                coll.Add j, tempHeader
                suffix = suffix + 1
            Loop Until Err.Number = 0
            arr(1, j) = tempHeader
        End If
        j = j + 1
    Next v
    On Error GoTo 0
End Sub

'*******************************************************************************
'Repeats merge cells values for the entire merge area
'*******************************************************************************
Private Property Let CleanData(ByRef arr() As Variant, ByVal rng As Range)
    RangeValues(arr) = rng
    If Not RangeHasMerge(rng) Then Exit Property
    '
    Dim topRow As Long: topRow = rng.Row - 1
    Dim leftCol As Long: leftCol = rng.Column - 1
    Dim i As Long
    Dim j As Long
    '
    For j = LBound(arr, 2) To UBound(arr, 2)
        If RangeHasMerge(rng.Columns(j)) Then
            For i = LBound(arr, 1) To UBound(arr, 1)
                If VarType(arr(i, j)) = vbEmpty Then
                    With rng.Cells(i, j)
                        If .MergeCells Then
                            If .MergeArea.Row - topRow < i Then
                                arr(i, j) = arr(i - 1, j)
                            ElseIf .MergeArea.Column - leftCol < j Then
                                arr(i, j) = arr(i, j - 1)
                            End If
                        End If
                    End With
                End If
            Next i
        End If
    Next j
End Property

'===============================================================================
'Columns Count
'===============================================================================
Public Property Get ColumnCount() As Long
    ColumnCount = m_headersColl.Count
End Property

'===============================================================================
'Rows Count
'===============================================================================
Public Property Get RowsCount() As Long
    On Error Resume Next
    RowsCount = UBound(m_data, 1) - LBound(m_data, 1) + 1
    On Error GoTo 0
End Property

'*******************************************************************************
'Get header index by name
'Can throw! Use HeaderExists before calling this!
'*******************************************************************************
Public Function IndexForHeader(ByVal headerName As String) As Long
    IndexForHeader = m_headersColl.Item(headerName)
End Function

'*******************************************************************************
'Check if header exists by name
'Does not throw!
'*******************************************************************************
Public Function HeaderExists(ByVal headerName As String) As Boolean
    On Error Resume Next
    m_headersColl.Item headerName
    HeaderExists = (Err.Number = 0)
    On Error GoTo 0
End Function

'*******************************************************************************
'Get header name by index
'*******************************************************************************
Public Function HeaderAtIndex(ByVal hIndex As Long) As String
    HeaderAtIndex = m_headers(hIndex)
End Function

'*******************************************************************************
'Get headers array
'*******************************************************************************
Public Function HeadersArray() As Variant()
    HeadersArray = m_headers
End Function

'*******************************************************************************
'Get a copy of the data array
'*******************************************************************************
Public Function DataByVal() As Variant()
    DataByVal = m_data
End Function

'*******************************************************************************
'Wrap the data array into a ByRef Variant - does not copy the array. Needs:
'https://github.com/cristianbuse/VBA-MemoryTools/blob/master/src/LibMemory.bas
'The data array is made 'static' at Init so it cannot be resized but values
'   can still be updated ByRef or erased
'*******************************************************************************
Public Function DataByRef() As Variant
    DataByRef = LibMemory.GetArrayByRef(m_data)
End Function

'*******************************************************************************
'Get array of formats
'*******************************************************************************
Public Function ColumnFormats() As Variant()
    ColumnFormats = m_columnFormats
End Function

'*******************************************************************************
'Get self instance. Useful in 'With New ...' blocks
'*******************************************************************************
Public Function Self() As ExcelTable
    Set Self = Me
End Function