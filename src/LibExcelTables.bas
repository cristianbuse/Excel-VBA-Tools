Attribute VB_Name = "LibExcelTables"
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

'*******************************************************************************
'' Public/Exposed methods:
''    - AddListRows
''    - DeleteListRows
''    - GetListObject
''    - IsListObjectFiltered
''    - SortListObjectIfNeeded
'*******************************************************************************

Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "LibExcelTables"

'*******************************************************************************
'Adds rows to a ListObject and returns the corresponding added Range
'Parameters:
'   - tbl: the table to add rows to
'   - [rowsToAdd]: the number of rows to add. Default is 1
'   - [startRow]: the row index from where to start adding. Default is 0 in
'       which case the rows would be appended at the end of the table
'   - [doEntireSheetRow]:
'       * TRUE  - adds entire rows including left and right of the target table
'       * FALSE - adds rows only below the table bounds shifting down (default)
'Raises error:
'   -    5: if 'rowsToAdd' is less than 1
'   -    9: if 'startRow' is invalid
'   -   91: if 'tbl' is not set
'   - 1004: if adding rows failed due to worksheet being protected while the
'           UserInterfaceOnly flag is set to False
'*******************************************************************************
Public Function AddListRows(ByVal tbl As ListObject _
                          , Optional ByVal rowsToAdd As Long = 1 _
                          , Optional ByVal startRow As Long = 0 _
                          , Optional ByVal doEntireSheetRow As Boolean = False _
) As Range
    Const fullMethodName As String = MODULE_NAME & ".AddListRows"
    Dim isSuccess As Boolean
    '
    If tbl Is Nothing Then
        Err.Raise 91, fullMethodName, "Table object not set"
    ElseIf startRow < 0 Or startRow > tbl.ListRows.Count + 1 Then
        Err.Raise 9, fullMethodName, "Invalid start row index"
    ElseIf rowsToAdd < 1 Then
        Err.Raise 5, fullMethodName, "Invalid number of rows to add"
    End If
    If startRow = 0 Then startRow = tbl.ListRows.Count + 1
    '
    UnfilterIfNeeded tbl
    '
    If startRow = tbl.ListRows.Count + 1 Then
        isSuccess = AppendListRows(tbl, rowsToAdd, doEntireSheetRow)
    Else
        isSuccess = InsertListRows(tbl, rowsToAdd, startRow, doEntireSheetRow)
    End If
    If Not isSuccess Then
        If tbl.Parent.ProtectContents And Not tbl.Parent.ProtectionMode Then
            Err.Raise 1004, fullMethodName, "Parent sheet is macro protected"
        Else
            Err.Raise 5, fullMethodName, "Cannot append rows"
        End If
    End If
    Set AddListRows = tbl.ListRows(startRow).Range.Resize(RowSize:=rowsToAdd)
End Function

'*******************************************************************************
'Utility for 'AddListRows' and 'DeleteListRows' methods
'*******************************************************************************
Private Sub UnfilterIfNeeded(ByVal tbl As ListObject)
    On Error Resume Next
    If IsListObjectFiltered(tbl) Then tbl.AutoFilter.ShowAllData
    If tbl.Parent.AutoFilterMode Then tbl.Parent.AutoFilter.ShowAllData
    On Error GoTo 0
End Sub

'*******************************************************************************
'Utility for 'AddListRows' method
'Inserts rows into a ListObject. Does not append!
'*******************************************************************************
Private Function InsertListRows(ByVal tbl As ListObject _
                              , ByVal rowsToInsert As Long _
                              , ByVal startRow As Long _
                              , ByVal doEntireSheetRow As Boolean) As Boolean
    Dim rngInsert As Range
    Dim fOrigin As XlInsertFormatOrigin: fOrigin = xlFormatFromLeftOrAbove
    Dim needsHeaders As Boolean
    '
    If startRow = 1 Then
        If Not tbl.ShowHeaders Then
            If tbl.Parent.ProtectContents Then
                Exit Function 'Not sure possible without headers
            Else
                needsHeaders = True
            End If
        End If
        fOrigin = xlFormatFromRightOrBelow
    End If
    '
    Set rngInsert = tbl.ListRows(startRow).Range.Resize(RowSize:=rowsToInsert)
    If doEntireSheetRow Then Set rngInsert = rngInsert.EntireRow
    '
    On Error Resume Next
    If needsHeaders Then tbl.ShowHeaders = True
    rngInsert.Insert xlShiftDown, fOrigin
    If needsHeaders Then tbl.ShowHeaders = False
    InsertListRows = (Err.Number = 0)
    On Error GoTo 0
End Function

'*******************************************************************************
'Utility for 'AddListRows' method
'Appends rows to the bottom of a ListObject. Does not insert!
'*******************************************************************************
Private Function AppendListRows(ByVal tbl As ListObject _
                              , ByVal rowsToAppend As Long _
                              , ByVal doEntireSheetRow As Boolean) As Boolean
    If tbl.ListRows.Count = 0 Then
        If Not UpgradeInsertRow(tbl) Then Exit Function
        If rowsToAppend = 1 Then
            AppendListRows = True
            Exit Function
        End If
        rowsToAppend = rowsToAppend - 1
    End If
    '
    Dim rngToAppend As Range
    Dim isProtected As Boolean: isProtected = tbl.Parent.ProtectContents
    '
    On Error GoTo ErrorHandler
    If isProtected And tbl.ShowTotals Then
        Set rngToAppend = tbl.TotalsRowRange
    ElseIf isProtected Then
        Set rngToAppend = AutoExpandOneRow(tbl)
    Else
        Set rngToAppend = tbl.Range.Rows(tbl.Range.Rows.Count + 1)
    End If
    '
    Set rngToAppend = rngToAppend.Resize(RowSize:=rowsToAppend)
    If doEntireSheetRow Then Set rngToAppend = rngToAppend.EntireRow
    rngToAppend.Insert xlShiftDown, xlFormatFromLeftOrAbove
    '
    If isProtected And tbl.ShowTotals Then 'Fix formatting
        tbl.ListRows(1).Range.Copy
        With tbl.ListRows(tbl.ListRows.Count - rowsToAppend + 1).Range
            .Resize(RowSize:=rowsToAppend).PasteSpecial xlPasteFormats
        End With
    ElseIf isProtected Then 'Delete the autoExpand row
        tbl.ListRows(tbl.ListRows.Count).Range.Delete xlShiftUp
    Else 'Resize table
        tbl.Resize tbl.Range.Resize(tbl.Range.Rows.Count + rowsToAppend)
    End If
    AppendListRows = True
Exit Function
ErrorHandler:
    AppendListRows = False
End Function

'*******************************************************************************
'Utility for 'AppendListRows' method
'Transforms the Insert row into a usable ListRow
'*******************************************************************************
Private Function UpgradeInsertRow(ByVal tbl As ListObject) As Boolean
    If tbl.InsertRowRange Is Nothing Then Exit Function
    If tbl.Parent.ProtectContents And Not tbl.ShowHeaders Then
        Exit Function 'Not implemented - can be done using a few inserts
    Else
        Dim needsHeaders As Boolean: needsHeaders = Not tbl.ShowHeaders
        '
        If needsHeaders Then tbl.ShowHeaders = True
        tbl.InsertRowRange.Insert xlShiftDown, xlFormatFromLeftOrAbove
        If needsHeaders Then tbl.ShowHeaders = False
    End If
    UpgradeInsertRow = True
End Function

'*******************************************************************************
'Utility for 'AppendListRows' method
'Adds one row via auto expand if the worksheet is protected and totals are off
'*******************************************************************************
Private Function AutoExpandOneRow(ByVal tbl As ListObject) As Range
    If Not tbl.Parent.ProtectContents Then Exit Function
    If tbl.ShowTotals Then Exit Function
    '
    Dim ac As AutoCorrect: Set ac = Application.AutoCorrect
    Dim isAutoExpand As Boolean: isAutoExpand = ac.AutoExpandListRange
    Dim tempRow As Range: Set tempRow = tbl.Range.Rows(tbl.Range.Rows.Count + 1)
    '
    If Not isAutoExpand Then ac.AutoExpandListRange = True
    tempRow.Insert xlShiftDown, xlFormatFromLeftOrAbove
    Set AutoExpandOneRow = tempRow.Offset(-1, 0)
    Const arbitraryValue As Long = 1 'Must not be Empty/Null/""
    AutoExpandOneRow.Value2 = arbitraryValue 'AutoExpand is triggered
    If Not isAutoExpand Then ac.AutoExpandListRange = False 'Revert if needed
End Function

'*******************************************************************************
'Deletes rows from a dynamic table (ListObject)
'Parameters:
'   - tbl: the table to delete from
'   - [rowsToDelete]: the number of rows to delete. Default is 1
'   - [startRow]: the row index from where to start deleting. Default is 0 in
'       which case the rows would be deleted from the end of the table
'   - [doEntireSheetRow]:
'       * TRUE  - deletes the entire rows including left and right of 'tbl'
'       * FALSE - deletes only within the table bounds shifting up (default)
'Raises error:
'   -    5: if startRow is out of bounds
'           if table has not enough rows to delete
'           if deletion fails
'   -   91: if target table is not set
'   - 1004: if parent worksheet has contents macro protected
'*******************************************************************************
Public Sub DeleteListRows(ByVal tbl As ListObject _
                        , Optional ByVal rowsToDelete As Long = 1 _
                        , Optional ByVal startRow As Long = 0 _
                        , Optional ByVal doEntireSheetRow As Boolean = False _
)
    Const fullMethodName As String = MODULE_NAME & ".DeleteListRows"
    '
    If tbl Is Nothing Then
        Err.Raise 91, fullMethodName, "Table object not set"
    ElseIf tbl.ListRows.Count = 0 Then
        Err.Raise 5, fullMethodName, "Table has no rows"
    ElseIf rowsToDelete < 1 Or rowsToDelete > tbl.ListRows.Count Then
        Err.Raise 5, fullMethodName, "Invalid number of rows to delete"
    ElseIf startRow < 0 Or startRow + rowsToDelete - 1 > tbl.ListRows.Count Then
        Err.Raise 5, fullMethodName, "Invalid start row"
    End If
    If startRow = 0 Then startRow = tbl.ListRows.Count - rowsToDelete + 1
    '
    UnfilterIfNeeded tbl
    '
    Dim rng As Range
    Set rng = tbl.ListRows(startRow).Range.Resize(RowSize:=rowsToDelete)
    If doEntireSheetRow Then Set rng = rng.EntireRow
    '
    On Error GoTo ErrorHandler
    rng.Delete xlShiftUp
Exit Sub
ErrorHandler:
    If tbl.Parent.ProtectContents And Not tbl.Parent.ProtectionMode Then
        Err.Raise 1004, fullMethodName, "Table is on a macro protected sheet"
    Else
        Err.Raise 5, fullMethodName, Err.Description
    End If
End Sub

'*******************************************************************************
'Returns:
'   - a ListObject from a given/source Workbook by searching the provided name
'   - Nothing - it table is not found
'Does not throw errors
'Notes:
' - slower alternative: Range("'" & Workbook.Name & "'!" & tableName).ListObject
'*******************************************************************************
Public Function GetListObject(ByVal tableName As String _
                            , ByVal sourceBook As Workbook) As ListObject
    If LenB(tableName) = 0 Or sourceBook Is Nothing Then Exit Function
    '
    Dim wSheet As Worksheet
    Dim tbl As ListObject
    '
    On Error Resume Next
    For Each wSheet In sourceBook.Worksheets
        Set tbl = wSheet.ListObjects(tableName)
        If Not tbl Is Nothing Then Exit For
    Next wSheet
    On Error GoTo 0
    '
    Set GetListObject = tbl
End Function

'*******************************************************************************
'Returns a boolean indicating if a given table is filtered or not
'Does not throw errors
'*******************************************************************************
Public Function IsListObjectFiltered(targetTable As ListObject) As Boolean
    On Error Resume Next
    IsListObjectFiltered = targetTable.AutoFilter.FilterMode
    On Error GoTo 0
End Function

'*******************************************************************************
'Sorts a table only if needed
'*******************************************************************************
Public Sub SortListObjectIfNeeded(ByVal tbl As ListObject _
                                , ByVal sortColumn As Long _
                                , ByVal sortOrder As XlSortOrder)
    Const methodName As String = MODULE_NAME & ".SortListObjectIfNeeded"
    If tbl Is Nothing Then
        Err.Raise 91, methodName, "Table not set"
    ElseIf sortColumn < 1 Or sortColumn > tbl.ListColumns.Count Then
        Err.Raise 5, methodName, "Invalid sort column index"
    End If
    If tbl.ListRows.Count <= 1 Then Exit Sub 'Nothing to sort
    '
    Dim rngSort As Range:     Set rngSort = tbl.ListColumns(sortColumn).DataBodyRange
    Dim arr() As Variant:     arr = rngSort.Value2
    Dim sortAsc As Boolean:   sortAsc = (sortOrder = xlAscending)
    Dim prev As Variant:      prev = arr(LBound(arr, 1), LBound(arr, 2))
    Dim isPrevErr As Boolean: isPrevErr = IsError(prev)
    Dim curr As Variant
    Dim isCurrErr As Boolean
    Dim needsSort As Boolean
    '
    For Each curr In arr
        isCurrErr = IsError(curr)
        If isCurrErr Then
            If Not isPrevErr Then needsSort = Not sortAsc
        ElseIf isPrevErr Then
            needsSort = sortAsc
        ElseIf sortAsc Then
            needsSort = (prev > curr)
            If needsSort Then needsSort = Not IsEmpty(curr)
        Else
            needsSort = (prev < curr)
        End If
        If needsSort Then Exit For
        isPrevErr = isCurrErr
        prev = curr
    Next curr
    If Not needsSort Then Exit Sub
    '
    If IsListObjectFiltered(tbl) Then tbl.AutoFilter.ShowAllData
    With tbl.Sort.SortFields
        .Clear
        .Add rngSort, xlSortOnValues, sortOrder, , xlSortNormal
    End With
    With tbl.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
