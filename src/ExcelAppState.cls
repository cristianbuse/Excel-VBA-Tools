VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelAppState"
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
''' Copyright (c) 2019 Ion Cristian Buse
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
''
''==============================================================================
'' Description:
''  - Class used to store/modify/restore Excel Application properties. Can be
''    useful is minimizing execution time for code interacting with the app
'' Example usage:
''    Sub Test()
''        Dim app As New ExcelAppState: app.StoreState: app.Sleep
''        'Do Whatever
''        app.RestoreState 'Optional because it's triggered from Class_Terminate
''    End Sub
''==============================================================================

Option Explicit

Private m_eventOn As Boolean
Private m_calcMode As XlCalculation
Private m_repaintOn As Boolean
Private m_alertsOn As Boolean
Private m_printOn As Boolean

Private m_hasStoredState As Boolean
Private m_hasStoredCalcMode As Boolean

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Safety restore
'When this instance reference count reaches 0 then restore state only if needed
'Useful in cases where an error occured and the state was not restored properly
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Class_Terminate()
    If m_hasStoredState Then RestoreState
End Sub

'*******************************************************************************
'Retain the current application settings
'*******************************************************************************
Public Sub StoreState()
    With Application
        m_eventOn = .EnableEvents
        m_repaintOn = .ScreenUpdating
        m_alertsOn = .DisplayAlerts
        m_printOn = .PrintCommunication
        '
        On Error Resume Next 'In case no Workbook is open (running from AddIn)
        m_calcMode = .Calculation
        m_hasStoredCalcMode = (Err.Number = 0)
        On Error GoTo 0
    End With
    m_hasStoredState = True
End Sub

'*******************************************************************************
'Restore the previously saved application settings
'*******************************************************************************
Public Sub RestoreState(Optional ByVal maxMilliSecondsToWait As Long = 0)
    If Not m_hasStoredState Then
        Err.Raise 5, TypeName(Me) & ".RestoreState", "Nothing to restore"
    End If
    With Application
        If .EnableEvents Xor m_eventOn Then .EnableEvents = m_eventOn
        If .ScreenUpdating Xor m_repaintOn Then .ScreenUpdating = m_repaintOn
        If m_hasStoredCalcMode Then
            On Error Resume Next
            If .Calculation <> m_calcMode Then .Calculation = m_calcMode
            If Err.Number = 0 Then WaitForCalculations maxMilliSecondsToWait
            On Error GoTo 0
        End If
        If .DisplayAlerts Xor m_alertsOn Then .DisplayAlerts = m_alertsOn
        If .PrintCommunication Xor m_printOn Then .PrintCommunication = m_printOn
    End With
    m_hasStoredState = False
End Sub

'*******************************************************************************
'Clear the previously saved application settings
'*******************************************************************************
Public Sub ClearStateIfNeeded()
    m_hasStoredState = False
End Sub

'*******************************************************************************
'Put App to sleep. Speeds up code execution for certain interactions with app
'*******************************************************************************
Public Sub Sleep(Optional ByVal eventsOn As Boolean = False _
               , Optional ByVal calcMode As XlCalculation = xlCalculationManual _
               , Optional ByVal repaintOn As Boolean = False _
               , Optional ByVal alertsOn As Boolean = False _
               , Optional ByVal printOn As Boolean = False)
    Select Case calcMode
    Case xlCalculationAutomatic, xlCalculationSemiautomatic, xlCalculationManual
    Case Else
        calcMode = xlCalculationManual
    End Select
    With Application
        If .EnableEvents Xor eventsOn Then .EnableEvents = eventsOn
        On Error Resume Next
        If .Calculation <> calcMode Then .Calculation = calcMode
        On Error GoTo 0
        If .ScreenUpdating Xor repaintOn Then .ScreenUpdating = repaintOn
        If .DisplayAlerts Xor alertsOn Then .DisplayAlerts = alertsOn
        If .PrintCommunication Xor printOn Then .PrintCommunication = printOn
    End With
End Sub

'*******************************************************************************
'Set Application to Normal working mode
'*******************************************************************************
Public Sub Wake(Optional ByVal maxMilliSecondsToWait As Long = 0 _
              , Optional ByVal eventsOn As Boolean = True _
              , Optional ByVal calcMode As XlCalculation = xlCalculationAutomatic _
              , Optional ByVal repaintOn As Boolean = True _
              , Optional ByVal alertsOn As Boolean = True _
              , Optional ByVal printOn As Boolean = True)
    Select Case calcMode
    Case xlCalculationAutomatic, xlCalculationSemiautomatic, xlCalculationManual
    Case Else
        calcMode = xlCalculationAutomatic
    End Select
    With Application
        If .EnableEvents Xor eventsOn Then .EnableEvents = eventsOn
        If .ScreenUpdating Xor repaintOn Then .ScreenUpdating = repaintOn
        On Error Resume Next
        If .Calculation <> calcMode Then .Calculation = calcMode
        If Err.Number = 0 Then WaitForCalculations maxMilliSecondsToWait
        On Error GoTo 0
        If .DisplayAlerts Xor alertsOn Then .DisplayAlerts = alertsOn
        If .PrintCommunication Xor printOn Then .PrintCommunication = printOn
    End With
End Sub

'*******************************************************************************
'Wait for the application to finish calculations
'*******************************************************************************
Public Sub WaitForCalculations(ByVal maxMilliSecondsToWait As Long)
    If Application.Calculation = xlCalculationManual Then Exit Sub
    If maxMilliSecondsToWait <= 0 Then Exit Sub
    If UDFMode() Then Exit Sub 'Code running from a UDF in a range/named range
    '
    Const msPerDay As Long = 24& * 60& * 60& * 1000&
    Dim remainingTriesToFixBug As Long: remainingTriesToFixBug = 3
    Dim tMax As Date: tMax = NowMs() + maxMilliSecondsToWait / msPerDay
    '
    Do While Application.CalculationState <> xlDone
        DoEvents
        If NowMs() > tMax Then Exit Do
        If Application.CalculationState = xlPending Then
            If remainingTriesToFixBug = 0 Then Exit Do
            remainingTriesToFixBug = remainingTriesToFixBug - 1
            TryFixingPendingBug
        End If
    Loop
End Sub

'*******************************************************************************
'Utility function for 'WaitForCalculations' method
'*******************************************************************************
Private Sub TryFixingPendingBug()
    'There is a bug in Excel when a Volatile function is present: even if
    '   the Status Bar shows 'Ready', Application.CalculationState could
    '   still return 'xlPending'
    'Force state out of pending by marking any formula cell as Dirty
    Dim formulaCell As Range
    '
    Set formulaCell = GetEditableFormulaCell()
    If Not formulaCell Is Nothing Then formulaCell.Dirty
End Sub

'*******************************************************************************
'Utility function for 'TryFixingPendingBug' method
'*******************************************************************************
Private Function GetEditableFormulaCell() As Range
    Dim wSheet As Worksheet
    Dim collMacroProtected As New Collection
    Dim formulaCell As Range
    '
    On Error Resume Next
    For Each wSheet In ThisWorkbook.Worksheets
        If wSheet.EnableCalculation Then
            If wSheet.ProtectContents And Not wSheet.ProtectionMode Then
                collMacroProtected.Add wSheet
            Else
                Set formulaCell = wSheet.UsedRange _
                                 .SpecialCells(xlCellTypeFormulas).Cells(1, 1)
                If Not formulaCell Is Nothing Then Exit For
            End If
        End If
    Next wSheet
    On Error GoTo 0
    '
    If formulaCell Is Nothing Then
        For Each wSheet In collMacroProtected
            Set formulaCell = GetUnlockedFormulaCell(wSheet)
            If Not formulaCell Is Nothing Then Exit For
        Next wSheet
    End If
    '
    Set GetEditableFormulaCell = formulaCell
End Function

'*******************************************************************************
'Utility function for 'GetEditableFormulaCell' method
'*******************************************************************************
Private Function GetUnlockedFormulaCell(ByVal wSheet As Worksheet) As Range
    Dim rngFormulas As Range
    Dim rngTemp As Range
    Dim isLocked As Variant
    Dim tempCell As Range
    '
    On Error Resume Next
    Set rngFormulas = wSheet.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If rngFormulas Is Nothing Then Exit Function
    '
    For Each rngTemp In rngFormulas.Columns
        isLocked = rngTemp.Locked
        If IsNull(isLocked) Then isLocked = False
        If Not isLocked Then
            For Each tempCell In rngTemp
                If Not tempCell.Locked Then
                    Set GetUnlockedFormulaCell = tempCell
                    Exit Function
                End If
            Next tempCell
        End If
    Next rngTemp
End Function

'*******************************************************************************
'An enhanced 'Now' - returns the date and time including milliseconds
'On Mac resolution depends on Excel version and can be 10ms or less (e.g. 4ms)
'On Win resolution is around 4ms
'*******************************************************************************
Private Function NowMs() As Date
    Const secondsPerDay As Long = 24& * 60& * 60&
#If Mac Then
    Const evalResolution As Double = 0.01
    Const evalFunc As String = "=Now()"
    Static useEval As Boolean
    Static isSet As Boolean
    '
    If Not isSet Then
        useEval = TimerResolution() > evalResolution
        isSet = True
    End If
    If useEval Then
        NowMs = Evaluate(evalFunc)
        Exit Function
    End If
#End If
    NowMs = Date + Round(Timer, 3) / secondsPerDay
End Function
#If Mac Then
Private Function TimerResolution() As Double
    Const secondsPerDay As Long = 24& * 60& * 60&
    Static r As Double
    If r = 0 Then
        Dim t As Double: t = Timer
        Do
            r = Round(Timer - t, 3)
            If r < 0# Then r = r + secondsPerDay
        Loop Until r > 0#
    End If
    TimerResolution = r
End Function
#End If

'*******************************************************************************
'Returns a boolean indicating if code was called from a UDF
'*******************************************************************************
Private Function UDFMode() As Boolean
    Dim dispAlerts As Boolean: dispAlerts = Application.DisplayAlerts
    '
    On Error Resume Next
    Application.DisplayAlerts = Not dispAlerts 'Cannot be changed in UDF mode
    On Error GoTo 0
    '
    UDFMode = (Application.DisplayAlerts = dispAlerts)
    If Not UDFMode Then Application.DisplayAlerts = dispAlerts 'Revert
End Function
