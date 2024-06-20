VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BookConnection"
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

'*******************************************************************************
'' Description:
''  - Opens a workbook, if needed (if not already open)
''  - Stores the open state so that it can be closed later if needed
'' Notes:
''  - 'Disconnect' is called when the instance is terminated
'*******************************************************************************

Option Explicit

Private m_book As Workbook
Private m_alreadyOpened As Boolean

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Class Events
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Class_Terminate()
    Disconnect
End Sub

'*******************************************************************************
'Open a Workbook by file path (if not already opened)
'*******************************************************************************
Public Function Connect(ByVal filePath As String _
                      , Optional ByVal updateLinks As Boolean = False _
                      , Optional ByVal readOnly As Boolean = True _
                      , Optional ByVal password As String _
                      , Optional ByVal minimizeNewlyOpened As Boolean = True _
) As Boolean
    Disconnect 'Just in case 'Connect' was called before
    '
    Dim bookCount As Long: bookCount = Workbooks.Count
    Dim wnd As Window
    '
    On Error Resume Next
    Set m_book = Workbooks.Open(fileName:=filePath _
                              , updateLinks:=updateLinks _
                              , readOnly:=readOnly _
                              , password:=password)
    If Not m_book Is Nothing Then
        m_alreadyOpened = (bookCount = Workbooks.Count)
        If Not m_alreadyOpened And minimizeNewlyOpened Then
            For Each wnd In m_book.Windows
                wnd.WindowState = xlMinimized
            Next wnd
        End If
        Connect = True
    End If
    On Error GoTo 0
End Function

'*******************************************************************************
'Close the Workbook (if not already opened before 'Connect')
'*******************************************************************************
Public Sub Disconnect(Optional ByVal saveChanges As Boolean = False)
    If m_book Is Nothing Then
        m_alreadyOpened = False
        Exit Sub
    End If
    If Not m_alreadyOpened Then
        On Error Resume Next
        m_book.Close saveChanges
        On Error GoTo 0
    End If
    '
    Set m_book = Nothing
    m_alreadyOpened = False
End Sub

'*******************************************************************************
'Get self instance. Useful in 'With New ...' blocks
'*******************************************************************************
Public Function Self() As BookConnection
    Set Self = Me
End Function

'===============================================================================
'Returns the 'connected' book
'===============================================================================
Public Property Get Workbook() As Workbook
    Set Workbook = m_book
End Property
