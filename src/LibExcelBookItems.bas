Attribute VB_Name = "LibExcelBookItems"
'''=============================================================================
''' Excel VBA Tools
''' -----------------------------------------------
''' https://github.com/cristianbuse/Excel-VBA-Tools
''' -----------------------------------------------
''' MIT License
'''
''' Copyright (c) 2018 Ion Cristian Buse
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
''    - Simple strings can be stored/retrieved using CustomXMLParts per book
''    - This module encapsulates the XML logic and exposes easy-to-use methods
''      without the need to write actual XML
'' Public/Exposed methods:
''    - BookItem - parametric property Get/Let
''    - GetBookItemNames
'' Notes:
''    - To delete a property simply set the value to a null string
''      e.g. BookItem(ThisWorkbook, "itemName") = vbNullString
'*******************************************************************************

Option Explicit
Option Private Module

Private Const XML_NAMESPACE As String = "ManagedExcelCustomXML"
Private Const rootName As String = "root"

'*******************************************************************************
'Returns the Root CustomXMLPart under the custom namespace
'part is created if missing!
'*******************************************************************************
Private Function GetRootXMLPart(ByVal book As Workbook) As CustomXMLPart
    Const xmlDeclaration As String = "<?xml version=""1.0"" encoding=""UTF-8""?>"
    Const rootTag As String = "<" & rootName & " xmlns=""" & XML_NAMESPACE _
                            & """></" & rootName & ">"
    Const rootXmlPart As String = xmlDeclaration & rootTag
    '
    With book.CustomXMLParts.SelectByNamespace(XML_NAMESPACE)
        If .Count = 0 Then
            Set GetRootXMLPart = book.CustomXMLParts.Add(rootXmlPart)
        Else
            Set GetRootXMLPart = .Item(1)
        End If
    End With
End Function

'*******************************************************************************
'Clears all CustomXMLParts under the custom namespace
'*******************************************************************************
Private Sub ClearRootXMLParts(ByVal book As Workbook)
    With book.CustomXMLParts.SelectByNamespace(XML_NAMESPACE)
        Dim i As Long
        For i = .Count To 1 Step -1
            .Item(i).Delete
        Next i
    End With
End Sub

'*******************************************************************************
'Get the Root Node under the custom namespace
'Node is created if missing!
'*******************************************************************************
Private Function GetRootNode(ByVal book As Workbook) As CustomXMLNode
    Dim root As CustomXMLNode
    If root Is Nothing Then
        With GetRootXMLPart(book)
            Dim nsPrefix As String
            nsPrefix = .NamespaceManager.LookupPrefix(XML_NAMESPACE)
            Set root = .SelectSingleNode("/" & nsPrefix & ":" & rootName & "[1]")
        End With
    End If
    Set GetRootNode = root
End Function

'*******************************************************************************
'Get an XML Node. Create it if missing
'*******************************************************************************
Private Function GetNode(ByVal book As Workbook _
                       , ByVal nodeName As String _
                       , ByVal addIfMIssing As Boolean) As CustomXMLNode
    Dim node As CustomXMLNode
    Dim expr As String
    '
    With GetRootNode(book)
        expr = .XPath & "/" & nodeName & "[1]"
        Set node = .SelectSingleNode(expr)
        If node Is Nothing And addIfMIssing Then
            .AppendChildNode nodeName
            Set node = .SelectSingleNode(expr)
        End If
    End With
    Set GetNode = node
End Function

'*******************************************************************************
'Retrieves/sets a book property value from a CustomXMLNode
'*******************************************************************************
Public Property Get BookItem(ByVal book As Workbook _
                           , ByVal itemName As String) As String
    ThrowIfInvalid book, itemName
    Dim node As CustomXMLNode
    Set node = GetNode(book, itemName, False)
    If Not node Is Nothing Then BookItem = node.Text
End Property
Public Property Let BookItem(ByVal book As Workbook _
                           , ByVal itemName As String _
                           , ByVal itemValue As String)
    ThrowIfInvalid book, itemName
    If LenB(itemValue) = 0 Then
        Dim node As CustomXMLNode
        Set node = GetNode(book, itemName, False)
        If Not node Is Nothing Then node.Delete
    Else
        GetNode(book, itemName, True).Text = itemValue
    End If
End Property
Private Sub ThrowIfInvalid(ByRef book As Workbook, ByRef itemName As String)
    Const methodName As String = "BookItem"
    If book Is Nothing Then
        Err.Raise 91, methodName, "Book not set"
    ElseIf LenB(itemName) = 0 Then
        Err.Raise 5, methodName, "Invalid item name"
    End If
End Sub

'*******************************************************************************
'Returns a collection of custom node names within the custom namespace
'*******************************************************************************
Public Function GetBookItemNames(ByVal book As Workbook) As Collection
    If book Is Nothing Then Err.Raise 91, "GetBookItemNames", "Book not set"
    '
    Dim coll As New Collection
    With GetRootNode(book).ChildNodes
        Dim i As Long
        ReDim arr(0 To .Count - 1)
        For i = 1 To .Count
            coll.Add .Item(i).BaseName
        Next i
    End With
    Set GetBookItemNames = coll
End Function
