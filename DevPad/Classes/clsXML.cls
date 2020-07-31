VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Developers Pad Implementation of an XML parser

'Far from efficient, and I don't know if it works
'for every XML file... but it does what we need it
'to do, without having to carry a big 3rd party DLL
'around with us!

Private m_cNodes()  As Node
Private m_lNodeCount As Long
Public Property Get Node(sName As String, Optional lInstance As Long = 1) As Node
Dim i As Long
Dim lCount As Long
    For i = 1 To m_lNodeCount
        If m_cNodes(i).Name = sName Then
            lCount = lCount + 1
            If lInstance = lCount Then
                LSet Node = m_cNodes(i)
                Exit For
            End If
        End If
    Next
End Property
Public Property Get NodeChild(NodeIndex As Long, sChildName As String, Optional lInstance As Long = 1) As Node
Dim i As Long
Dim lCount As Long
    For i = 1 To m_lNodeCount
        If m_cNodes(i).ParentNode = NodeIndex And m_cNodes(i).Name = sChildName Then
            lCount = lCount + 1
            If lInstance = lCount Then
                LSet NodeChild = m_cNodes(i)
                Exit For
            End If
        End If
    Next
End Property
Public Property Get NodeChildIndex(NodeIndex As Long, sChildName As String, Optional lInstance As Long = 1) As Long
Dim i As Long
Dim lCount As Long
    For i = 1 To m_lNodeCount
        If m_cNodes(i).ParentNode = NodeIndex And m_cNodes(i).Name = sChildName Then
            lCount = lCount + 1
            If lInstance = lCount Then
                NodeChildIndex = i
                Exit For
            End If
        End If
    Next
End Property
Public Property Get NodeChildCount(NodeIndex As Long, sChildName As String) As Long
Dim i As Long

    For i = 1 To m_lNodeCount
        If m_cNodes(i).ParentNode = NodeIndex And m_cNodes(i).Name = sChildName Then
            NodeChildCount = NodeChildCount + 1
        End If
    Next
End Property
Public Property Get NodeAttribute(NodeIndex As Long, sName As String) As String
Dim i As Long
    For i = 1 To m_cNodes(NodeIndex).AttributeCount
        If m_cNodes(NodeIndex).Attributes(i).Name = sName Then
            NodeAttribute = m_cNodes(NodeIndex).Attributes(i).Value
            Exit For
        End If
    Next
End Property
Public Sub ParseXMLFile(sFile As String)
    Dim iFileNum    As Integer
    Dim sBuf        As String
    Dim lLen        As Long
    Dim i           As Long
    Dim j           As Long
    Dim bInNode     As Boolean
    Dim lParentNode() As Long
    Dim lNextSpace  As Long
    Dim lNextTag    As Long
    Dim sName       As String
    Dim sValue      As String
    Dim lLevel      As Long
    Dim sChar       As String
    Dim lPos        As Long
    Dim lLastPos    As Long
    Dim sDelimiter(3) As String
    Dim lStartPos   As Long
    Dim lValuePos   As Long
    'clear array
    Erase m_cNodes
    m_lNodeCount = 0
    'init array
    ReDim lParentNode(0 To 0)
    'init delimiters
    sDelimiter(0) = " "
    sDelimiter(1) = ">"
    sDelimiter(2) = "/>"
    sDelimiter(3) = vbCrLf
    
    'get a file num
    iFileNum = FreeFile
    'open the file
    Open sFile For Binary Access Read Lock Write As #iFileNum
    'read the file into the buffer
    lLen = LOF(iFileNum)
    sBuf = Space$(lLen)
    Get #iFileNum, , sBuf
    'close the file
    Close #iFileNum
    
    'read the file into a buffer
    For i = 1 To lLen
        'get the current char
        sChar = Mid$(sBuf, i, 1)
        Select Case sChar
        Case "<"
            If bInNode = False Then
                If Mid$(sBuf, i + 1, 1) = "/" Then
                    'end of node...
                    If lLevel <> 0 Then
                        If m_cNodes(lParentNode(lLevel)).HasChildren = False Then
                            If lValuePos <> 0 Then
                                'get the nodes stored value
                                sValue = Mid$(sBuf, lValuePos, i - lValuePos)
                                'remove spaces here...
                                sValue = Trim$(sValue)
                                'remove vbCrLf at either end, if needed
                                If Left$(sValue, 2) = vbCrLf Then sValue = Right$(sValue, Len(sValue) - 2)
                                If Right$(sValue, 2) = vbCrLf Then sValue = Left$(sValue, Len(sValue) - 2)
                                'and here...
                                sValue = Trim$(sValue)
                                
                                m_cNodes(lParentNode(lLevel)).Value = sValue
                            End If
                        End If
                    End If
                    'go up a level...
                    lLevel = lLevel - 1
                    'skip to end of tag
                    lNextTag = InStr(i + 1, sBuf, ">")
                    i = lNextTag
                ElseIf Mid$(sBuf, i + 1, 1) = "!" Then
                    'skip... a comment
                    'skip to end of tag
                    lNextTag = InStr(i + 1, sBuf, ">")
                    i = lNextTag
                Else
                    'new node
                    bInNode = True
                    lLastPos = lLen
                    'get the name of the node
                    For j = 0 To UBound(sDelimiter)
                        lPos = InStr(i + 1, sBuf, sDelimiter(j))
                        If lPos < lLastPos And lPos <> 0 Then lLastPos = lPos
                    Next j
                    sName = Mid$(sBuf, i + 1, lLastPos - (i + 1))
                    'add node to array
                    m_lNodeCount = m_lNodeCount + 1
                    ReDim Preserve m_cNodes(1 To m_lNodeCount)
                    m_cNodes(m_lNodeCount).Name = sName
                    m_cNodes(m_lNodeCount).ParentNode = lParentNode(lLevel)
                    m_cNodes(m_lNodeCount).Index = m_lNodeCount
                    If lLevel <> 0 Then m_cNodes(lParentNode(lLevel)).HasChildren = True
                    i = i + Len(sName)
                End If
                lValuePos = 0
            End If
        Case "="
            If bInNode Then
                'add property to current node
                'get the property name
                lStartPos = i - 1
                If Mid$(sBuf, i - 1, 1) = " " Then lStartPos = lStartPos - 1
                
                lLastPos = 0
                'get the name of the node
                For j = 0 To UBound(sDelimiter)
                    lPos = InStrRev(Left$(sBuf, lStartPos), sDelimiter(j))
                    If lPos > lLastPos Then lLastPos = lPos
                Next j
                'get the attribute name...
                sName = Mid$(sBuf, lLastPos + 1, lStartPos - lLastPos)
                
                lStartPos = i + 1
                If Mid$(sBuf, i + 1, 1) = " " Then lStartPos = lStartPos + 1
                'temp val
                sValue = Mid$(sBuf, lStartPos, 1)
                If sValue = """" Or sValue = "'" Then
                    'value is within quotes
                    'find the end of the quote
                    lPos = InStr(lStartPos + 1, sBuf, sValue)
                    If lPos = 0 Then lPos = lLen
                    sValue = Mid$(sBuf, lStartPos + 1, lPos - (lStartPos + 1))
                Else
                    'get the attribute value
                    lLastPos = lLen
                    'get the name of the node
                    For j = 0 To UBound(sDelimiter)
                        lPos = InStr(lStartPos, sBuf, sDelimiter(j))
                        If lPos < lLastPos And lPos <> 0 Then lLastPos = lPos
                    Next j
                    sValue = Mid$(sBuf, lStartPos, lLastPos - (lStartPos))
                End If
                'add to attributes list
                DoEvents
                With m_cNodes(m_lNodeCount)
                    .AttributeCount = .AttributeCount + 1
                    ReDim Preserve .Attributes(1 To .AttributeCount)
                    .Attributes(.AttributeCount).Name = sName
                    .Attributes(.AttributeCount).Value = sValue
                End With
            End If
        Case ">"
            If bInNode Then
                lValuePos = 0
                bInNode = False
                If Mid$(sBuf, i - 1, 1) = "/" Then
                    'end of the node...
                Else
                    'save the current node
                    lLevel = lLevel + 1
                    ReDim Preserve lParentNode(0 To lLevel)
                    lParentNode(lLevel) = m_lNodeCount
                    lValuePos = i + 1
                End If
            End If
        End Select
    Next i
End Sub
