VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Developers Pad
' Version 1, BETA 2
' http://www.developerspad.com/
'
' © 1999-2000 VB Web Development
' You may not redistribute this source code,
' or distribute re-compiled versions of
' Developers Pad
'
Option Explicit
Private m_lMax As Long
Private m_sSection As String
Private m_sKey As String
Public Property Let MaxCount(lNew As Long)
    m_lMax = lNew
End Property
Public Property Let RegSection(sNew As String)
    m_sSection = sNew
End Property
Public Property Let RegKey(sNew As String)
    m_sKey = sNew
End Property
Public Property Get Items() As Collection
    Dim i As Long
    Dim sCurItem As String
    Dim cItems As Collection
    Set cItems = New Collection
    For i = 1 To m_lMax
        sCurItem = GetSetting(REG_KEY, m_sSection, m_sKey & i)
        If sCurItem = Empty Then Exit For
        cItems.Add sCurItem
    Next
    Set Items = cItems
End Property
Public Sub Add(ByVal sNewItem As String)
    ' this sub adds a new file
    ' to the list
    Dim i As Integer
    Dim sCurItem As String
    If sNewItem = "" Then Exit Sub
    If OnList(sNewItem) Then
        MoveItemToTop sNewItem
    Else
        ' Copy RecentFile1 to RecentFile2, and so on.
        For i = m_lMax - 1 To 1 Step -1
            sCurItem = GetSetting(REG_KEY, m_sSection, m_sKey & i)
            If sCurItem <> Empty Then
                'Key = sRecent & (i + 1)
                SaveSetting REG_KEY, m_sSection, m_sKey & (i + 1), sCurItem
            End If
        Next i
        ' Write the open file to first recent file.
        SaveSetting REG_KEY, m_sSection, m_sKey & "1", sNewItem
        ' delete any item left over...
        On Error Resume Next
        DeleteSetting REG_KEY, m_sSection, m_sKey & m_lMax + 1
    End If
End Sub
Public Sub Remove(Item As Variant)
    Dim lIndex As Long
    Dim i As Long
    Dim sCurItem As String
    'removes an item...
    lIndex = ItemIndex(Item)
    If lIndex <> 0 Then
        'move item 5 to 4, 4 to 3 etc...
        For i = lIndex + 1 To m_lMax Step 1
            sCurItem = GetSetting(REG_KEY, m_sSection, m_sKey & i)
            If sCurItem <> Empty Then
                ' save it to the next one down
                SaveSetting REG_KEY, m_sSection, m_sKey & i - 1, sCurItem
                DeleteSetting REG_KEY, m_sSection, m_sKey & i
            End If
        Next i
    End If
End Sub
Public Property Get ItemIndex(Item As Variant) As Long
    Dim i As Long
    If IsNumeric(Item) Then
        ItemIndex = Int(Item)
    Else
        'get the item
        For i = 1 To m_lMax
            If GetSetting(REG_KEY, m_sSection, m_sKey & i) = CStr(Item) Then
                ' Save the value
                ItemIndex = i
                Exit For
            End If
        Next
    End If
End Property
Public Property Get Item(TheItem As Variant) As String
    Dim lItemPos As Integer
    lItemPos = ItemIndex(TheItem)
    Item = GetSetting(REG_KEY, m_sSection, m_sKey & lItemPos)
End Property
Public Sub MoveItemToTop(Item As Variant)
    ' This sub moves the specified file
    ' to the top of the list
    ' from wherever it is
    Dim i As Long
    Dim sCurItem As String
    Dim lItemPos As Integer
    Dim sValue As String
    lItemPos = ItemIndex(Item)
    'get its value
    sValue = GetSetting(REG_KEY, m_sSection, m_sKey & lItemPos)
    ' Copy RecentFile1 to RecentFile2, and so on.
    For i = lItemPos - 1 To 1 Step -1
        'Key = sItem & i
        sCurItem = GetSetting(REG_KEY, m_sSection, m_sKey & i)
        If sCurItem <> Empty Then
            'Key = txtRecent & (i + 1)
            If Not i = m_lMax Then
                ' save it to the next one down
                SaveSetting REG_KEY, m_sSection, m_sKey & i + 1, sCurItem
            End If
        End If
    Next i
    'Key = txtRecent & "1"
    SaveSetting REG_KEY, m_sSection, m_sKey & "1", sValue
End Sub

Public Function OnList(ByVal sItem As String) As Boolean
    Dim i As Long
    Dim sCurItem As String
    For i = 1 To m_lMax
        sCurItem = GetSetting(REG_KEY, m_sSection, m_sKey & i)
        If sCurItem = sItem Then
            ' Save the value
            OnList = True
            Exit For
        End If
    Next
End Function
Private Sub Class_Initialize()
    m_lMax = 10
End Sub
