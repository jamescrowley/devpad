VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAddInMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IDevPadMenu

Private Sub IDevPadMenu_AddMenuItem(ByVal Caption As String, ByVal Key As String, ByVal ParentIndex As Long, Optional ByVal IconIndex As Long, Optional ByVal ItemData As Long, Optional ByVal InsertItem As Boolean)
    If InsertItem Then
        frmMainForm.ctlPopMenu.InsertItem Caption, ParentIndex, Key, "", ItemData, -1 'IconIndex
    Else
        frmMainForm.ctlPopMenu.AddItem Caption, Key, "", ItemData, ParentIndex, IconIndex
    End If
End Sub

Private Property Let IDevPadMenu_Caption(ByVal Key As Variant, ByVal RHS As String)
    frmMainForm.ctlPopMenu.Caption(Key) = RHS
End Property

Private Property Get IDevPadMenu_Caption(ByVal Key As Variant) As String
    IDevPadMenu_Caption = frmMainForm.ctlPopMenu.Caption(Key)
End Property

Private Property Let IDevPadMenu_Checked(ByVal Key As Variant, ByVal RHS As Boolean)
    frmMainForm.ctlPopMenu.Checked(Key) = RHS
End Property
Private Property Get IDevPadMenu_Checked(ByVal Key As Variant) As Boolean
    IDevPadMenu_Checked = frmMainForm.ctlPopMenu.Checked(Key)
End Property

Private Property Let IDevPadMenu_Enabled(ByVal Key As Variant, ByVal RHS As Boolean)
    frmMainForm.ctlPopMenu.Enabled(Key) = RHS
End Property
Private Property Get IDevPadMenu_Enabled(ByVal Key As Variant) As Boolean
    IDevPadMenu_Enabled = frmMainForm.ctlPopMenu.Enabled(Key)
End Property

Private Property Get IDevPadMenu_ItemIndex(ByVal Key As String) As Long
    IDevPadMenu_ItemIndex = frmMainForm.ctlPopMenu.MenuIndex(Key)
End Property

Private Property Get IDevPadMenu_ItemKey(ByVal Index As Long) As String
    IDevPadMenu_ItemKey = frmMainForm.ctlPopMenu.MenuKey(Index)
End Property
Private Property Let IDevPadMenu_ItemKey(ByVal Index As Long, ByVal RHS As String)
    frmMainForm.ctlPopMenu.MenuKey(Index) = RHS
End Property

Private Sub IDevPadMenu_Remove(ByVal Key As Variant)
    frmMainForm.ctlPopMenu.RemoveItem Key
End Sub
