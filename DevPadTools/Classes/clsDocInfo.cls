VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDocInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' Developers Pad
' Version 1, BETA 2
' http://www.developerspad.com/
'
' � 1999-2000 VB Web Development
' You may not redistribute this source code,
' or distribute re-compiled versions of
' Developers Pad
'
Option Explicit

Implements IDevPadTools
Private frmForm As Form
'Private WithEvents cAddInMenus As clsAddInMenu
'
'Private Sub cAddInMenus_Click(lItem As Long)
'    MsgBox "clicked!"
'End Sub


Private Sub IDevPadTools_Connect(ByVal DevPadApp As DevPadAddInTlb.IDevPadApp)
    If cFunc Is Nothing Then Set cFunc = DevPadApp
    cFunc.Menus.AddMenuItem "Paste HTML", "MyPopPaste", cFunc.Menus.ItemIndex("EditPaste"), , , True
    'cFunc.frmMain.ctlPopMenu.InsertItem "Paste HTML", "EditPaste"
End Sub

Private Sub IDevPadTools_Disconnect()

End Sub

Private Sub IDevPadTools_HideDialog()
    If Not frmForm Is Nothing Then Unload frmForm
    Set frmForm = Nothing
End Sub

Private Function IDevPadTools_MenuClick(ByVal ItemIndex As Long) As Boolean
    
End Function

Private Sub IDevPadTools_ShowAbout()

End Sub

Private Sub IDevPadTools_ShowDialog()
    If frmForm Is Nothing Then
        Set frmForm = New frmInfo
        Load frmForm
    End If
    If frmForm.Visible = False Then
        frmForm.Show vbModeless
    Else
        frmForm.SetFocus
    End If
End Sub

