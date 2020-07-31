VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScreenShot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDevPadTools
Private frmForm As Form

Private Sub Class_Terminate()
    IDevPadTools_HideDialog
End Sub

Private Sub IDevPadTools_Connect(ByVal AddInFuncs As Object, ByVal AddInMenus As Object)
If cFunc Is Nothing Then Set cFunc = AddInFuncs
End Sub
Private Sub IDevPadTools_HideDialog()
    If Not frmForm Is Nothing Then Unload frmForm
    Set frmForm = Nothing
End Sub

Private Sub IDevPadTools_ShowAbout()

End Sub

Private Sub IDevPadTools_ShowDialog()
    If frmForm Is Nothing Then
        Set frmForm = New frmScreenShot
        Load frmForm
    End If
    frmForm.Show vbModeless, cFunc.frmMain
End Sub
Private Sub IDevPadTools_Disconnect()

End Sub
