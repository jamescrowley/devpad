Attribute VB_Name = "modDialogHook"
Option Explicit

' ==========================================================================
' Module:   modDialogHook
' Filename: modDialogHook.bas
' Author:   Steve McMahon, Edited by James Crowley
' Date:     04 December 2000
'
' Provides functions which can be called via AddressOf for common
' dialog hook support.
' ==========================================================================

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private m_cHookedDialog As Long

Property Let HookedDialog(ByRef cThis As GCommonDialog)
    'Set cHookedDialog = cThis
    m_cHookedDialog = ObjPtr(cThis)
End Property
Property Get HookedDialog() As GCommonDialog
Dim oT As GCommonDialog
    If (m_cHookedDialog <> 0) Then
        ' Turn the pointer into an illegal, uncounted interface
        CopyMemory oT, m_cHookedDialog, 4
        ' Do NOT hit the End button here! You will crash!
        ' Assign to legal reference
        Set HookedDialog = oT
        ' Still do NOT hit the End button here! You will still crash!
        ' Destroy the illegal reference
        CopyMemory oT, 0&, 4
    End If
End Property
Public Sub ClearHookedDialog()
    m_cHookedDialog = 0
End Sub
Public Function DialogHookFunction(ByVal hDlg As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim cD As GCommonDialog
    Set cD = HookedDialog
    If Not (cD Is Nothing) Then
        DialogHookFunction = cD.DialogHook(hDlg, msg, wParam, lParam)
    End If
End Function
