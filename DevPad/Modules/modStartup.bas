Attribute VB_Name = "modStartup"
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

Public Const SC_RESTORE = &HF120&
Public Const mcTHISAPPID = "DevPad"
Private Const WM_SYSCOMMAND = &H112

Public m_hMutex As Long
Public m_hWndPrevious As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
'Private cRegister As New clsRegister
Private Sub Main()
    Dim sCode As String
    Load frmSplash
    frmSplash.Main
    Unload frmSplash
End Sub

Private Function IsThisApp(ByVal hWnd As Long) As Boolean
   ' Check if the windows property is set for this
   ' window handle:
   If GetProp(hWnd, mcTHISAPPID & "_APPLICATION") = 1 Then
      IsThisApp = True
   End If
End Function
Public Sub ParseCommand(ByVal sCmd As String)
On Error GoTo ErrHandler

    Dim sFile As String
    Dim sFiles() As String
    Dim i As Long
    ' Bring me to the foreground and restore
    ' if iconic:
    If Len(sCmd) > 0 Then
        RestoreAndActivate frmMainForm.hWnd
        ' Your function to parse the command line:
        sCmd = Trim$(sCmd)
        'remove the quotes
        sFiles = Split(sCmd, """ """)
        For i = 0 To UBound(sFiles)
            'MsgBox StripChar("""", sFiles(i))
            LoadFileDefault StripChar("""", sFiles(i))
        Next
        frmMainForm.SetDocStatus
    Else
        'no command line... do default
        If GetSetting(REG_KEY, "Settings", "RestoreWorkspace", "0") = "1" And GetSetting(REG_KEY, "Settings", "LastWorkspace", "") <> "" Then
            'load the last workspace
            cWorkspace.Load GetSetting(REG_KEY, "Settings", "LastWorkspace", "")
        ElseIf GetSetting(REG_KEY, "Settings", "SU_LoadDocument", "1") = "1" Then
            cDocuments.New
        Else
            frmMainForm.SetDocStatus
        End If
    End If
    
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Startup.ParseCommand" ', sCmd
End Sub
Public Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    If IsThisApp(hWnd) Then
        EnumWindowsProc = 0
        m_hWndPrevious = hWnd
    Else
        EnumWindowsProc = 1
    End If
End Function
Public Sub RestoreAndActivate(ByVal hWnd As Long)
   If (IsIconic(hWnd)) Then
      SendMessageLong hWnd, WM_SYSCOMMAND, SC_RESTORE, 0
   End If
   ActivateWindow hWnd
End Sub
Public Sub ActivateWindow(ByVal lhWnd As Long)
    SetForegroundWindow lhWnd
End Sub
Public Function EndApp()
   ' Call this to remove the Mutex.  It will be cleared
   ' anyway by windows, but this ensures it works.
   If (m_hMutex <> 0) Then
      CloseHandle m_hMutex
   End If
   m_hMutex = 0
End Function
Public Sub TagWindow(ByVal hWnd As Long)
   ' Applies a window property to allow the window to
   ' be clearly identified.
   SetProp hWnd, mcTHISAPPID & "_APPLICATION", 1
End Sub
