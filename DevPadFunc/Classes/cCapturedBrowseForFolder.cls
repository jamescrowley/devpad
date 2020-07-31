VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCaptureBF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

'Private Type RECT
'   left As Long
'   top As Long
'   right As Long
'   bottom As Long
'End Type

Private WithEvents m_cBF As clsBrowseForFolder
Attribute m_cBF.VB_VarHelpID = -1
Private m_iWp As ICaptureBF

'Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Const WM_SYSCOMMAND = &H112&
Private Const SC_CLOSE = &HF060&
'Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_CLOSE = &H10
'Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


'
'Private Const GWL_STYLE = (-16)
'Private Const GWL_EXSTYLE = (-20)
'Private Const WS_OVERLAPPED = &H0&
'Private Const WS_POPUP = &H80000000
'Private Const WS_CHILD = &H40000000
'Private Const WS_MINIMIZE = &H20000000
'Private Const WS_VISIBLE = &H10000000
'Private Const WS_DISABLED = &H8000000
'Private Const WS_CLIPSIBLINGS = &H4000000
'Private Const WS_CLIPCHILDREN = &H2000000
'Private Const WS_MAXIMIZE = &H1000000
'Private Const WS_CAPTION = &HC00000                 '/* WS_BORDER | WS_DLGFRAME  */
'Private Const WS_BORDER = &H800000
'Private Const WS_DLGFRAME = &H400000
'Private Const WS_VSCROLL = &H200000
'Private Const WS_HSCROLL = &H100000
'Private Const WS_SYSMENU = &H80000
'Private Const WS_THICKFRAME = &H40000
'Private Const WS_GROUP = &H20000
'Private Const WS_TABSTOP = &H10000
'
'Private Const WS_MINIMIZEBOX = &H20000
'Private Const WS_MAXIMIZEBOX = &H10000
'
'Private Const WS_TILED = WS_OVERLAPPED
'Private Const WS_ICONIC = WS_MINIMIZE
'Private Const WS_SIZEBOX = WS_THICKFRAME
'
''/*
'' * Extended Window Styles
'' */
'Private Const WS_EX_DLGMODALFRAME = &H1&
'Private Const WS_EX_NOPARENTNOTIFY = &H4&
'Private Const WS_EX_TOPMOST = &H8&
'Private Const WS_EX_ACCEPTFILES = &H10&
'Private Const WS_EX_TRANSPARENT = &H20&
'Private Const WS_EX_MDICHILD = &H40&
'Private Const WS_EX_TOOLWINDOW = &H80&
'Private Const WS_EX_WINDOWEDGE = &H100&
'Private Const WS_EX_CLIENTEDGE = &H200&
'Private Const WS_EX_CONTEXTHELP = &H400&
'
'Private Const WS_EX_RIGHT = &H1000&
'Private Const WS_EX_LEFT = &H0&
'Private Const WS_EX_RTLREADING = &H2000&
'Private Const WS_EX_LTRREADING = &H0&
'Private Const WS_EX_LEFTSCROLLBAR = &H4000&
'Private Const WS_EX_RIGHTSCROLLBAR = &H0&
'
'Private Const WS_EX_CONTROLPARENT = &H10000
'Private Const WS_EX_STATICEDGE = &H20000
'Private Const WS_EX_APPWINDOW = &H40000
'
'Private Const WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
'Private Const WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Private m_bDontUnload As Boolean

' Modified by VB Web
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_HIDE = 0


Public Sub Reload(ByVal sSelPath As String)
Attribute Reload.VB_Description = "Closes the open Browse for Folder dialog and opens a new one.  Required if you make or delete any folder when the dialog is open."
Dim hwnd As Long
   m_bDontUnload = True
   hwnd = m_cBF.DialoghWnd
   pUnloadBrowseDialog hwnd
   m_bDontUnload = False
   m_cBF.InitialDir = sSelPath
   Show m_iWp
End Sub

Public Property Get Browse() As clsBrowseForFolder
Attribute Browse.VB_Description = "Returns the cBrowseForFolder object.  Use this to set up the options, initial folder and so on."
   Set Browse = m_cBF
End Property
Public Sub Show(ByRef iwp As ICaptureBF)

Dim sR As String
   Set m_iWp = iwp
   sR = m_cBF.BrowseForFolder
   If Not (m_bDontUnload) Then
      If Not m_iWp Is Nothing Then
         m_iWp.Unload
      End If
   End If
End Sub

Private Sub Class_Initialize()
   'DebugMsg "cCaptureBrowseForFolder:Initialize"
   Set m_cBF = New clsBrowseForFolder
End Sub

Private Sub Class_Terminate()
   Set m_cBF = Nothing
   Set m_iWp = Nothing
   'DebugMsg "cCaptureBrowseForFolder:Terminate"
End Sub

Private Sub m_cBF_Initialized()
Dim tR As RECT
Dim lhWndC As Long ' Capture PicBox hWnd
Dim lhWndT As Long ' Dialog hWnd
Dim lStyle As Long
Dim lhWndTV As Long ' TreeView hWnd
Dim hButtonOK As Long
Dim hButtonCancel As Long
   lhWndC = m_iWp.CapturehWnd
   GetClientRect lhWndC, tR
   lhWndT = m_cBF.DialoghWnd
   lStyle = GetWindowLong(lhWndT, GWL_STYLE)
   lStyle = lStyle And Not (WS_BORDER Or WS_DLGFRAME Or WS_CAPTION Or WS_BORDER Or WS_SIZEBOX Or WS_THICKFRAME)
   'lStyle = lStyle And Not (WS_BORDER Or WS_CAPTION Or WS_SIZEBOX)
   lStyle = lStyle Or WS_CHILD
   SetWindowLong lhWndT, GWL_STYLE, lStyle
   lStyle = GetWindowLong(lhWndT, GWL_EXSTYLE)
   lStyle = lStyle And Not WS_EX_DLGMODALFRAME
   SetWindowLong lhWndT, GWL_EXSTYLE, lStyle
   
   
   
   ' Modified by VB Web
   hButtonOK = FindWindowEx(lhWndT, 0&, "Button", "OK")
   hButtonCancel = FindWindowEx(lhWndT, 0&, "Button", "Cancel")
   ShowWindow hButtonOK, SW_HIDE
   ShowWindow hButtonCancel, SW_HIDE
   'setparent lhwndc,
   SetParent lhWndT, lhWndC
   
   MoveWindow lhWndT, 0, 0, tR.Right - tR.Left, tR.Bottom - tR.Top, 1
   
   lhWndTV = FindWindowEx(lhWndT, 0, "SysTreeView32", "")
   If lhWndTV <> 0 Then
      MoveWindow lhWndTV, 0, 0, tR.Right - tR.Left, tR.Bottom - tR.Top, 1
   End If
   
   lStyle = GetWindowLong(lhWndC, GWL_EXSTYLE)
   lStyle = lStyle Or WS_EX_CONTROLPARENT
   SetWindowLong lhWndC, GWL_EXSTYLE, lStyle
   
   SetThin3DBorder lhWndTV
 '  SetThin3DBorder lhWndC
   m_iWp.CaptureBrowseForFolder = Me
End Sub
Public Sub ResizeDlg()
    Dim tR As RECT
    Dim lhWndC As Long ' Capture PicBox hWnd
    Dim lhWndT As Long ' Dialog hWnd
    Dim lhWndTV As Long ' TreeView hWnd
    lhWndC = m_iWp.CapturehWnd
    GetClientRect lhWndC, tR
    lhWndT = m_cBF.DialoghWnd
    MoveWindow lhWndT, 0, 0, tR.Right - tR.Left, tR.Bottom - tR.Top, 1
    lhWndTV = FindWindowEx(lhWndT, 0, "SysTreeView32", "")
    If lhWndTV <> 0 Then
       MoveWindow lhWndTV, 0, 0, tR.Right - tR.Left, tR.Bottom - tR.Top, 1
    End If
End Sub

Public Sub Unload()
Attribute Unload.VB_Description = "Raised when the browse for folder dialog is unloaded.  Call Unload Me in your form in response to this event."
Dim hwnd As Long
   
   On Error Resume Next
   hwnd = m_cBF.DialoghWnd
   If Err.Number = 0 Then
      If hwnd <> 0 Then
         pUnloadBrowseDialog hwnd
      End If
   End If
   m_iWp.Unload
   Set m_iWp = Nothing
End Sub
Private Sub pUnloadBrowseDialog(ByVal hwnd As Long)
Dim lR As Long
    ShowWindow hwnd, SW_HIDE
    SetParent hwnd, 0
    lR = SendMessageLong(hwnd, WM_SYSCOMMAND, SC_CLOSE, 0)
    DestroyWindow hwnd
    Debug.Assert (IsWindow(hwnd) = 0)
End Sub

Private Sub m_cBF_SelectionChanged(ByVal sPath As String, bAllowOk As Boolean)
   m_iWp.SelectionChanged sPath
End Sub

Private Sub m_cBF_ValidationFailed(ByVal sPath As String, bKeepOpen As Boolean)
   ' MsgBox sPath
   ' bKeepOpen = False
End Sub
