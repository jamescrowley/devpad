VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFlatOpt"
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
Implements ISubclass
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_COMMAND = &H111
Private Const WM_PAINT = &HF
Private Const BN_CLICKED = 0
Private Const WM_KEYDOWN = &H100
Private Const BM_SETSTATE = &HF3
Private Const BM_SETCHECK = &HF1

Private m_hWnd             As Long
Private m_hWndParent       As Long
Private m_bSubclass        As Boolean

Public Sub Attach(ByVal lhWnd As Long)
    DetachMessages
    m_hWnd = lhWnd
    m_hWndParent = GetParent(lhWnd)
    AttachMessages
End Sub
Private Sub AttachMessages()
    If (m_hWnd <> 0) Then
       ' If InDevelopment Then Exit Sub
        AttachMessage Me, m_hWnd, WM_PAINT
        'AttachMessage Me, m_hWnd, WM_LBUTTONDOWN
        'AttachMessage Me, m_hWnd, WM_LBUTTONUP
        'AttachMessage Me, m_hWnd, WM_KEYDOWN
        AttachMessage Me, m_hWnd, BM_SETSTATE
        AttachMessage Me, m_hWnd, BM_SETCHECK
       ' AttachMessage Me, m_hWnd, wM_notify
        AttachMessage Me, m_hWndParent, WM_COMMAND
        m_bSubclass = True
    End If
End Sub
Private Sub DetachMessages()
    If (m_bSubclass) Then
        DetachMessage Me, m_hWnd, WM_PAINT
'        DetachMessage Me, m_hWnd, WM_LBUTTONDOWN
'        DetachMessage Me, m_hWnd, WM_LBUTTONUP
'        DetachMessage Me, m_hWnd, WM_KEYDOWN
        DetachMessage Me, m_hWnd, BM_SETSTATE
        DetachMessage Me, m_hWnd, BM_SETCHECK
      '  DetachMessage Me, m_hWnd, wM_notify
        DetachMessage Me, m_hWndParent, WM_COMMAND
    End If
    m_hWnd = 0: m_hWndParent = 0
End Sub

Private Sub Draw(clrTopLeft As OLE_COLOR, clrBottomRight As OLE_COLOR)
Dim rcItem As RECT
Dim pDC As Long
   

    GetClientRect m_hWnd, rcItem
    pDC = GetDC(m_hWnd)

    rcItem.Left = 0
    rcItem.Top = rcItem.Top + 2
    rcItem.Bottom = rcItem.Bottom - 2
    rcItem.Right = 13
    Draw3DRect pDC, rcItem, vbButtonFace, vbButtonFace, True
    rcItem.Left = 1
    rcItem.Top = rcItem.Top + 1
    rcItem.Bottom = rcItem.Bottom
    rcItem.Right = 1
    Draw3DRect pDC, rcItem, vbButtonFace, vbButtonFace, True
  '  rcItem.Left = 1
    rcItem.Top = rcItem.Top
    rcItem.Bottom = rcItem.Bottom - 1
    rcItem.Right = 12
    Draw3DRect pDC, rcItem, vbButtonShadow, vb3DHighlight, True
    DeleteDC pDC
End Sub

Public Sub OnPaint()
    Draw vbButtonShadow, vb3DHighlight
End Sub

Private Sub Class_Terminate()
   DetachMessages
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
End Property
Private Property Get ISubclass_MsgResponse() As EMsgResponse
    Select Case CurrentMessage
    Case WM_PAINT, BM_SETSTATE, BM_SETCHECK 'WM_LBUTTONDOWN, WM_LBUTTONUP, WM_KEYDOWN,  ', wM_notify
        ISubclass_MsgResponse = emrPreprocess
    Case Else
        ISubclass_MsgResponse = emrPostProcess
    End Select
End Property
Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim bDown As Boolean
Dim bFocus As Boolean
   ' Debug.Print iMsg
    Select Case iMsg
    Case WM_COMMAND
        OnPaint
        Debug.Print wParam \ &H10000
        Select Case wParam \ &H10000
        Case BN_CLICKED
            OnPaint
        End Select
    Case WM_PAINT, BM_SETSTATE, BM_SETCHECK ','WM_KEYDOWN, WM_LBUTTONDOWN, WM_LBUTTONUP,
        'If iMsg = BM_SETCHECK Then Debug.Print "setcheck"
        OnPaint
   End Select
End Function
