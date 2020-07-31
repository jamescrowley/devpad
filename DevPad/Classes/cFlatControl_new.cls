VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFlatCombo"
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
' Based on the FlatControl Class @ vbaccelerator.com
' Modified by VB Web @ vbweb.co.uk

Implements ISubclass
Private Const WM_PAINT = &HF
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long

Private m_hWnd             As Long
Private m_bSubclass        As Boolean
Private m_BackColor        As OLE_COLOR

Public Property Let BackColor(vNew As OLE_COLOR)
    m_BackColor = vNew
End Property

Public Sub Attach(lhWnd As Long)
    DetachMessages
    m_hWnd = lhWnd
    AttachMessages
End Sub
Private Sub AttachMessages()
   ' If InDevelopment Then Exit Sub
    If (m_hWnd <> 0) Then
        AttachMessage Me, m_hWnd, WM_PAINT
        m_bSubclass = True
    End If
End Sub
Private Sub DetachMessages()
    If (m_bSubclass) Then
        DetachMessage Me, m_hWnd, WM_PAINT
    End If
    m_hWnd = 0
End Sub

Private Sub DrawCombo()
Dim rcItem As RECT
Dim pDC As Long

    GetClientRect m_hWnd, rcItem
    pDC = GetDC(m_hWnd)
    ' Cover up dark 3D shadow.
    Draw3DRect pDC, rcItem, m_BackColor, m_BackColor
    InflateRect rcItem, -1, -1
    Draw3DRect pDC, rcItem, vbButtonShadow, vb3DHighlight
    DeleteDC pDC
End Sub

Private Sub Class_Initialize()
    m_BackColor = vbButtonFace
End Sub

Private Sub Class_Terminate()
   DetachMessages
End Sub
Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    ISubclass_MsgResponse = emrPreprocess
End Property
Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    DrawCombo
End Function

