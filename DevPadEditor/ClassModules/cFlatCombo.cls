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
Private Const PS_SOLID = 0
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

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
    'If InDevelopment Then Exit Sub
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
Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    DrawCombo
End Function
Private Function Draw3DRect(ByVal hdc As Long, ByRef rcItem As RECT, ByVal oTopLeftColor As OLE_COLOR, ByVal oBottomRightColor As OLE_COLOR, Optional ByVal bMod As Boolean = False)
Dim hPen As Long
Dim hPenOld As Long
Dim tp As POINTAPI
   hPen = CreatePen(PS_SOLID, 1, TranslateColor(oTopLeftColor))
   hPenOld = SelectObject(hdc, hPen)
   MoveToEx hdc, rcItem.Left, rcItem.Bottom - 1, tp
   LineTo hdc, rcItem.Left, rcItem.Top
   LineTo hdc, rcItem.Right - 1, rcItem.Top
   SelectObject hdc, hPenOld
   DeleteObject hPen
   If (rcItem.Left <> rcItem.Right) Then
      hPen = CreatePen(PS_SOLID, 1, TranslateColor(oBottomRightColor))
      hPenOld = SelectObject(hdc, hPen)
      LineTo hdc, rcItem.Right - 1, rcItem.Bottom - 1
      LineTo hdc, rcItem.Left - (IIf(bMod, 1, 0)), rcItem.Bottom - 1
      SelectObject hdc, hPenOld
      DeleteObject hPen
   End If
End Function
