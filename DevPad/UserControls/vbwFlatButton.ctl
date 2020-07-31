VERSION 5.00
Begin VB.UserControl vbwFlatButton 
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2865
   ScaleHeight     =   2175
   ScaleWidth      =   2865
   Begin VB.Image imgImage 
      Height          =   240
      Left            =   255
      Picture         =   "vbwFlatButton.ctx":0000
      Top             =   0
      Width           =   240
   End
   Begin VB.Image imgImageDisabled 
      Height          =   240
      Left            =   30
      Picture         =   "vbwFlatButton.ctx":014A
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "vbwFlatButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISubclass
Public Enum EDrawStyle
   FC_DRAWNORMAL = &H1
   FC_DRAWRAISED = &H2
   FC_DRAWPRESSED = &H4
End Enum
Private Const WM_TIMER = &H113
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Private m_bSubclass         As Boolean
Private m_bMouseOver        As Boolean
Private m_bDown             As Boolean

Public Event Click()
Private Sub AttachMessages()
    AttachMessage Me, hWnd, WM_TIMER
    m_bSubclass = True
End Sub
Private Sub DetachMessages()
    If (m_bSubclass) Then
         DetachMessage Me, hWnd, WM_TIMER
    End If
End Sub

Private Sub Draw(ByVal dwStyle As EDrawStyle, clrTopLeft As OLE_COLOR, clrBottomRight As OLE_COLOR)
Dim rcItem As RECT
    rcItem.Left = 0
    rcItem.Right = Width / Screen.TwipsPerPixelX
    rcItem.Top = 0
    rcItem.Bottom = Height / Screen.TwipsPerPixelY
    Draw3DRect UserControl.hdc, rcItem, clrTopLeft, clrBottomRight
End Sub

Private Sub Init()
    If Ambient.UserMode = False Then Exit Sub
    DetachMessages
    AttachMessages
End Sub

Private Sub OnPaint(bPaintEvent As Boolean)
    If m_bDown And Enabled Then
        Draw FC_DRAWPRESSED, vbButtonShadow, vb3DHighlight
        If Not bPaintEvent Then imgImage.Top = (UserControl.Height / 2) - (imgImage.Height / 2) + 8
    ElseIf m_bMouseOver And Enabled Then
        Draw FC_DRAWRAISED, vb3DHighlight, vbButtonShadow
        If Not bPaintEvent Then UserControl_Resize
    Else
        Draw FC_DRAWNORMAL, vbButtonFace, vbButtonFace
        If Not bPaintEvent Then UserControl_Resize
    End If
End Sub

Private Sub OnTimer(ByVal bCheckMouse As Boolean)
Dim bOver As Boolean
Dim rcItem As RECT
Dim tp As POINTAPI
   
   If (bCheckMouse) Then
      bOver = True
      GetCursorPos tp
      GetWindowRect hWnd, rcItem
      If (PtInRect(rcItem, tp.x, tp.y) = 0) Then
         bOver = False
      End If
   End If
   
   If Not (bOver) Then
      KillTimer hWnd, 1
      m_bMouseOver = False
      m_bDown = False
   End If

End Sub

Private Sub imgImage_Click()
    UserControl_Click
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    ISubclass_MsgResponse = emrPostProcess
End Property
Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case iMsg
    Case WM_TIMER
       OnTimer True
       If Not (m_bMouseOver) Then
          OnPaint False
       End If
    End Select
End Function

Private Sub UserControl_Click()
    If Enabled Then RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
    imgImage.Left = 30
End Sub

Private Sub UserControl_LostFocus()
    m_bDown = False
    OnPaint False
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_bDown = True
    OnPaint False
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not (m_bMouseOver) And Button = 0 Then
        m_bMouseOver = True
        OnPaint False
        SetTimer hWnd, 1, 10, 0
    End If
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_bDown = False
    OnPaint False
End Sub
Private Sub imgImage_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseDown Button, Shift, x, y
End Sub
Private Sub imgImage_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseMove Button, Shift, x, y
End Sub
Private Sub imgImage_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseUp Button, Shift, x, y
End Sub
Private Sub UserControl_Paint()
OnPaint True
'    If m_bDown And Enabled Then
'        Draw FC_DRAWPRESSED, vbButtonShadow, vb3DHighlight
'    ElseIf m_bMouseOver And Enabled Then
'        Draw FC_DRAWRAISED, vb3DHighlight, vbButtonShadow
'    Else
'        Draw FC_DRAWNORMAL, vbButtonFace, vbButtonFace
'    End If
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Init
End Sub
Public Property Get Enabled() As Boolean
    Enabled = Not imgImageDisabled.Visible
End Property
Public Property Let Enabled(New_Enabled As Boolean)
    imgImageDisabled.Visible = Not New_Enabled
    imgImage.Visible = New_Enabled
End Property
Private Sub UserControl_Resize()
    imgImage.Top = (UserControl.Height / 2) - (imgImage.Height / 2)
End Sub
Private Sub UserControl_Terminate()
    DetachMessages
End Sub
