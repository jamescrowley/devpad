VERSION 5.00
Begin VB.UserControl vbwSpinner 
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2340
   ScaleHeight     =   1575
   ScaleWidth      =   2340
   Begin VB.TextBox txtFocus 
      Height          =   285
      Left            =   -1000
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   615
      Width           =   675
   End
   Begin VB.Timer tmrChange 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1575
      Top             =   675
   End
   Begin VB.TextBox txtEntry 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "0"
      Top             =   0
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   510
      X2              =   720
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Image imgDown 
      Height          =   60
      Left            =   495
      Picture         =   "vbwSpinner.ctx":0000
      Top             =   180
      Width           =   225
   End
   Begin VB.Image imgUp 
      Height          =   60
      Left            =   495
      Picture         =   "vbwSpinner.ctx":00BA
      Top             =   30
      Width           =   225
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   510
      X2              =   720
      Y1              =   135
      Y2              =   135
   End
End
Attribute VB_Name = "vbwSpinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private nChanging As Integer
Private bIgnore As Boolean
Public Event Change()
Private m_lMin As Long
Private m_lMax As Long
Private Sub imgDown_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If txtEntry.Enabled = False Then Exit Sub
    imgDown.Top = 195
    nChanging = -1
    tmrChange.Enabled = True
End Sub
Private Sub imgDown_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If txtEntry.Enabled = False Then Exit Sub
    imgDown.Top = 180
    nChanging = 0
    Increment -1
End Sub
Private Sub imgUp_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If txtEntry.Enabled = False Then Exit Sub
    nChanging = 1
    imgUp.Top = 15
    tmrChange.Enabled = True
End Sub
Private Sub imgUp_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If txtEntry.Enabled = False Then Exit Sub
    nChanging = 0
    imgUp.Top = 30
    Increment 1
End Sub
Private Sub tmrChange_Timer()
Static lDelay As Long
    If nChanging = 0 Then
        tmrChange.Enabled = False
        lDelay = 0
    Else
        If lDelay < 3 Then
            lDelay = lDelay + 1
        Else
            Increment
        End If
    End If
End Sub
Private Sub Increment(Optional nAmount As Integer = 0)
    If nAmount = 0 Then nAmount = nChanging
    If txtEntry.Text = "" Then txtEntry.Text = "0"
    
    If txtEntry.Text = m_lMax And nAmount > 0 Then
        'ignore
    ElseIf txtEntry.Text = m_lMin And nAmount < 0 Then
        'ignore
    Else
        txtEntry.Text = txtEntry.Text + nAmount
    End If
    txtFocus.SetFocus
    RaiseEvent Change
End Sub

Private Sub txtEntry_Change()
    If bIgnore Then Exit Sub
    bIgnore = True
    If IsNumeric(txtEntry) Then
        If txtEntry.Text > m_lMax Then
            txtEntry.Text = m_lMax
        ElseIf txtEntry.Text < m_lMin Then
            txtEntry.Text = m_lMin
        End If
        RaiseEvent Change
    End If
    bIgnore = False
End Sub

Private Sub txtEntry_KeyPress(KeyAscii As Integer)
    KeyAscii = NumericOnly(KeyAscii)
End Sub
Private Sub UserControl_Initialize()
    SetThin3DBorder txtEntry.hwnd
End Sub
Public Property Get txtText() As TextBox
    Set txtText = txtEntry
End Property
Public Property Get Value() As Long
    Value = txtEntry.Text
End Property
Public Property Let Value(lNewValue As Long)
    txtEntry.Text = lNewValue
End Property
Public Property Get Min() As Long
    Min = m_lMin
End Property
Public Property Let Min(lNewValue As Long)
    m_lMin = lNewValue
End Property
Public Property Get Max() As Long
    Max = m_lMax
End Property
Public Property Let Max(lNewValue As Long)
    m_lMax = lNewValue
End Property
Public Property Get Enabled() As Boolean
    Enabled = txtEntry.Enabled
End Property
Public Property Let Enabled(bNewValue As Boolean)
    txtEntry.Enabled = bNewValue
End Property
