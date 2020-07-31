VERSION 5.00
Begin VB.UserControl vbwColourPicker 
   ClientHeight    =   1980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2235
   ScaleHeight     =   1980
   ScaleWidth      =   2235
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      ForeColor       =   &H80000008&
      Height          =   1965
      Left            =   0
      ScaleHeight     =   1935
      ScaleWidth      =   2205
      TabIndex        =   0
      Top             =   0
      Width           =   2235
      Begin VB.CommandButton cmdMore 
         Caption         =   "More Colours..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   1
         Top             =   1650
         Width           =   2230
      End
      Begin VB.Label lblColour 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   15
         TabIndex        =   2
         Top             =   30
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Shape shpActive 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   0
         Top             =   15
         Visible         =   0   'False
         Width           =   270
      End
   End
End
Attribute VB_Name = "vbwColourPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_vColour As OLE_COLOR
Private m_lColourCount As Long
Private m_lActive As Long
Public Event ColourChanged(NewColour As OLE_COLOR)
Public Event CancelPick()
Public Property Get Colour() As OLE_COLOR
    Colour = m_vColour
End Property
Public Property Let Colour(vNew As OLE_COLOR)
Dim i As Long
Dim bFound As Boolean
    m_vColour = vNew
    For i = 1 To m_lColourCount
        If lblColour(i).BackColor = vNew Then
            pSetActive (i)
            bFound = True
            Exit For
        End If
    Next
    If bFound = False Then
        shpActive.Visible = False
        m_lActive = 0
    End If
End Property
Private Sub pLoadColourPicker()
Dim i As Long
Dim lRed As Long
Dim lOther As Long
Dim lAmount As Long
Dim lType As Long
    
    Dim lAlt As Long
    m_lColourCount = 0
    For lType = 1 To 8
        lAmount = 100
        lRed = 255
        lOther = 255
        lAlt = 223
        If lType = 1 Then
            lOther = 250 + (lAmount / 2)
            lRed = 0
            lAmount = (lAmount / 2)
        ElseIf lType = 3 Then
            lOther = 280
            lAlt = 280
            lAmount = 58
        Else
            lRed = 261
            lOther = 261
            lAmount = 87 '.6666666666667 '68
        End If
        For i = 1 To 6
            m_lColourCount = m_lColourCount + 1
            On Error Resume Next
            Load lblColour(m_lColourCount)
            
            On Error GoTo 0
            With lblColour(m_lColourCount)
                .Left = 30 + (270 * (lType - 1))
                .Top = 30 + (270 * (i - 1))
                If lOther >= lAmount Then
                    lOther = lOther - lAmount
                    lAlt = lOther + (15 * i)
                    If lAlt < 0 Then lAlt = 0
                ElseIf lRed >= lAmount Then
                    lAmount = 55
                    lOther = 0
                    lRed = lRed - lAmount
                    lAlt = lAlt - lAmount '(15 * i)
                    If lAlt < 0 Then lAlt = 0
                End If
                Select Case lType
                Case 1
                    .BackColor = RGB(lOther, lOther, lOther)
                Case 2
                    .BackColor = RGB(lRed, lOther, lOther)
                Case 3
                    .BackColor = RGB(lRed, lAlt, lOther)
                Case 4
                    .BackColor = RGB(lRed, lRed, lOther)
                Case 5
                    .BackColor = RGB(lOther, lRed, lOther)
                Case 6
                    .BackColor = RGB(lOther, lRed, lRed)
                Case 7
                    .BackColor = RGB(lOther, lOther, lRed)
                Case 8
                    .BackColor = RGB(lRed, lOther, lRed)
                
'                Case 8
'                    .BackColor = RGB(lRed, lRed, lRed)
                
                End Select
                .Visible = True
            End With
            'If lOther < lAmount And lRed < (lAmount) Then Exit For
        Next
    Next
End Sub
'
'Private Sub Combo1_DropDown()
'Dim tR As RECT
'Dim tPicker As RECT
'Dim tP As POINTAPI
'    SetParent Picture1.hWnd, UserControl.Parent.hWnd
'    GetWindowRect Combo1.hWnd, tR
'    GetWindowRect Picture1.hWnd, tPicker
'    ScreenToClient UserControl.Parent.hWnd, tP
'    MoveWindow Picture1.hWnd, tP.x + tR.Left, tP.y + tR.Top + (tPicker.Bottom - tPicker.Top), tPicker.Right - tPicker.Left, tPicker.Bottom - tPicker.Top, 1
'    ShowWindow Picture1.hWnd, SW_SHOW
'    Picture1.SetFocus
'End Sub
'
'Private Sub Picture1_LostFocus()
'    ShowWindow Picture1.hWnd, SW_HIDE
'End Sub

Private Sub cmdMore_Click()
    Dim cCmDlg As GCommonDialog
    Dim lColour As Long
    Set cCmDlg = New GCommonDialog
    lColour = CLng(m_vColour)
    If cCmDlg.VBChooseColor(lColour, True, True, , UserControl.Parent.hWnd) Then
        m_vColour = lColour
        RaiseEvent ColourChanged(m_vColour)
    End If
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        m_lActive = 0
        RaiseEvent CancelPick
    Case vbKeyReturn
        pColourPicked
    Case vbKeyDown
        pMoveActive True, False
    Case vbKeyUp
        pMoveActive False, False
    Case vbKeyLeft
        pMoveActive False, True
    Case vbKeyRight
        pMoveActive True, True
    End Select
End Sub

Private Sub UserControl_ExitFocus()
    RaiseEvent CancelPick
End Sub

Private Sub UserControl_Initialize()
    pLoadColourPicker
End Sub

Private Sub lblColour_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    pSetActive (Index)
    Picture1.SetFocus
End Sub

Private Sub lblColour_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    pColourPicked
End Sub
Private Sub pColourPicked()
    shpActive.Visible = False
    If m_lActive <> 0 Then
        m_vColour = lblColour(m_lActive).BackColor
        RaiseEvent ColourChanged(m_vColour)
    End If
End Sub
Private Sub pMoveActive(bForward As Boolean, bAcross As Boolean)
Dim nIndex As Integer
    nIndex = m_lActive
    'across a column
    
    If bAcross Then
        nIndex = nIndex + IIf(bForward, 6, -6)
    Else
        nIndex = nIndex + IIf(bForward, 1, -1)
    End If
        If nIndex > (lblColour.Count - 1) Then
            If (nIndex = lblColour.Count + 6 And bAcross = True) Or (nIndex = lblColour.Count And bAcross = False) Then
                nIndex = 1
            Else
                'reached end...
                nIndex = nIndex - lblColour.Count + 2
            End If
        ElseIf nIndex <= 0 Then
            If (nIndex = -6 And bAcross = True) Or (nIndex = 0 And bAcross = False) Then
                'reached beginning
                nIndex = lblColour.Count - 1
            Else
                nIndex = lblColour.Count - Abs(nIndex) - 2
            End If
        End If
   ' Else
        
        
   ' End If
    pSetActive nIndex
End Sub
Private Sub pSetActive(Index As Integer)
    shpActive.Left = lblColour(Index).Left - 15
    shpActive.Top = lblColour(Index).Top - 15
    shpActive.Visible = True
    m_lActive = Index
End Sub

Private Sub UserControl_Resize()
    Height = 1965
    Width = 2235
End Sub

Private Sub UserControl_Show()
    Picture1.SetFocus
End Sub
