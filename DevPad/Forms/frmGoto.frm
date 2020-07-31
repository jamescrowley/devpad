VERSION 5.00
Begin VB.Form frmGoto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Jump-To..."
   ClientHeight    =   1770
   ClientLeft      =   30
   ClientTop       =   255
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGoto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboOption 
      Height          =   315
      ItemData        =   "frmGoto.frx":000C
      Left            =   1590
      List            =   "frmGoto.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   75
      Width           =   1725
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "1000"
      Height          =   345
      Left            =   2520
      TabIndex        =   3
      Top             =   1305
      Width           =   1215
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "1022"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   345
      Left            =   3810
      TabIndex        =   4
      Top             =   1305
      Width           =   1200
   End
   Begin VB.TextBox txtGoto 
      Height          =   285
      Left            =   2655
      TabIndex        =   1
      Top             =   480
      Width           =   2235
   End
   Begin VB.Label lblVal 
      Caption         =   "1127"
      Height          =   270
      Left            =   975
      TabIndex        =   7
      Top             =   495
      Width           =   1635
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   90
      X2              =   5010
      Y1              =   1170
      Y2              =   1170
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      Caption         =   "1240"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   105
      MouseIcon       =   "frmGoto.frx":0038
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1365
      Width           =   360
   End
   Begin VB.Label lblMsg 
      Caption         =   "152"
      Height          =   225
      Left            =   975
      TabIndex        =   6
      Top             =   120
      Width           =   510
   End
   Begin VB.Image imgAlert 
      Height          =   480
      Left            =   255
      Picture         =   "frmGoto.frx":0342
      Top             =   105
      Width           =   480
   End
   Begin VB.Label lblInfo 
      Enabled         =   0   'False
      Height          =   255
      Left            =   2670
      TabIndex        =   5
      Top             =   825
      Width           =   2205
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   90
      X2              =   4995
      Y1              =   1185
      Y2              =   1185
   End
End
Attribute VB_Name = "frmGoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Private Const EM_LINEFROMCHAR = &HC9
Private cCombo As clsFlatCombo
Private Sub cboOption_Click()
    'simulate goto_change to update status text
    txtGoto_Change
End Sub
Private Sub cmdClose_Click()
    'close the form
    Unload Me
End Sub
Private Sub cmdGo_Click()
On Error GoTo ErrHandler

    Dim lLineStart As Long
    'abort if no documents open
    If DocOpen = False Then Exit Sub
    With ActiveDoc
        'hide the goto form
        Hide
        Select Case cboOption.Text
        Case LoadResString(1116) 'Line
            ' Goto Line
            ' Get line
            lLineStart = .LineIndex(txtGoto.Text) 'SendMessage(.hWnd, EM_LINEINDEX, txtGoto.Text - 1, 0&)
            If lLineStart = -1 Then
                '"Invalid line number"
                cDialog.ErrHandler vbObjectError + 1007, LoadResString(1122), "Goto.Goto", "LINENUMBER: " & txtGoto
                Exit Sub
            End If
            .SelStart = lLineStart
        Case LoadResString(1117) 'Char pos
            ' Goto char pos
            .SelStart = txtGoto.Text
        Case LoadResString(1119) 'start
            ' Goto the beginning
            .SelStart = 0
        Case LoadResString(1120) 'end
            ' Goto end of document
            '.SelStart = Len(ActiveDoc.Text)
        Case Else '!!
            Debug.Print "Coding Error"
        End Select
        On Error Resume Next
        'activate the form
        .SetFocus
        'unload this form
        Unload Me
    End With
Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Goto.cmdGo"
End Sub
Private Sub Form_Load()
    'load the resource strings
    LoadResStrings Controls
    'restore the last goto option
    cboOption.ListIndex = GetSetting(REG_KEY, "Settings", "LastGotoOption", 0)
    'make the combo flat
    Set cCombo = New clsFlatCombo
    cCombo.Attach cboOption.hWnd
    'update status text
    Call txtGoto_Change
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'save goto option
    SaveSetting REG_KEY, "Settings", "LastGotoOption", cboOption.ListIndex
End Sub

Private Sub lblHelp_Click()
    cDialog.ShowHelpTopic 9, hWnd
End Sub

Private Sub txtGoto_Change()
On Error GoTo ErrHandler

    Dim lVal            As Long
    Dim bState          As Boolean 'stores visibility of txtGo
    Dim bButtonState    As Boolean 'stores enabled state of cmdGo
    
    lVal = -1
    
    bButtonState = True
    Select Case cboOption.Text
    Case LoadResString(1116) 'Line
        'show the input box
        bState = True
        'if the goto text isn't empty, get the start of the specified
        'line
        If txtGoto.Text <> "" Then lVal = ActiveDoc.LineIndex(txtGoto.Text)
        If lVal = -1 Then
            'invalid line...
            lblInfo = LoadResString(1122) '"Invalid Line"
            'not allowed to click go!
            bButtonState = False
        Else
            'return status info
            lblInfo = LoadResString(1123) & lVal  ' "Line starts at char "
        End If
    Case LoadResString(1117) 'Char pos
        bState = True
        lblLabel = LoadResString(1129) '"Enter the Character Position"
        'get the line from char pos, if the textbox
        'isn't empty
        'If txtGoto.Text <> "" Then lVal = ActiveDoc.LineFromChar(txtGoto.Text)
        If lVal = -1 Then
            'invalid char pos
            lblInfo = LoadResString(1124) '"Invalid Char Pos"
            bButtonState = False
        Else
            'give status info
            lblInfo = LoadResString(1125) & lVal '"Char is on line "
        End If
    Case Else 'end or start
        bState = False
    End Select
    'Show/hide input controls as required
    txtGoto.Visible = bState
    lblVal.Visible = bState
    lblInfo.Visible = bState
    If txtGoto.Text <> "" Or bState = False Then
        'if txtGoto isn't empty, or we are hiding the input controls
        'use bButtonState
        cmdGo.Enabled = bButtonState
    ElseIf txtGoto.Text = Empty Then
        ' if the textbox is empty and we
        ' are expecting input then
        ' disable go button
        cmdGo.Enabled = False
    End If
    Exit Sub
ErrHandler:
    'ignore overflow errors...
    If Err <> 6 Then cDialog.ErrHandler Err, Error, "Goto.txtGoto_Change"
End Sub
Private Sub txtGoto_KeyPress(KeyAscii As Integer)
    'ensure we only accept numeric input
    KeyAscii = NumericOnly(KeyAscii)
End Sub
