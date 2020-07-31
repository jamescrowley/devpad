VERSION 5.00
Begin VB.Form frmMessage 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2160
   ClientLeft      =   30
   ClientTop       =   240
   ClientWidth     =   5475
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmErrMessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtValue 
      Height          =   285
      Left            =   975
      TabIndex        =   0
      Top             =   795
      Width           =   3045
   End
   Begin VB.PictureBox picDetails 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   105
      Left            =   15
      ScaleHeight     =   105
      ScaleWidth      =   180
      TabIndex        =   8
      Top             =   2040
      Width           =   180
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "1002"
      Height          =   345
      Left            =   4080
      TabIndex        =   4
      Top             =   1230
      Width           =   1200
   End
   Begin VB.TextBox txtDetails 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1395
      Left            =   45
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "frmErrMessage.frx":000C
      Top             =   2430
      Width           =   5325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "1001"
      Height          =   345
      Left            =   2760
      TabIndex        =   3
      Top             =   1230
      Width           =   1215
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "1008"
      Height          =   345
      Left            =   2775
      TabIndex        =   2
      Top             =   1230
      Width           =   1200
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "1007"
      Height          =   345
      Left            =   1470
      TabIndex        =   1
      Top             =   1230
      Width           =   1200
   End
   Begin VB.Image imgEdit 
      Height          =   480
      Left            =   180
      Picture         =   "frmErrMessage.frx":0012
      Top             =   960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "Debugging Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   60
      TabIndex        =   9
      Top             =   2160
      Width           =   2970
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      Caption         =   "1240"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4920
      MouseIcon       =   "frmErrMessage.frx":031C
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   1860
      Width           =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   165
      X2              =   5325
      Y1              =   1740
      Y2              =   1740
   End
   Begin VB.Image imgAlert 
      Height          =   480
      Left            =   225
      Picture         =   "frmErrMessage.frx":0626
      Top             =   225
      Width           =   480
   End
   Begin VB.Image imgUp 
      Height          =   60
      Left            =   675
      Picture         =   "frmErrMessage.frx":10E0
      Top             =   1230
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgDown 
      Height          =   60
      Left            =   165
      Picture         =   "frmErrMessage.frx":119A
      Top             =   1185
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblMsg 
      Caption         =   "The message...."
      Height          =   900
      Left            =   960
      TabIndex        =   7
      Top             =   225
      Width           =   4215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   150
      X2              =   5310
      Y1              =   1755
      Y2              =   1755
   End
End
Attribute VB_Name = "frmMessage"
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
Option Explicit
Private m_vResult As ShowYesNoResult
Public Property Get Result() As ShowYesNoResult
    Result = m_vResult
End Property
Public Sub ResetResult()
    m_vResult = None
End Sub
'*** Click Events ***
Private Sub cmdCancel_Click()
    ' cancel Action
    m_vResult = Cancelled
    Hide
End Sub
Private Sub cmdNo_Click()
    'no is clicked
    m_vResult = No
    Hide
End Sub
Private Sub cmdYes_Click()
    'yes is clicked
    m_vResult = Yes
    Hide
End Sub
Private Sub cmdOK_Click()
    'ok...
    Hide
End Sub


Private Sub lblHelp_Click()
    cDialog.ShowHelpTopic 10, hWnd
End Sub

Private Sub picDetails_Click()
    'hide or show the more detailed info...
    If Height = 2535 Then
        Height = 4260
        'change to arrow pointing up
        Set picDetails.Picture = imgUp.Picture
        'adjust top
        picDetails.Top = picDetails.Top - 10
    Else
        Height = 2535
        'adjust top
        picDetails.Top = picDetails.Top + 10
        'change to arrow pointing down
        Set picDetails.Picture = imgDown.Picture
    End If
End Sub
'*** Init ***
Public Sub InitForm(bYesNo As Boolean, bCancel As Boolean, bEdit As Boolean)
    'prepares the form, depending on what things we
    'want displayed
    'adjust the pos of buttons
    cmdOK.Left = IIf(bCancel, 2760, 4080)
    cmdYes.Left = IIf(bCancel, 1470, 2760)
    cmdNo.Left = IIf(bCancel, 2760, 4080)
    'set which is the default button
    cmdYes.Default = bYesNo
    cmdOK.Default = Not bYesNo
    'display alert/or edit icon
    imgAlert.Visible = Not bEdit
    imgEdit.Visible = bEdit
    'display edit box if we want input
    txtValue.Visible = bEdit
    'display current buttons
    cmdOK.Visible = Not bYesNo
    cmdYes.Visible = bYesNo
    cmdNo.Visible = bYesNo
    cmdCancel.Visible = bCancel ' hide cancel
    
    cmdOK.Cancel = Not bCancel And Not bYesNo
    cmdCancel.Cancel = bCancel
    'hide the details box
    picDetails.Visible = False
End Sub
Private Sub Form_Load()
    Height = 2535
    LoadResStrings Controls
    RemoveCloseItem hWnd
    imgAlert.Left = 225
    imgAlert.Top = 225
    imgEdit.Left = 225
    imgEdit.Top = 225
    Set picDetails.Picture = imgDown.Picture
End Sub

