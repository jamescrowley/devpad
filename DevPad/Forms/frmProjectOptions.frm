VERSION 5.00
Begin VB.Form frmProjectProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Project Properties"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3930
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProjectOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtRoot 
      Height          =   285
      Left            =   1185
      TabIndex        =   11
      Top             =   405
      Width           =   2655
   End
   Begin VB.TextBox txtFileName 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1185
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   90
      Width           =   2655
   End
   Begin VB.TextBox txtProjectName 
      Height          =   285
      Left            =   1185
      TabIndex        =   0
      Top             =   735
      Width           =   2655
   End
   Begin VB.TextBox txtAuthor 
      Height          =   285
      Left            =   1185
      TabIndex        =   1
      Top             =   1065
      Width           =   2655
   End
   Begin VB.TextBox txtDescription 
      Height          =   1125
      Left            =   1185
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1395
      Width           =   2655
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "1001"
      Default         =   -1  'True
      Height          =   345
      Left            =   1380
      TabIndex        =   5
      Top             =   2730
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "1002"
      Height          =   345
      Left            =   2655
      TabIndex        =   4
      Top             =   2730
      Width           =   1215
   End
   Begin VB.Label lblLabel 
      Caption         =   "1173"
      Height          =   255
      Index           =   4
      Left            =   75
      TabIndex        =   12
      Top             =   420
      Width           =   1095
   End
   Begin VB.Label lblLabel 
      Caption         =   "1172"
      Height          =   255
      Index           =   3
      Left            =   75
      TabIndex        =   10
      Top             =   105
      Width           =   1095
   End
   Begin VB.Label lblLabel 
      Caption         =   "1169"
      Height          =   255
      Index           =   2
      Left            =   75
      TabIndex        =   8
      Top             =   750
      Width           =   1095
   End
   Begin VB.Label lblLabel 
      Caption         =   "1170"
      Height          =   255
      Index           =   1
      Left            =   75
      TabIndex        =   7
      Top             =   1065
      Width           =   1095
   End
   Begin VB.Label lblLabel 
      Caption         =   "1171"
      Height          =   255
      Index           =   0
      Left            =   75
      TabIndex        =   6
      Top             =   1380
      Width           =   1095
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      Caption         =   "1240"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   135
      MouseIcon       =   "frmProjectOptions.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2790
      Width           =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   120
      X2              =   3870
      Y1              =   2625
      Y2              =   2625
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   3870
      Y1              =   2640
      Y2              =   2640
   End
End
Attribute VB_Name = "frmProjectProperties"
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
Public bCancel As Boolean

Private Sub cmdCancel_Click()
    'dialog has been cancelled
    bCancel = True
    'hide
    Hide
End Sub
Private Sub cmdOK_Click()
    bCancel = False
    'hide the dialog
    Hide
End Sub
Private Sub Form_Load()
    'load resource strings
    LoadResStrings Controls
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        'hide the form
        Hide
        'and abort unload
        Cancel = -1
    End If
End Sub
Private Sub lblHelp_Click()
    'display the help topic for this dlg
    cDialog.ShowHelpTopic 4, hWnd
End Sub
