VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   4575
   ClientLeft      =   3555
   ClientTop       =   3000
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   1
   Icon            =   "frmAbout.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox picRegInfo 
      Height          =   645
      Left            =   1155
      ScaleHeight     =   585
      ScaleWidth      =   3165
      TabIndex        =   2
      Top             =   1935
      Width           =   3225
      Begin VB.Label lblName 
         Caption         =   "<Registered Name>"
         Height          =   240
         Left            =   60
         TabIndex        =   4
         Top             =   45
         Width           =   3015
      End
      Begin VB.Label lblCompany 
         Caption         =   "<Company>"
         Height          =   240
         Left            =   60
         TabIndex        =   3
         Top             =   330
         Width           =   2505
      End
   End
   Begin VB.PictureBox picPic 
      BackColor       =   &H00B1523B&
      FillStyle       =   0  'Solid
      Height          =   2460
      Left            =   105
      ScaleHeight     =   2400
      ScaleWidth      =   825
      TabIndex        =   1
      Top             =   120
      Width           =   885
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "1001"
      Default         =   -1  'True
      Height          =   345
      Left            =   3465
      TabIndex        =   0
      Top             =   4140
      Width           =   1125
   End
   Begin DevPad.ctlFrame ctlFrame1 
      Height          =   600
      Left            =   1155
      TabIndex        =   5
      Top             =   120
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   1058
      Begin VB.Label lblLabel 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Developers Pad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   90
         TabIndex        =   7
         Top             =   60
         Width           =   2040
      End
      Begin VB.Label lblBuild 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Build 1.5.200"
         Height          =   240
         Left            =   90
         TabIndex        =   6
         Top             =   315
         Width           =   2145
      End
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   0
      X1              =   105
      X2              =   4575
      Y1              =   3990
      Y2              =   3990
   End
   Begin VB.Label lblLabel 
      Caption         =   "© 1999-2000 VB Web Development"
      Height          =   270
      Index           =   1
      Left            =   1170
      TabIndex        =   11
      Top             =   825
      Width           =   2685
   End
   Begin VB.Label lblLabel 
      Caption         =   "All Rights Reserved"
      Height          =   255
      Index           =   2
      Left            =   1170
      TabIndex        =   10
      Top             =   1110
      Width           =   2505
   End
   Begin VB.Label lblWeb 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.developerspad.com"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1170
      MouseIcon       =   "frmAbout.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   1410
      Width           =   2400
   End
   Begin VB.Label lblLabel 
      Caption         =   $"frmAbout.frx":0316
      Height          =   1035
      Index           =   0
      Left            =   165
      TabIndex        =   8
      Top             =   2880
      Width           =   4350
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   105
      X2              =   4575
      Y1              =   2745
      Y2              =   2745
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   2
      X1              =   105
      X2              =   4575
      Y1              =   2760
      Y2              =   2745
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   4
      X1              =   105
      X2              =   4575
      Y1              =   4005
      Y2              =   3990
   End
End
Attribute VB_Name = "frmAbout"
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

Private Sub cmdClose_Click()
    'close
    Unload Me
End Sub

Private Sub Form_Load()
Dim cReg As clsRegistry
    'Retreive the build number
    lblBuild.Caption = "Build " & App.Major & "." & App.Minor & "." & App.Revision
    'init the registry class
    Set cReg = New clsRegistry
    cReg.ClassKey = HKEY_LOCAL_MACHINE
    cReg.SectionKey = "Software\Developers Pad\InstallInfo"
    'Get the registered name
    cReg.ValueKey = "Name"
    lblName.Caption = cReg.Value
    'Get the registered company
    cReg.ValueKey = "Company"
    lblCompany.Caption = cReg.Value

    LoadResStrings Controls
    'Set the 'thin3d' borders...
    SetThin3DBorder picRegInfo.hWnd
    SetThin3DBorder picPic.hWnd
End Sub

Private Sub lblWeb_Click()
    'open website
    ShellFunc "http://www.developerspad.com"
End Sub
