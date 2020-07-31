VERSION 5.00
Begin VB.Form frmCommandLine 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Command Line"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOutput 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4380
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmCommandLine.frx":0000
      Top             =   0
      Width           =   5595
   End
End
Attribute VB_Name = "frmCommandLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents cCommand As clsCommandLine

Private Sub cCommand_OutputAvailable(sOutput As String)
    txtOutput.Text = txtOutput.Text & sOutput
    DoEvents
End Sub

Private Sub Form_Load()
    Set cCommand = New clsCommandLine
    cCommand.GetCommandOutput "c:\consoleapp.exe", True, True, True
End Sub
Private Sub Form_Resize()
    'resize the textbox
    txtOutput.Move 5, 5, ScaleWidth - 10, ScaleHeight - 10
End Sub

