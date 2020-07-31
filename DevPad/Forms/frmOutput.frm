VERSION 5.00
Begin VB.Form frmOutput 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Output"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4740
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
   ScaleHeight     =   3240
   ScaleWidth      =   4740
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
      Height          =   3195
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   15
      Width           =   4695
   End
End
Attribute VB_Name = "frmOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents cCommand As clsCommandLine
Attribute cCommand.VB_VarHelpID = -1
Private Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" ( _
  ByVal lpPathName As String) As Long
'Private Declare Function GetCurrentDirectory Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Sub cCommand_OutputAvailable(sOutput As String)
    txtOutput.Text = txtOutput.Text & sOutput
    DoEvents
End Sub

Private Sub Form_Load()
    Set cCommand = New clsCommandLine
End Sub
Public Sub RunCommand(sCommand As String)
    If Left(sCommand, 2) = "cd" Then
        'ChDir (Right$(sCommand, Len(sCommand) - 3))
        If SetCurrentDirectory("c:\") Then
        Else
            Debug.Print "Invalid Path"
        End If
    End If
    sCommand = Environ$("COMSPEC") & " /c " & sCommand
    'txtOutput.Text = cCommand.GetCommandOutput(sCommand, True, True, True)
    cCommand.GetCommandOutput sCommand, True, True, True
End Sub
Private Sub Form_Resize()
    'resize the textbox
    txtOutput.Move 5, 5, ScaleWidth - 10, ScaleHeight - 10
End Sub

Private Sub txtOutput_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sCommand As String
Dim lPos As Long
Dim lBegPos As Long
    If KeyCode = vbKeyReturn Then
        lPos = InStr(txtOutput.SelStart, txtOutput.Text, vbCrLf)
        If lPos = 0 Then lPos = Len(txtOutput.Text) + 1
        lBegPos = InStrRev(Left$(txtOutput.Text, txtOutput.SelStart - 1), vbCrLf) + 2
        If lBegPos = 2 Then lBegPos = 1
        sCommand = Mid$(txtOutput.Text, lBegPos, (lPos - 1) - (lBegPos - 2))
        RunCommand sCommand
    End If
End Sub
