VERSION 5.00
Begin VB.Form frmGenerate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DLL Base Address Generator"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGenerate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   345
      Left            =   3810
      TabIndex        =   10
      Top             =   1860
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   345
      Left            =   3810
      TabIndex        =   9
      Top             =   1470
      Width           =   1215
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   345
      Left            =   3810
      TabIndex        =   8
      Top             =   1005
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   345
      Left            =   3810
      TabIndex        =   4
      Top             =   3285
      Width           =   1215
   End
   Begin VB.TextBox txtBaseAddress 
      Height          =   285
      HideSelection   =   0   'False
      Left            =   1290
      TabIndex        =   0
      Top             =   1035
      Width           =   2385
   End
   Begin VB.ListBox lstUsed 
      Height          =   1620
      Left            =   1290
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1440
      Width           =   2385
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      Caption         =   "&Help"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   165
      MouseIcon       =   "frmGenerate.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   3345
      Width           =   360
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   4
      X1              =   120
      X2              =   5055
      Y1              =   3150
      Y2              =   3150
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   2
      X1              =   30
      X2              =   7500
      Y1              =   885
      Y2              =   885
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   4470
      Picture         =   "frmGenerate.frx":0316
      Top             =   165
      Width           =   480
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "DLL Base Address Generator"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   165
      TabIndex        =   6
      Top             =   75
      Width           =   3615
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Generates base addresses for your components"
      Height          =   255
      Left            =   300
      TabIndex        =   5
      Top             =   345
      Width           =   3825
   End
   Begin VB.Label lblLabel 
      Caption         =   "Used List"
      Height          =   255
      Index           =   0
      Left            =   105
      TabIndex        =   3
      Top             =   1425
      Width           =   1575
   End
   Begin VB.Label lblLabel 
      Caption         =   "Base Address"
      Height          =   255
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   1065
      Width           =   1140
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   900
      Left            =   0
      Top             =   0
      Width           =   5145
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   30
      X2              =   7500
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   5
      X1              =   105
      X2              =   5010
      Y1              =   3165
      Y2              =   3165
   End
End
Attribute VB_Name = "frmGenerate"
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
'Public cParent As clsDLLBaseGenerator
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_COPY = &H301
Private Const LB_FINDSTRING = &H18F
Private Sub cmdAdd_Click()
    If cmdAdd.Enabled = False Then Exit Sub
    lstUsed.AddItem txtBaseAddress.Text
    EnableDelete
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub EnableDelete()
    If lstUsed.ListCount = 0 Then
        cmdDelete.Enabled = False
    Else
        cmdDelete.Enabled = True
        lstUsed.ListIndex = 0
    End If
    Call txtBaseAddress_Change
End Sub
Private Sub cmdDelete_Click()
On Error Resume Next
Dim lLastIndex As Long
    lLastIndex = lstUsed.ListIndex
    lstUsed.RemoveItem lstUsed.ListIndex
    If lstUsed.ListCount > lLastIndex Then
        lstUsed.ListIndex = lLastIndex
    Else
        lstUsed.ListIndex = lstUsed.ListCount - 1
    End If
    EnableDelete
End Sub

Private Sub cmdGenerate_Click()
    Dim i As Long
    Screen.MousePointer = vbHourglass
    '// 65536 * 256
Start:
    '// DLL Base Addresses must be a multiple of 64K (65536)
    '// LowerBound = 16777216 (which is 64K * 256)
    '// UpperBound = 2147483648 (which is 64K * 32768)
    '// 1) Generate Random Number between 256 and 32768
    '// 2) Multiply by 64K
    '// 3) Convert generated value to a Long value
    '// 4) Convert generated value to Hex
    '// 5) Convert generated string to lowercase (so that the letters are lower case)
    '// 6) Add &H in front of string
    txtBaseAddress = "&H" & LCase(Hex(CLng((32768 - 256 + 1) * Rnd + 256) * 65536))
    '// check if this value has already been used
    For i = 0 To lstUsed.ListCount - 1
        If lstUsed.List(i) = txtBaseAddress Then
            '// found on used list
            GoTo Start
        End If
    Next
    Screen.MousePointer = vbDefault
    Call txtBaseAddress_Change
    
    On Error Resume Next
    If txtBaseAddress.Visible Then txtBaseAddress.SetFocus
    txtBaseAddress.SelStart = 0
    txtBaseAddress.SelLength = Len(txtBaseAddress.Text)
End Sub
Private Sub Form_Load()
    Dim intFileNum As Integer
    Dim strLine As String
    'cFunc.pLoadResStrings Controls
    intFileNum = FreeFile
    On Error GoTo FileErr
    Open App.Path & "\UsedBaseAddress.txt" For Input As intFileNum
    lstUsed.Clear

    Do While Not EOF(intFileNum)
        Line Input #intFileNum, strLine
        lstUsed.AddItem strLine
    Loop
    Close #intFileNum
FileErr:
    If lstUsed.ListCount <> 0 Then
        cmdDelete.Enabled = True
        lstUsed.ListIndex = 0
    End If
    Call cmdGenerate_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
'// save items
    Dim intFileNum As Integer
    Dim i As Long
    intFileNum = FreeFile
    Open App.Path & "\UsedBaseAddress.txt" For Output As intFileNum
    For i = 0 To lstUsed.ListCount - 1
        Print #intFileNum, lstUsed.List(i)
    Next
    Close #intFileNum
End Sub


Private Sub lblHelp_Click()
    ShowHTMLHelpTopic 1, hWnd
End Sub

Private Sub txtBaseAddress_Change()
    If txtBaseAddress.Text = Empty Then
        cmdAdd.Enabled = False
    Else
        cmdAdd.Enabled = True
    End If
    lstUsed.ListIndex = SendMessage(lstUsed.hWnd, LB_FINDSTRING, -1, ByVal txtBaseAddress.Text)
    If lstUsed.ListIndex = -1 Or lstUsed.List(lstUsed.ListIndex) <> txtBaseAddress.Text Then
        cmdAdd.Enabled = True
    Else
        cmdAdd.Enabled = False
    End If
End Sub

Private Sub txtBaseAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyC And Shift = vbCtrlMask Then
        If cFunc.Dialogs.ShowYesNo("Add to used list?", False) = 1 Then ' 1= yes
            Call cmdAdd_Click
        End If
        SendMessage txtBaseAddress.hWnd, WM_COPY, 0&, 0&
    End If
End Sub
