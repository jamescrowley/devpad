VERSION 5.00
Begin VB.Form frmErrorConverter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Error Lookup"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmErrorConverter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   345
      Left            =   3810
      TabIndex        =   9
      Top             =   2670
      Width           =   1215
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      ItemData        =   "frmErrorConverter.frx":000C
      Left            =   930
      List            =   "frmErrorConverter.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   990
      Width           =   1650
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "&Look up"
      Default         =   -1  'True
      Height          =   345
      Left            =   3810
      TabIndex        =   3
      Top             =   975
      Width           =   1215
   End
   Begin VB.TextBox txtError 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   135
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1665
      Width           =   4875
   End
   Begin VB.TextBox txtErr 
      Height          =   285
      Left            =   2865
      TabIndex        =   0
      Top             =   1005
      Width           =   855
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   4
      X1              =   120
      X2              =   5055
      Y1              =   2535
      Y2              =   2535
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      Caption         =   "&Help"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   165
      MouseIcon       =   "frmErrorConverter.frx":0032
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   2730
      Width           =   360
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
      Picture         =   "frmErrorConverter.frx":033C
      Top             =   165
      Width           =   480
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Error Lookup"
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
      TabIndex        =   8
      Top             =   75
      Width           =   1305
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Lookup the error text for VB and Win API errors"
      Height          =   255
      Left            =   300
      TabIndex        =   7
      Top             =   345
      Width           =   3825
   End
   Begin VB.Label lblLabel 
      Caption         =   "#"
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   6
      Top             =   1020
      Width           =   210
   End
   Begin VB.Label lblLabel 
      Caption         =   "Type"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1020
      Width           =   765
   End
   Begin VB.Label lblMsg 
      Caption         =   "Description"
      Height          =   255
      Left            =   135
      TabIndex        =   1
      Top             =   1380
      Width           =   3135
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
      Y1              =   2550
      Y2              =   2550
   End
End
Attribute VB_Name = "frmErrorConverter"
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
'Public cParent As clsErrorConverter
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

' *** Status Codes
Private Const INVALID_HANDLE_VALUE = -1&
Private Const ERROR_SUCCESS = 0&

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdGet_Click()
    Dim i As Long
    Select Case cboType.ListIndex
    Case 0
        On Error Resume Next
        txtError = Error$(txtErr.Text)
    Case 1
        txtError = ReturnAPIError(txtErr.Text)
        If txtError = "" Then txtError.Text = "***Not Found***"
    End Select
End Sub

Private Sub Form_Load()
  '  cFunc.pLoadResStrings Controls
    cboType.ListIndex = 0
End Sub


Private Sub txtErr_Change()
    cmdGet.Enabled = IsNumeric(txtErr.Text) Or cboType.ListIndex = 2
End Sub


Public Function ReturnAPIError(ErrorCode As Long) As String
    '// Takes an API error number, and returns
    '// a descriptive text string of the error
    '// Thanks to Theirry Waty
    Dim sBuffer          As String
    
    ' *** Allocate the string, then get the system to
    ' *** tell us the error message associated with
    ' *** this error number
    
    sBuffer = String(256, 0)
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, ErrorCode, 0&, sBuffer, Len(sBuffer), 0&
    
    ' *** Strip the last null, then the last CrLf pair if it exists
    
    sBuffer = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    If Right$(sBuffer, 2) = Chr$(13) & Chr$(10) Then
       sBuffer = Mid$(sBuffer, 1, Len(sBuffer) - 2)
    End If
    
    ReturnAPIError = sBuffer

End Function
