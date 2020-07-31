VERSION 5.00
Begin VB.Form frmDTConv 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DataType Converter"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5505
   Icon            =   "frmConvertNums.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboConvertFrom 
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
      ItemData        =   "frmConvertNums.frx":000C
      Left            =   1800
      List            =   "frmConvertNums.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   120
      Width           =   2295
   End
   Begin VB.ComboBox cboDataType 
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
      ItemData        =   "frmConvertNums.frx":0032
      Left            =   1800
      List            =   "frmConvertNums.frx":005D
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   495
      Width           =   2295
   End
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "frmConvertNums.frx":00CA
      Top             =   960
      Width           =   3975
   End
   Begin VB.TextBox txtResult 
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
      Left            =   1560
      TabIndex        =   5
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox txtText 
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
      Left            =   1560
      TabIndex        =   3
      Top             =   2160
      Width           =   2535
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4320
      TabIndex        =   1
      Top             =   510
      Width           =   1092
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "&Convert"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4320
      TabIndex        =   0
      Top             =   120
      Width           =   1092
   End
   Begin VB.Label lblLabel 
      Caption         =   "Convert from"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblLabel 
      Caption         =   "Convert to"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label lblLabel 
      Caption         =   "Result"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lblLabel 
      Caption         =   "Text To Convert"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   1335
   End
End
Attribute VB_Name = "frmDTConv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////////////
'//                This Code is Copyright VB Web 1999                   //
'//         You MAY NOT re-distribute this source code. Instead,        //
'//                     please provide a link to                        //
'//       http://www.vbweb.co.uk/dev/?devpad.htm                        //
'//                                                                     //
'//       If you would like to become a beta tester, please email       //
'//   devpadbeta@vbweb.f9.co.uk with your Name, VB Version and PC Spec  //
'//                                                                     //
'//           Please report any bugs to bugs@vbweb.f9.co.uk             //
'/////////////////////////////////////////////////////////////////////////
Option Explicit
'Public cParent As clsDTConv
Private sDescriptions(0 To 12) As String

Private Sub cboConvertFrom_Click()
    Select Case cboConvertFrom.ListIndex
    Case 1 'c++ hex
        cboDataType.ListIndex = 0
        cboDataType.Enabled = False
    Case Else
        cboDataType.Enabled = True
    End Select
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
'// Convert as required
Private Sub cmdConvert_Click()
On Error GoTo ErrHandler
    Select Case cboConvertFrom.ListIndex
    Case 1 'c++ hex
        txtResult.Text = "&H" & CLng(Right$(txtText.Text, Len(txtText.Text) - 2))
    Case Else
        Select Case cboDataType.ListIndex
        Case 0 '// Hex
            txtResult.Text = "&H" & LCase(Hex$(txtText.Text))
        Case 1 '// Byte
            txtResult.Text = CByte(txtText.Text)
        Case 2 '// Integer
            txtResult.Text = CInt(txtText.Text)
        Case 3 '// Long
            txtResult.Text = CLng(txtText.Text)
        Case 4 '// Single
            txtResult.Text = CSng(txtText.Text)
        Case 5 '// Double
            txtResult.Text = CDbl(txtText.Text)
        Case 6 '// Octal
            txtResult.Text = Oct$(txtText.Text)
        Case 7 '// Boolean
            txtResult.Text = CBool(txtText.Text)
        Case 8 '// Currency
            txtResult.Text = CCur(txtText.Text)
        Case 9 '// Number
            txtResult.Text = Val(txtText.Text)
        Case 10 '// Decimal
            txtResult.Text = CDec(txtText.Text)
        Case 11 '// Date
            txtResult.Text = CDate(txtText.Text)
        Case 12 '// c++ hex
            txtResult.Text = "0x" & Format(Right$(txtText.Text, Len(txtText.Text) - 2), "0000")
        End Select
    End Select
    
    Exit Sub
ErrHandler:
    Select Case Err
    Case 6
        cFunc.ErrHandler 6, "Overflow: Please enter a smaller value", "DataType.Convert", , Error
    Case 13
        cFunc.ErrHandler 13, "Please enter a valid value", "DataType.Convert", , Error
    Case Else
        cFunc.ErrHandler Err, Error, "DataType.Convert"
    End Select
End Sub
'// Set information values
Private Sub Form_Load()
On Error Resume Next
    'cFunc.pLoadResStrings Controls
    sDescriptions(0) = "You can convert a C++ Hex code (ie 0x003) to a standard hex code (ie &H3), or vice versa, or convert a hex code to an integer"
    sDescriptions(12) = sDescriptions(0)
    sDescriptions(1) = "Storage Size: 1 byte" & vbCrLf & "Range: 0 To 255"
    sDescriptions(2) = "Storage Size: 2 bytes" & vbCrLf & "Range: -32,768 To 32,767"
    sDescriptions(3) = "Storage Size: 4 bytes" & vbCrLf & "Range: -2,147,483,648 to 2,147,483,647"
    sDescriptions(4) = "Storage Size: 4 byte" & vbCrLf & "Range: -3.402823E38 to -1.401298E-45 for negative values; 1.401298E-45 to 3.402823E38 for positive values"
    sDescriptions(5) = "Storage Size: 8 byte" & vbCrLf & "Range: -1.79769313486232E308 to -4.94065645841247E-324 for negative values; 4.94065645841247E-324 to 1.79769313486232E308 for positive values"
    
    
    sDescriptions(7) = "Storage Size: 2 bytes" & vbCrLf & "Range: True or False"
    sDescriptions(8) = "Storage Size: 8 byte" & vbCrLf & "Range: -922,337,203,685,477.5808 to 922,337,203,685,477.5807"
    sDescriptions(9) = "Any numeric value"
    
    
    sDescriptions(10) = "Storage Size: 14 bytes" & vbCrLf & "Range: +/-79,228,162,514,264,337,593,543,950,335 with no decimal point; +/-7.9228162514264337593543950335 with 28 places to the right of the decimal; smallest non-zero number is +/-0.0000000000000000000000000001"
    sDescriptions(11) = "Storage Size: 8 bytes" & vbCrLf & "Range: January 1, 100 to December 31, 9999"
   ' txtText = ActiveForm.txtText.SelText
    cboDataType.ListIndex = GetSetting(App.Title, "Settings", "DataType", 0)
    UpdateInfo
End Sub

'// Update Information text box when options change
Private Sub cboDataType_Change()
    UpdateInfo
End Sub

Private Sub cboDataType_Click()
    UpdateInfo
End Sub

Private Sub UpdateInfo()
    txtInfo.Text = sDescriptions(cboDataType.ListIndex)
    If txtInfo.Text = "" Then txtInfo.Text = "No Information"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "Settings", "DataType", cboDataType.ListIndex
End Sub
