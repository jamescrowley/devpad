VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search CodeHound.com"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search..."
      Default         =   -1  'True
      Height          =   345
      Left            =   3735
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   345
      Left            =   3735
      TabIndex        =   6
      Top             =   1890
      Width           =   1215
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   1350
      Width           =   1935
   End
   Begin VB.ComboBox cboCat 
      Height          =   315
      ItemData        =   "frmSearch.frx":000C
      Left            =   1680
      List            =   "frmSearch.frx":0022
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   990
      Width           =   1035
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      Caption         =   "&Help"
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   105
      MouseIcon       =   "frmSearch.frx":0047
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1935
      Width           =   360
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   4
      X1              =   60
      X2              =   4995
      Y1              =   1755
      Y2              =   1755
   End
   Begin VB.Label lblLabel 
      Caption         =   "Searching for..."
      Height          =   255
      Index           =   0
      Left            =   165
      TabIndex        =   4
      Top             =   1380
      Width           =   1335
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Search the web for code using codehound.com"
      Height          =   255
      Left            =   330
      TabIndex        =   3
      Top             =   375
      Width           =   4080
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Find Code"
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
      TabIndex        =   2
      Top             =   75
      Width           =   1830
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   4470
      Picture         =   "frmSearch.frx":0351
      Top             =   165
      Width           =   480
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   6
      X1              =   30
      X2              =   7500
      Y1              =   885
      Y2              =   885
   End
   Begin VB.Label lblLabel 
      Caption         =   "Category"
      Height          =   255
      Index           =   11
      Left            =   165
      TabIndex        =   1
      Top             =   1020
      Width           =   1335
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
      Index           =   7
      X1              =   30
      X2              =   7500
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   5
      X1              =   45
      X2              =   4950
      Y1              =   1770
      Y2              =   1770
   End
End
Attribute VB_Name = "frmSearch"
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

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    Dim sURL As String
    sURL = "http://www.codehound.com/"
    sURL = sURL & LCase$(cboCat.Text) & "/"
    sURL = sURL & "results/results.asp?Q=" & URLEncode(txtSearch.Text) & "&S=1&V=567"
    ShellExecute 0&, vbNullString, sURL, vbNullString, vbNullString, vbNormalFocus
End Sub

Private Sub Form_Load()
    cboCat.ListIndex = GetSetting(App.Title, "Settings", "SearchCat", 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "Settings", "SearchCat", cboCat.ListIndex
End Sub
Private Function URLEncode(str As String) As String

Dim strTemp As String
Dim strChar As String
Dim nTemp As Integer
Dim nAsciiVal As Integer

  strTemp = ""
  strChar = ""

  For nTemp = 1 To Len(str)
    nAsciiVal = Asc(Mid(str, nTemp, 1))
    If ((nAsciiVal < 123) And (nAsciiVal > 96)) Then
      strTemp = strTemp & Chr(nAsciiVal)
    ElseIf ((nAsciiVal < 91) And (nAsciiVal > 64)) Then
      strTemp = strTemp & Chr(nAsciiVal)
    ElseIf ((nAsciiVal < 58) And (nAsciiVal > 47)) Then
      strTemp = strTemp & Chr(nAsciiVal)
    Else
      strChar = Trim(Hex(nAsciiVal))
      If nAsciiVal < 16 Then
        strTemp = strTemp & "%0" & strChar
      Else
        strTemp = strTemp & "%" & strChar
      End If
    End If
  Next
  
  URLEncode = strTemp
        
End Function
