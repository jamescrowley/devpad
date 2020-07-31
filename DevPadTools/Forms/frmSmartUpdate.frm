VERSION 5.00
Begin VB.Form frmSmartUpdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Smart Update"
   ClientHeight    =   4020
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
   Icon            =   "frmSmartUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPages 
      BorderStyle     =   0  'None
      Height          =   2265
      Index           =   1
      Left            =   105
      ScaleHeight     =   2265
      ScaleWidth      =   4905
      TabIndex        =   18
      Top             =   1050
      Width           =   4905
      Begin VB.CommandButton cmdDownload 
         Caption         =   "Download Update"
         Enabled         =   0   'False
         Height          =   345
         Left            =   1522
         TabIndex        =   25
         Top             =   1830
         Width           =   1890
      End
      Begin VB.Image imgCurrent 
         Height          =   210
         Left            =   3855
         Picture         =   "frmSmartUpdate.frx":000C
         Top             =   750
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label lblResult 
         Height          =   280
         Left            =   75
         TabIndex        =   24
         Top             =   1365
         Width           =   4590
      End
      Begin VB.Image imgStatus 
         Height          =   210
         Index           =   2
         Left            =   60
         Top             =   1020
         Width           =   210
      End
      Begin VB.Label lblItem 
         Caption         =   "Checking versions"
         Height          =   285
         Index           =   3
         Left            =   360
         TabIndex        =   23
         Top             =   1020
         Width           =   2985
      End
      Begin VB.Image imgStatus 
         Height          =   210
         Index           =   1
         Left            =   60
         Top             =   735
         Width           =   210
      End
      Begin VB.Label lblItem 
         Caption         =   "Downloading update information"
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   22
         Top             =   735
         Width           =   2985
      End
      Begin VB.Image imgStatus 
         Height          =   210
         Index           =   0
         Left            =   60
         Top             =   450
         Width           =   210
      End
      Begin VB.Image imgCross 
         Height          =   210
         Left            =   3840
         Picture         =   "frmSmartUpdate.frx":005A
         Top             =   465
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Image imgTick 
         Height          =   210
         Left            =   3480
         Picture         =   "frmSmartUpdate.frx":00FD
         Top             =   450
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label lblItem 
         Caption         =   "Connecting to server"
         Height          =   280
         Index           =   0
         Left            =   360
         TabIndex        =   20
         Top             =   450
         Width           =   2985
      End
      Begin VB.Label lblLabel 
         Caption         =   "Please wait while the wizard checks for updates"
         Height          =   300
         Index           =   3
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   4935
      End
   End
   Begin VB.PictureBox picPages 
      BorderStyle     =   0  'None
      Height          =   2265
      Index           =   0
      Left            =   105
      ScaleHeight     =   2265
      ScaleMode       =   0  'User
      ScaleWidth      =   4875.183
      TabIndex        =   15
      Top             =   1080
      Width           =   4905
      Begin VB.Label lblLabel 
         Caption         =   "To check for an updated release of Developers Pad, connect to the internet, and then click Next."
         Height          =   555
         Index           =   1
         Left            =   15
         TabIndex        =   17
         Top             =   525
         Width           =   4785
      End
      Begin VB.Label lblLabel 
         Caption         =   "Welcome to the Developers Pad Smart Update Wizard"
         Height          =   300
         Index           =   0
         Left            =   0
         TabIndex        =   16
         Top             =   -15
         Width           =   4935
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >"
      Default         =   -1  'True
      Height          =   345
      Left            =   2460
      TabIndex        =   14
      Top             =   3600
      Width           =   1200
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Height          =   345
      Left            =   1200
      TabIndex        =   13
      Top             =   3600
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3855
      TabIndex        =   12
      Top             =   3600
      Width           =   1200
   End
   Begin VB.PictureBox picSpy 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
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
      Height          =   1410
      Left            =   0
      ScaleHeight     =   1410
      ScaleWidth      =   5145
      TabIndex        =   0
      Top             =   4000
      Width           =   5145
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Style:"
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
         Index           =   9
         Left            =   150
         TabIndex        =   8
         Top             =   945
         Width           =   825
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Caption:"
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
         Index           =   10
         Left            =   150
         TabIndex        =   7
         Top             =   375
         Width           =   825
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Class:"
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
         Index           =   11
         Left            =   150
         TabIndex        =   6
         Top             =   660
         Width           =   825
      End
      Begin VB.Label lblLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "hWnd:"
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
         Index           =   12
         Left            =   150
         TabIndex        =   5
         Top             =   105
         Width           =   825
      End
      Begin VB.Shape shpBorder 
         BorderColor     =   &H00C0C0C0&
         Height          =   1410
         Left            =   0
         Top             =   0
         Width           =   5145
      End
      Begin VB.Label lblStyle 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1080
         TabIndex        =   4
         Top             =   945
         Width           =   3870
      End
      Begin VB.Label lblhWnd 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   105
         Width           =   3870
      End
      Begin VB.Label lblCaption 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   375
         Width           =   3870
      End
      Begin VB.Label lblClass 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   660
         Width           =   3870
      End
   End
   Begin VB.Label lblItem 
      Caption         =   "Connecting to server"
      Height          =   285
      Index           =   1
      Left            =   465
      TabIndex        =   21
      Top             =   1800
      Width           =   2985
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      Caption         =   "&Help"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   105
      MouseIcon       =   "frmSmartUpdate.frx":0192
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   3645
      Width           =   360
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   4
      X1              =   120
      X2              =   5055
      Y1              =   3480
      Y2              =   3480
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
      Picture         =   "frmSmartUpdate.frx":049C
      Top             =   165
      Width           =   480
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Smart Update"
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
      TabIndex        =   10
      Top             =   75
      Width           =   1515
   End
   Begin VB.Label lblMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "Checks for Developers Pad updates"
      Height          =   255
      Left            =   300
      TabIndex        =   9
      Top             =   345
      Width           =   3825
   End
   Begin VB.Shape shpShape 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   900
      Left            =   0
      Top             =   0
      Width           =   5145
   End
   Begin VB.Line linLine 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   30
      X2              =   7500
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Line linLine 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   5
      X1              =   105
      X2              =   5010
      Y1              =   3495
      Y2              =   3495
   End
End
Attribute VB_Name = "frmSmartUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum REQUEST_STATUS
    rsConnecting = 1
    rsGettingData = 2
    rsParsing = 3
    rsIdle = 0
    rsCancelled = -1
End Enum
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private WithEvents cXML As DOMDocument
Attribute cXML.VB_VarHelpID = -1
Private vStatus         As REQUEST_STATUS
Private nCurrentPage    As Integer

Private Sub cmdBack_Click()
    imgStatus(0).Picture = Nothing
    imgStatus(1).Picture = Nothing
    imgStatus(2).Picture = Nothing
    lblResult.Caption = ""
    cmdDownload.Enabled = False
    pShowPage nCurrentPage - 1
End Sub
Private Sub cmdNext_Click()
    pShowPage nCurrentPage + 1
    Select Case nCurrentPage
    Case 1
        'perform update check
        pStartUpdateCheck
    End Select
End Sub

Private Sub Form_Load()
    'init XML object
    Set cXML = New DOMDocument
    'show the first page
    pShowPage (0)
End Sub
Private Sub pShowPage(nPage As Integer)
    'show the page
    picPages(nPage).ZOrder
    picPages(nPage).Visible = True
    'disable < back?
    cmdBack.Enabled = Not (nPage = 0)
    cmdNext.Enabled = Not (nPage = 1)
    nCurrentPage = nPage
End Sub
Private Sub cmdDownload_Click()
    ShellExecute hwnd, vbNullString, "http://www.developerspad.com/download/", vbNullString, vbNullString, vbNormalFocus
End Sub
Private Sub cmdCancel_Click()
    If vStatus <> rsIdle Then
        vStatus = rsCancelled
        cXML.abort
    Else
        Unload Me
    End If
End Sub

Private Sub pStartUpdateCheck()
    'perform update check
    cmdBack.Enabled = False
    imgStatus(0).Picture = imgCurrent.Picture
    DoEvents
    vStatus = rsConnecting
    cXML.async = True 'Asyncronous
    cXML.Load ("http://www.developerspad.com/version.xml")
End Sub

Private Sub cXML_onreadystatechange()
    If vStatus <> rsIdle Then
        Select Case cXML.readyState
        Case 1, 2
            vStatus = rsConnecting
        Case 3
            imgStatus(0).Picture = imgTick.Picture
            imgStatus(1).Picture = imgCurrent.Picture
            vStatus = rsGettingData
        Case 4
            If cXML.parseError.ErrorCode <> 0 Then
                If vStatus = rsCancelled Then
                    lblResult.Caption = "Cancelled"
                Else
                    lblResult.Caption = cXML.parseError.reason
                End If
                imgStatus(0).Picture = imgCross.Picture
                imgStatus(1).Picture = imgCross.Picture
                imgStatus(2).Picture = imgCross.Picture
                
            Else
                imgStatus(1).Picture = imgTick.Picture
                imgStatus(2).Picture = imgCurrent.Picture
                vStatus = rsParsing
                pCheckVersion
            End If
            cmdBack.Enabled = True
            vStatus = rsIdle
        End Select
    End If
End Sub
Private Sub pCheckVersion()
    Dim oNodeList   As IXMLDOMNodeList
    Dim oNode       As IXMLDOMNode
    Dim oSubNode    As IXMLDOMNode
    Dim sName       As String
    Dim sVersion    As String
    Dim sDescription As String
    Dim bNewVersion As Boolean

    Set oNodeList = cXML.documentElement.childNodes
    For Each oNode In oNodeList
        If oNode.nodeName = "item" Then
        
            'get it's name
            sName = oNode.Attributes(0).nodeValue
            For Each oSubNode In oNode.childNodes
                If oSubNode.nodeName = "version" Then
                    sVersion = oSubNode.nodeTypedValue
                ElseIf oSubNode.nodeName = "description" Then
                    sDescription = oSubNode.nodeTypedValue
                End If
            Next
            'check version
            If sVersion <> App.Major & "." & App.Minor & "." & App.Revision Then
                'new version
                bNewVersion = True
                'txtDetails.Text = txtDetails.Text & "---" & sName & "---" & vbCrLf & "Version " & sVersion & vbCrLf & vbCrLf & sDescription & vbCrLf & vbCrLf
            End If
        End If
    Next
    If bNewVersion Then
        lblResult.Caption = "Update(s) are available. "
    Else
        lblResult.Caption = "No new updates are available at this time."
    End If
    cmdDownload.Enabled = (bNewVersion)
    imgStatus(2).Picture = imgTick.Picture
End Sub

