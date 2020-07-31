VERSION 5.00
Begin VB.Form frmInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Document Info"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5100
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   5100
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboState 
      Height          =   315
      ItemData        =   "frmInfo.frx":000C
      Left            =   1680
      List            =   "frmInfo.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3540
      Width           =   1170
   End
   Begin VB.ComboBox cboForms 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   2955
   End
   Begin VB.ComboBox cboSpeed 
      Height          =   315
      ItemData        =   "frmInfo.frx":002A
      Left            =   3105
      List            =   "frmInfo.frx":0046
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   4680
      Width           =   1905
   End
   Begin VB.ComboBox cboSize 
      Height          =   315
      ItemData        =   "frmInfo.frx":0095
      Left            =   1065
      List            =   "frmInfo.frx":00A2
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   4680
      Width           =   840
   End
   Begin VB.CommandButton cmdProperties 
      Caption         =   "&Properties..."
      Height          =   345
      Left            =   2520
      TabIndex        =   12
      Top             =   5205
      Width           =   1215
   End
   Begin VB.TextBox txtDownloadTime 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "0"
      Top             =   4275
      Width           =   2655
   End
   Begin VB.TextBox txtDocType 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0"
      Top             =   2430
      Width           =   2655
   End
   Begin VB.TextBox txtSize 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "0"
      Top             =   3960
      Width           =   2655
   End
   Begin VB.TextBox txtLines 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "0"
      Top             =   2085
      Width           =   2655
   End
   Begin VB.TextBox txtChars 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "0"
      Top             =   1740
      Width           =   2655
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   345
      Left            =   3825
      TabIndex        =   13
      Top             =   5205
      Width           =   1215
   End
   Begin VB.TextBox txtFullPath 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3105
      Width           =   3390
   End
   Begin VB.TextBox txtFilename 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2760
      Width           =   3390
   End
   Begin VB.Image imgRefresh 
      Height          =   240
      Left            =   4710
      Picture         =   "frmInfo.frx":00B5
      ToolTipText     =   "Refresh"
      Top             =   1005
      Width           =   240
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   8
      X1              =   75
      X2              =   4980
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label lblLabel 
      Caption         =   "Document"
      Height          =   255
      Index           =   11
      Left            =   165
      TabIndex        =   27
      Top             =   990
      Width           =   1335
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
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   4470
      Picture         =   "frmInfo.frx":01FF
      Top             =   165
      Width           =   480
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Document Info"
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
      TabIndex        =   26
      Top             =   75
      Width           =   1830
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Get download times for your documents!"
      Height          =   255
      Left            =   330
      TabIndex        =   25
      Top             =   375
      Width           =   4080
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   2
      X1              =   75
      X2              =   5010
      Y1              =   4575
      Y2              =   4575
   End
   Begin VB.Label Label2 
      Caption         =   "Modem Speed"
      Height          =   210
      Left            =   1980
      TabIndex        =   24
      Top             =   4740
      Width           =   1140
   End
   Begin VB.Label lblLabel 
      Caption         =   "Download Times"
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
      Index           =   10
      Left            =   150
      TabIndex        =   23
      Top             =   3585
      Width           =   1935
   End
   Begin VB.Label lblLabel 
      Caption         =   "File Information"
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
      Index           =   9
      Left            =   165
      TabIndex        =   22
      Top             =   1410
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Size Units"
      Height          =   210
      Left            =   165
      TabIndex        =   21
      Top             =   4725
      Width           =   885
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   0
      X1              =   75
      X2              =   5010
      Y1              =   3435
      Y2              =   3435
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      Caption         =   "&Help"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   165
      MouseIcon       =   "frmInfo.frx":0509
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   5265
      Width           =   360
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   4
      X1              =   75
      X2              =   5055
      Y1              =   5100
      Y2              =   5100
   End
   Begin VB.Label lblLabel 
      Caption         =   "Download Time"
      Height          =   255
      Index           =   8
      Left            =   150
      TabIndex        =   20
      Top             =   4260
      Width           =   1335
   End
   Begin VB.Label lblLabel 
      Caption         =   "Document Type"
      Height          =   255
      Index           =   7
      Left            =   165
      TabIndex        =   19
      Top             =   2430
      Width           =   1335
   End
   Begin VB.Label lblLabel 
      Caption         =   "Full Path"
      Height          =   255
      Index           =   4
      Left            =   165
      TabIndex        =   18
      Top             =   3090
      Width           =   855
   End
   Begin VB.Label lblLabel 
      Caption         =   "Filename"
      Height          =   255
      Index           =   5
      Left            =   165
      TabIndex        =   17
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label lblLabel 
      Caption         =   "Size"
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   16
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label lblLabel 
      Caption         =   "Lines"
      Height          =   255
      Index           =   2
      Left            =   165
      TabIndex        =   15
      Top             =   2085
      Width           =   1335
   End
   Begin VB.Label lblLabel 
      Caption         =   "Characters"
      Height          =   255
      Index           =   3
      Left            =   165
      TabIndex        =   14
      Top             =   1740
      Width           =   1335
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   5
      X1              =   75
      X2              =   5010
      Y1              =   5115
      Y2              =   5115
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   1
      X1              =   75
      X2              =   4965
      Y1              =   3450
      Y2              =   3450
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   3
      X1              =   75
      X2              =   4965
      Y1              =   4590
      Y2              =   4590
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
      Index           =   9
      X1              =   75
      X2              =   4950
      Y1              =   1335
      Y2              =   1335
   End
   Begin VB.Menu mnuSpeed 
      Caption         =   "Modem Speed"
      Visible         =   0   'False
      Begin VB.Menu mnuSpeed14K 
         Caption         =   "14.4K"
      End
      Begin VB.Menu mnuSpeed28K 
         Caption         =   "28.8K"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSpeed56K 
         Caption         =   "56.6K"
      End
      Begin VB.Menu mnuSpeedISDNSingle 
         Caption         =   "ISDN Single Channel"
      End
      Begin VB.Menu mnuSpeedISDNDual 
         Caption         =   "ISDN Dual Channel"
      End
      Begin VB.Menu mnuSpeedASDL 
         Caption         =   "ASDL"
      End
      Begin VB.Menu mnuSpeedFullT1 
         Caption         =   "Full T1"
      End
      Begin VB.Menu mnuSpeedT3 
         Caption         =   "T3"
      End
   End
   Begin VB.Menu mnuSize 
      Caption         =   "Size"
      Visible         =   0   'False
      Begin VB.Menu mnuSizeBytes 
         Caption         =   "Bytes"
      End
      Begin VB.Menu mnuSizeKB 
         Caption         =   "KB"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSizeMB 
         Caption         =   "MB"
      End
   End
End
Attribute VB_Name = "frmInfo"
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
'Public cParent As clsDocInfo
Private Type SHELLEXECUTEINFO
        cbSize As Long
        fMask As Long
        hWnd As Long
        lpVerb As String
        lpFile As String
        lpParameters As String
        lpDirectory As String
        nShow As Long
        hInstApp As Long
        '  Optional fields
        lpIDList As Long
        lpClass As String
        hkeyClass As Long
        dwHotKey As Long
        hIcon As Long
        hProcess As Long
End Type

Private Const SEE_MASK_INVOKEDLIST = &HC
Private Declare Function ShellExecuteEx Lib "shell32" (lpSEI As SHELLEXECUTEINFO) As Long

Private lByteCount As Long
Private lSavedByteCount As Long

Private lSpeeds(0 To 7) As Long
Private frmActive As IDevPadDocument
Private bIgnore As Boolean

'// Cancel
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cboForms_Click()
    If cboForms.ListIndex <> -1 Then
        Set frmActive = cFunc.Documents.ItemByID(cboForms.ItemData(cboForms.ListIndex))
        If frmActive Is Nothing Then RefreshList
    End If
End Sub

Private Sub cboForms_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then RefreshList
End Sub

Private Sub cboSize_Click()
    CalculateSize
End Sub

Private Sub cboSpeed_Click()
    If bIgnore = False Then CalculateDownloadTime
End Sub

Private Sub cboState_Click()
    If bIgnore = False Then
        CalculateDownloadTime
        CalculateSize
    End If
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdProperties_Click()
    ShowFileProperties txtFullPath.Text, hWnd
End Sub

Private Sub Command1_Click()
    frmCOMInfo.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then RefreshList
End Sub
Private Sub Form_Load()
    RefreshList
End Sub

Private Sub RefreshList()
    Dim i As Long
    Dim sLastItem As String
    On Error GoTo ErrHandler
    bIgnore = True
    sLastItem = cboForms.Text
    cboForms.Clear
    cboForms.ListIndex = -1
    If cFunc.Documents.Count = 0 Then Exit Sub
    For i = 1 To cFunc.Documents.Count  ' - 1
        cboForms.AddItem cFunc.Documents.Item(i).DocumentCaption
        cboForms.ItemData(cboForms.NewIndex) = cFunc.Documents.Item(i).DocID
    Next
    On Error Resume Next
    cboForms.Text = sLastItem
    If Err Then cboForms.Text = cFunc.Documents.ActiveDoc.DocumentCaption

    cboSize.ListIndex = GetSetting(App.Title, "Settings", "DefaultSize", 1)
    cboSpeed.ListIndex = GetSetting(App.Title, "Settings", "DefaultSpeed", 0)
    cboState.ListIndex = GetSetting(App.Title, "Settings", "DefaultState", 0)
    If cboState.ListIndex = 0 Then
        'for saved documents only..
        'check it is saved
        If frmActive.Saved = False Then cboState.ListIndex = 1
    End If
    
    GetData
    
    bIgnore = False
    Exit Sub
ErrHandler:
    cFunc.Dialogs.ErrHandler Err, Error, "Info.Load"
End Sub
Private Sub GetData()
    With frmActive
        '// Get the number of characters, lines and document type
        txtChars = .CharacterCount
        'txtLines = .Linlinecount ' SendMessage(.hWnd, EM_GETLINECOUNT, 0&, 0&)
        txtDocType = .FileMode & " File"
        '// Set the speeds
        lSpeeds(0) = "14400" '14K
        lSpeeds(1) = "28800" '28K
        lSpeeds(2) = "56600" '56K
        lSpeeds(3) = "64000" 'IDSN Single
        lSpeeds(4) = "128000" 'IDSN Dual
        lSpeeds(5) = "256000" '1500000 eventually! - ADSL
        lSpeeds(6) = "1544000" 'T1
        lSpeeds(7) = "4000000" 'T3
        '// check the last used units
        'CheckMenu GetSetting(App.Title, "Settings", "DefaultSize", "mnuSizeKB"), "mnuSize"
        
        
        '// set the file attributes
        txtFilename = .DocumentCaption
        txtFullPath = .FileName
        
    End With
    CalculateSize
    CalculateDownloadTime
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    '// Save the checked speed and size
    'Dim ctlControl As Control
    SaveSetting App.Title, "Settings", "DefaultSize", cboSize.ListIndex
    SaveSetting App.Title, "Settings", "DefaultSpeed", cboSpeed.ListIndex
    SaveSetting App.Title, "Settings", "DefaultState", cboState.ListIndex
End Sub


Private Sub CalculateSize()
    If frmActive Is Nothing Then Exit Sub
    '// get the current number of characters in document
    With frmActive
        If .IsRTF Then
            lByteCount = (Len(.RichText))
        Else
            lByteCount = (txtChars.Text)
        End If
    End With
    If cboState.ListIndex = 0 Then 'saved
        With frmActive
            '// calculate the saved filesize
            If .Saved = True Then
                lSavedByteCount = FileLen(.FileName)
                GetCorrectSizeUnits lSavedByteCount, txtSize
            Else
                txtSize = "File Not Saved"
            End If
            cmdProperties.Enabled = .Saved
        End With
    Else
        '// calculate it
        GetCorrectSizeUnits lByteCount, txtSize
        cmdProperties.Enabled = frmActive.Saved
    End If
End Sub
Private Sub GetCorrectSizeUnits(lBytes As Long, txtText As TextBox)
    If cboSize.Text = "KB" Then
        txtText = Round(lBytes / 1024, 2) & " KB"
    ElseIf cboSize.Text = "MB" Then
        txtText = Round(lBytes / 1048576, 2) & " MB"
    Else
        txtText = lBytes & " Bytes"
    End If
End Sub

Private Sub CalculateDownloadTime()
    If frmActive Is Nothing Or cboSpeed.ListIndex = -1 Then Exit Sub
    'Dim ctlControl As Control
    Dim lSpeed As Long
    Dim sSpeed As String
    Dim lTime As Long
    On Error GoTo CalcError

    lSpeed = lSpeeds(cboSpeed.ListIndex)
    lSpeed = lSpeed / 8
    sSpeed = cboSpeed.Text
    If cboState.ListIndex = 0 Then 'saved
        ' calculate the saved file's download time
        lTime = lSavedByteCount / lSpeed
    Else
        ' calculate the unsaved file's download time
        lTime = lByteCount / lSpeed
    End If
    ' calculate the current download time (unsaved)
    txtDownloadTime = CalculateDownloadUnits(lTime, lSpeed, sSpeed)
    Exit Sub
CalcError:
    'pErrHandler Err, Error, "DocInfo.CalcDownloadTime", , "Speed=" & lSpeed & " Type=" & sSpeed & " Time=" & lTime & " Bytes=" & lByteCount
End Sub
Private Function CalculateDownloadUnits(lTime As Long, lSpeed As Long, sSpeed As String) As String
    Dim sMeas As String
    If lTime >= 60 Then
        '// convert to minutes
        lTime = CLng(lTime / 60)
        '// convert to hours
        If lTime >= 60 Then
            lTime = CLng(lTime / 60)
            sMeas = " hour" & IIf(lTime = 1, "", "s") & " @ " & sSpeed
        Else
            sMeas = " minute" & IIf(lTime = 1, "", "s") & " @ " & sSpeed
        End If
    Else
        '// seconds is OK
        sMeas = " second" & IIf(lTime = 1, "", "s") & " @ " & sSpeed
    End If
    CalculateDownloadUnits = lTime & sMeas
End Function


'*********************************************************
'Shows the file properties dialog
'*********************************************************
Private Sub ShowFileProperties(lsFile As String, llWin As Long)
    Dim sei As SHELLEXECUTEINFO
    sei.hWnd = llWin
    sei.lpVerb = "properties"
    sei.lpFile = lsFile
    sei.fMask = SEE_MASK_INVOKEDLIST
    sei.cbSize = Len(sei)
    ShellExecuteEx sei
End Sub

Private Sub imgRefresh_Click()
    RefreshList
End Sub
