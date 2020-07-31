VERSION 5.00
Object = "{C4925FC3-1606-11D4-82BB-004005AAE138}#5.2#0"; "VBWIML.OCX"
Object = "{1F5999A2-1D0B-11D4-82CF-004005AAE138}#6.1#0"; "VBWTBA~1.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmPreview 
   ClientHeight    =   5460
   ClientLeft      =   3060
   ClientTop       =   3345
   ClientWidth     =   6420
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPreview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   6420
   StartUpPosition =   1  'CenterOwner
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   3720
      Left            =   735
      TabIndex        =   4
      Top             =   1050
      Width           =   2790
      ExtentX         =   4921
      ExtentY         =   6562
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.PictureBox picStatus 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   6420
      TabIndex        =   2
      Top             =   5205
      Width           =   6420
   End
   Begin vbwIml.vbalImageList vbalImgLst 
      Left            =   4545
      Top             =   3375
      _ExtentX        =   953
      _ExtentY        =   953
      Size            =   7520
      Images          =   "frmPreview.frx":058A
      KeyCount        =   8
      Keys            =   "ÿÿÿÿÿÿÿ"
   End
   Begin VB.PictureBox picHolder 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
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
      Height          =   720
      Left            =   0
      ScaleHeight     =   720
      ScaleWidth      =   6420
      TabIndex        =   0
      Top             =   0
      Width           =   6420
      Begin VB.ComboBox cboAddress 
         Height          =   315
         Left            =   30
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   360
         Width           =   6135
      End
      Begin vbwTBar.cToolbar tbrMain 
         Left            =   1065
         Top             =   15
         _ExtentX        =   7408
         _ExtentY        =   582
      End
      Begin vbwTBar.cToolbarHost tbhMain 
         Height          =   345
         Left            =   45
         TabIndex        =   3
         Top             =   0
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   609
         BorderStyle     =   0
      End
   End
   Begin VB.Menu mnuSizeTOP 
      Caption         =   "&SizePop"
      Visible         =   0   'False
      Begin VB.Menu mnuSize 
         Caption         =   "536 x 196 (640 x 480, Default)"
         Index           =   0
      End
      Begin VB.Menu mnuSize 
         Caption         =   "600 x 300 (640 x 480, Maximized)"
         Index           =   1
      End
      Begin VB.Menu mnuSize 
         Caption         =   "760 x 420 (800 x 600, Maximized)"
         Index           =   2
      End
      Begin VB.Menu mnuSize 
         Caption         =   "795 x 470 (832 x 624, Maximized)"
         Index           =   3
      End
      Begin VB.Menu mnuSize 
         Caption         =   "955 x 600 (1024 x 768, Maximized)"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmPreview"
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
Private cCombo      As clsFlatCombo
Private cStatusBar  As clsStatusBar
Private bIgnore     As Boolean

Public Sub NavigateTo(sURL As String, sTitle As String)
    If InStr(1, brwWebBrowser.LocationURL, sURL & "?") = 1 Then
        'preserve query string
        sURL = brwWebBrowser.LocationURL
    End If
    If brwWebBrowser.LocationURL = sURL Then
        brwWebBrowser.Refresh
    Else
        'navigates to the specified URL
        brwWebBrowser.Navigate sURL
        'set the caption
        'Caption = LoadResString(1258) & sTitle '"IE Preview: "
    End If
End Sub

Private Sub brwWebBrowser_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
If Progress > 0 Then
'cStatusBar.PanelText("Status") = CStr((Progress / ProgressMax) * 100) & "% Complete"
End If
End Sub

Private Sub brwWebBrowser_StatusTextChange(ByVal Text As String)
    'update the status bar's text
    cStatusBar.PanelText("Status") = Text
End Sub

Private Sub brwWebBrowser_TitleChange(ByVal Text As String)
    'set form's caption
    Caption = Text ' brwWebBrowser.LocationName
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'refresh webbrowser
    If KeyCode = vbKeyF5 Then
        brwWebBrowser.Refresh
    ElseIf KeyCode = vbKeyF1 Then
        cDialog.ShowHelpTopic 12, hWnd
    End If
End Sub
Private Sub Form_Load()
    On Error GoTo ErrHandler
    'Init status bar
    Set cStatusBar = New clsStatusBar
    With cStatusBar
        'use picStatus pic box
        .Create picStatus
        'add a panel
        .AddPanel estbrStandard, "Ready", , , True, , , "Status"
        .AddPanel estbrStandard, "", , , , True, , "Size"
        'we want a size grip
        .SizeGrip = True
        'set offset
        .SetLeftTopOffsets 0, 2
        'adjust picbox height
        picStatus.Height = .Height
        'set status text
        .PanelText(1) = "Please wait..."
    End With
    'make combo flat
    Set cCombo = New clsFlatCombo
    cCombo.Attach cboAddress.hWnd

    With tbrMain
        'build the toolbar
        'set image list
        .ImageSource = CTBExternalImageList
        .SetImageList vbalImgLst, CTBImageListNormal
        'create tb
        .CreateToolbar 16, , True, False
        'no wrap
        .Wrappable = False
        'add buttons
        .AddButton "Back", 0, , , Empty, CTBNormal, "Back"
        .AddButton "Forward", 1, , , Empty, , "Forward"
        .AddButton Empty, -1, , , , CTBSeparator
        .AddButton "Stop", 2, , , Empty, , "Stop"
        .AddButton "Refresh", 3, , , Empty, , "Refresh"
        .AddButton "Home", 4, , , Empty, , "Home"
        .AddButton Empty, -1, , , , CTBSeparator
        .AddButton "Search", 5, , , Empty, , "Search"
        .AddButton "Working Offline", 7, , , Empty, CTBCheck, "Offline"
    End With
    
    With tbhMain
        'init toolbar host
        'no border
        .BorderStyle = etbhBorderStyleNone
        'set width
        .Width = ScaleWidth
        'set height
        .Height = tbrMain.ToolbarHeight * Screen.TwipsPerPixelY
        'capture main toolbar
        .Capture tbrMain
    End With
    'default to working offline
    brwWebBrowser.Offline = True
    'clear address
    cboAddress.Text = ""
    'set the caption
    Caption = LoadResString(1258)
    'resize picHolder to update tb
    picHolder_Resize
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Preview.Load"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Hide
        Cancel = -1
    End If
End Sub

Private Sub mnuSize_Click(Index As Integer)
Dim sCaption As String
Dim lWidth As Long
Dim lHeight As Long
Dim lXPos As Long
    sCaption = mnuSize(Index).Caption
    lXPos = InStr(1, sCaption, "x")
    lWidth = Left$(sCaption, lXPos - 1)
    lHeight = Mid$(sCaption, lXPos + 2, InStr(lXPos + 2, sCaption, " ") - lXPos - 2)
    Height = (lHeight * Screen.TwipsPerPixelY) + picStatus.Height + brwWebBrowser.Top + 405
    Width = (lWidth * Screen.TwipsPerPixelX) + 120
End Sub

Private Sub picStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim tR As RECT
Dim tp As POINTAPI
    GetCursorPos tp
    ScreenToClient picStatus.hWnd, tp
    cStatusBar.GetPanelRect "Size", tR
    'take away sizer bar
    tR.Right = tR.Right - 10
    If PtInRect(tR, tp.X, tp.Y) Then
        PopupMenu mnuSizeTOP
    End If
End Sub

'*** Status Bar ***
Private Sub picStatus_Paint()
    'draw statusbar
    If Not cStatusBar Is Nothing Then cStatusBar.Draw
End Sub
Private Sub picStatus_Resize()
    'draw statusbar
    picStatus_Paint
End Sub
Private Sub brwWebBrowser_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    Dim i       As Integer
    Dim bFound  As Boolean
    'loop through the combo to see if we have visited this
    'url before...
    For i = 0 To cboAddress.ListCount - 1
        If cboAddress.List(i) = brwWebBrowser.LocationURL Then
            bFound = True
            Exit For
        End If
    Next i
    'we have found it... remove it!
    If bFound Then cboAddress.RemoveItem i
    'add the URL to the combo
    '(it will appear at the top)
    cboAddress.AddItem brwWebBrowser.LocationURL, 0
    
    'ignore triggered events...
    bIgnore = True
    'select the first item
    cboAddress.ListIndex = 0
    bIgnore = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    'resize toolbar host
    tbhMain.Move 0, 0, picHolder.ScaleWidth, tbrMain.ToolbarHeight * Screen.TwipsPerPixelX
    'resize the webbrowser
    brwWebBrowser.Move 0, picHolder.Height, ScaleWidth, ScaleHeight - picHolder.Height - picStatus.Height
    cStatusBar.PanelText("Size") = brwWebBrowser.Width / Screen.TwipsPerPixelX & " x " & brwWebBrowser.Height / Screen.TwipsPerPixelY
End Sub
Private Sub picHolder_Resize()
On Error Resume Next
    'resize toolbar host
    tbhMain.Width = picHolder.ScaleWidth
    'adjust holder and address pos
    cboAddress.Move 0, tbrMain.ToolbarHeight2 * Screen.TwipsPerPixelX, ScaleWidth ' - 20
End Sub
Private Sub cboAddress_Click()
    'ignore flag set, abort
    If bIgnore Then Exit Sub
    'navigate to the selected url
    brwWebBrowser.Navigate cboAddress.Text
End Sub
Private Sub cboAddress_KeyPress(KeyAscii As Integer)
    'go to the entered url
    If KeyAscii = vbKeyReturn Then cboAddress_Click
End Sub
'*** Toolbar ***
Private Sub tbrMain_ButtonClick(ByVal lButton As Long)
On Error Resume Next
    'get the clicked toolbar
    Select Case tbrMain.ButtonKey(lButton)
    'and perform correct option
    Case "Back"
        brwWebBrowser.GoBack
    Case "Forward"
        brwWebBrowser.GoForward
    Case "Refresh"
        brwWebBrowser.Refresh
    Case "Home"
        brwWebBrowser.GoHome
    Case "Search"
        brwWebBrowser.GoSearch
    Case "Stop"
        brwWebBrowser.Stop
    Case "Offline"
        'reverse offline/online state
        brwWebBrowser.Offline = Not brwWebBrowser.Offline
        'update toolbar image and tooltip
        tbrMain.ButtonImage("Offline") = IIf(brwWebBrowser.Offline, 7, 6)
        tbrMain.ButtonToolTip("Offline") = IIf(brwWebBrowser.Offline, "Working Offline", "Working Online")
    End Select
End Sub

