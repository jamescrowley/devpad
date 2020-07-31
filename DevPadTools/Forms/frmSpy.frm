VERSION 5.00
Begin VB.Form frmSpy 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   ClientHeight    =   1305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5145
   ControlBox      =   0   'False
   FillColor       =   &H00E0E0E0&
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   1305
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblClass 
      BackStyle       =   0  'Transparent
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
      Left            =   1080
      TabIndex        =   7
      Top             =   660
      Width           =   3870
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
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
      Left            =   1080
      TabIndex        =   6
      Top             =   375
      Width           =   3870
   End
   Begin VB.Label lblhWnd 
      BackStyle       =   0  'Transparent
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
      Left            =   1080
      TabIndex        =   5
      Top             =   105
      Width           =   3870
   End
   Begin VB.Label lblStyle 
      BackStyle       =   0  'Transparent
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
      Left            =   1080
      TabIndex        =   4
      Top             =   945
      Width           =   3870
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00C0C0C0&
      Height          =   1305
      Left            =   0
      Top             =   0
      Width           =   5145
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H80000018&
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
      Index           =   0
      Left            =   150
      TabIndex        =   3
      Top             =   105
      Width           =   825
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H80000018&
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
      Index           =   1
      Left            =   150
      TabIndex        =   2
      Top             =   660
      Width           =   825
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H80000018&
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
      Index           =   2
      Left            =   150
      TabIndex        =   1
      Top             =   375
      Width           =   825
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H80000018&
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
      Index           =   3
      Left            =   150
      TabIndex        =   0
      Top             =   945
      Width           =   825
   End
End
Attribute VB_Name = "frmSpy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'
'Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'
'Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
''Public cParent As clsClassSpy
'Private Const HWND_TOPMOST = -1
'Private Const HWND_NOTOPMOST = -2
'Private Const HWND_TOP = 0
'Private Const SWP_NOMOVE = &H2
'Private Const SWP_NOSIZE = &H1
'Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
'Private Const GWL_EXSTYLE = (-20)
'Private Const GWL_STYLE = (-16)
'Private Const WS_BORDER = &H800000
'Private Const WS_CAPTION = &HC00000
'Private Const WS_CHILD = &H40000000
'Private Const WS_CHILDWINDOW = (WS_CHILD)
'Private Const WS_CLIPCHILDREN = &H2000000
'Private Const WS_CLIPSIBLINGS = &H4000000
'Private Const WS_DISABLED = &H8000000
'Private Const WS_DLGFRAME = &H400000
'Private Const WS_EX_ACCEPTFILES = &H10&
'Private Const WS_EX_DLGMODALFRAME = &H1&
'Private Const WS_EX_NOPARENTNOTIFY = &H4&
'Private Const WS_EX_TOPMOST = &H8&
'Private Const WS_EX_TRANSPARENT = &H20&
'Private Const WS_GROUP = &H20000
'Private Const WS_HSCROLL = &H100000
'Private Const WS_MINIMIZEBOX = &H20000
'Private Const WS_MAXIMIZEBOX = &H10000
'Private Const WS_POPUP = &H80000000
'Private Const WS_SYSMENU = &H80000
'Private Const WS_TABSTOP = &H10000
'Private Const WS_THICKFRAME = &H40000
'Private Const WS_VSCROLL = &H200000
'Private Declare Function GetActiveWindow Lib "user32" () As Long
'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
'
'Private Type POINTAPI
'    X As Long
'    Y As Long
'End Type
'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'
'Private Const WM_GETTEXT = &HD
'
'Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'Private bOnTop As Boolean
'
'Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
'Private Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'Private Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
'Private Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
'Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
'Private Declare Function lSetFocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
'Private Declare Function GetDesktopWindow Lib "user32" () As Long
'Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
'
'Function WindowSPY()
''On Error Resume Next
'    'Call This In A Timer
'    Dim pt32 As POINTAPI, ptx As Long, pty As Long, sWindowText As String * 100
'    Dim sClassName As String * 100, hWndOver As Long, hwndParent As Long
'    Dim sParentClassName As String * 100, wID As Long, lWindowStyle As Long
'    Dim hInstance As Long, sParentWindowText As String * 100
'    Dim sModuleFileName As String * 100, r As Long
'    Static hWndLast As Long
'    Call GetCursorPos(pt32)
'    ptx = pt32.X
'    pty = pt32.Y
'    hWndOver = WindowFromPointXY(ptx, pty)
'
'    If hWndOver = 0 Then Exit Function
'
'    If hWndOver <> hWndLast Then
'        hWndLast = hWndOver
'        txtWinhWnd.Text = hWndOver
'        r = GetWindowText(hWndOver, sWindowText, 100)
'        txtWinTxt.Text = Left(sWindowText, r)
'        r = GetClassName(hWndOver, sClassName, 100)
'        txtWinClass.Text = Left(sClassName, r)
'        txtWinStyle.Text = GetWndStyle(hWndOver)
'        hwndParent = GetParent(hWndOver)
'
'
'        If hwndParent <> 0 Then
'            txtWinParenthWnd.Text = hwndParent
'            r = GetWindowText(hwndParent, sParentWindowText, 100)
'            txtWinParentTxt.Text = Left(sParentWindowText, r)
'
'            r = GetClassName(hwndParent, sParentClassName, 100)
'            txtWinParentClass.Text = Left(sParentClassName, r)
'            '// Program crashes if we try to get
'            '// the WindowWord of a MsgBox (ClassName = #32770)
'            If DangerousClass(txtWinParentClass.Text) = False Then wID = GetWindowWord(hWndOver, 0&)
'            txtWinIDNum.Text = wID
'        Else
'            txtWinIDNum.Text = "N/A"
'            txtWinParenthWnd.Text = "N/A"
'            txtWinParentTxt.Text = "N/A"
'            txtWinParentClass.Text = "N/A"
'        End If
'        '// Program crashes if we try to get
'        '// the WindowWord of a MsgBox (ClassName = #32770)
'    '    Debug.Print txtWinParentClass.Text
'        If DangerousClass(txtWinParentClass.Text) = False Then hInstance = GetWindowWord(hWndOver, 0&)
'        r = GetModuleFileName(hInstance, sModuleFileName, 100)
'        txtWinModule.Text = Left(sModuleFileName, r)
'    End If
'End Function
'Private Function DangerousClass(sClass As String) As Boolean
'   ' If sClass = "Shell_TrayWnd" Then Stop
'    If InStr(1, GetSetting(App.Title, "Settings", "DangerousClasses", "#32770,Shell_TrayWnd"), sClass, vbTextCompare) Then
'        DangerousClass = True
'    End If
'End Function
'
'
'
'Private Sub cmdExport_Click()
'    Dim sString As String
'    Dim sFile As String
'
'    sString = sString & lblLabel(0) & " - " & txtWinhWnd & vbCrLf
'    sString = sString & lblLabel(1) & " - " & txtWinClass & vbCrLf
'    sString = sString & lblLabel(2) & " - " & txtWinTxt & vbCrLf
'    sString = sString & lblLabel(3) & " - " & txtWinStyle & vbCrLf
'    sString = sString & lblLabel(4) & " - " & txtWinIDNum & vbCrLf
'    sString = sString & lblLabel(5) & " - " & txtWinParenthWnd & vbCrLf
'    sString = sString & lblLabel(6) & " - " & txtWinParentTxt & vbCrLf
'    sString = sString & lblLabel(7) & " - " & txtWinParentClass & vbCrLf
'    sString = sString & lblLabel(8) & " - " & txtWinModule & vbCrLf
'
'   ' sFile = InputBox("Please enter the file to save as", "Save Log...", App.Path & "\output1.txt")
'   ' If pshowcommondialog(scd_save, "Export...", SCD_OPENFLAGS, "Text Files|*.txt", , hwnd) = False Then Exit Sub
''    Dim iFileNum As Integer
''    iFileNum = FreeFile
''    Open pCmDlg.FileName For Output As iFileNum
''    Print #iFileNum, sString
''    Close #iFileNum
'    Dim frmForm As Form
''
'    Set frmForm = cFunc.frmMain.LoadNewDoc(vbwText)
'    frmForm.txtText.SelText = sString
'   ' loadnewdoc
'   ' ShellExecute 0, vbNullString, pCmDlg.FileName, vbNullString, vbNullString, vbNormalFocus
'End Sub
'
'Private Sub cmdOnTop_Click()
'    If bOnTop Then
'        'Hide
'       ' Show , cParent.frmMain
'        Set cmdOnTop.Picture = imgUp.Picture
'        MakeNormal hwnd
'        SetParent hwnd, 0
'    Else
'        'Hide
'        Set cmdOnTop.Picture = imgDown.Picture
'       ' Show ', Me
'        MakeTopMost hwnd
'        'SetParent hwnd, GetDesktopWindow
'    End If
'    bOnTop = Not bOnTop
'    If Visible Then txtWinhWnd.SetFocus
'End Sub
'
'Private Sub cmdPause_Click()
'    If tmrUpdate.Enabled = True Then
'        tmrUpdate.Enabled = False
'        lblMsg = "Press Alt+S to continue spying"
'        cmdPause.Caption = "&Spy"
'    Else
'        tmrUpdate.Enabled = True
'        lblMsg = "Press Alt+P to pause spying"
'        cmdPause.Caption = "&Pause"
'    End If
'End Sub
'
'
'Private Sub cmdBig_Click()
'    ShowControls True
'    Height = 3090
'    Width = 5295
'
'End Sub
'Private Sub cmdSmall_Click()
'    ShowControls False
'    Height = 1095
'    Width = 4560
''    Parent.MakeTopMost hwnd
'End Sub
'Private Sub ShowControls(bShow As Boolean)
'    Dim i As Byte
'    For i = 3 To 8
'        lblLabel(i).Visible = bShow
'    Next
'    txtWinModule.Visible = bShow
'    txtWinParentClass.Visible = bShow
'    txtWinParenthWnd.Visible = bShow
'    txtWinParentTxt.Visible = bShow
'    txtWinStyle.Visible = bShow
'    txtWinIDNum.Visible = bShow
'    cmdBig.Visible = Not (bShow)
'    On Error Resume Next
'    txtWinhWnd.SetFocus
'End Sub
'
'
'Private Sub Form_Load()
'   ' MakeTopMost hwnd
'   Set MousePointer = vbCustom
'   Set MouseIcon = picture1.Picture
'    bOnTop = Not CBool(GetSetting(App.Title, "Settings", "ClassSpyOnTop", True))
'    Call cmdOnTop_Click
'End Sub
'
'Private Sub Form_Resize()
'    Dim ctlControl As Control
'    For Each ctlControl In Controls
'        If TypeOf ctlControl Is TextBox Then
'            ctlControl.Width = ScaleWidth - ctlControl.Left
'        End If
'    Next
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    SaveSetting App.Title, "Settings", "ClassSpyOnTop", bOnTop
'
'End Sub
'
'Private Sub tmrUpdate_Timer()
'    WindowSPY
'End Sub
'Function GetWndStyle(hwnd) As String
'    Dim af As Long, s As String
'
'    ' Get normal style
'    af = GetWindowLong(hwnd, GWL_STYLE)
'
'    If af And WS_BORDER Then s = s & "Border "
'    If af And WS_CAPTION Then s = s & "Caption "
'    If af And WS_CHILD Then s = s & "Child "
'    If af And WS_CLIPCHILDREN Then s = s & "ClipChildren "
'    If af And WS_CLIPSIBLINGS Then s = s & "ClipSiblings "
'    If af And WS_DLGFRAME Then s = s & "DlgFrame "
'    If af And WS_GROUP Then s = s & "Group "
'    If af And WS_HSCROLL Then s = s & "HScroll "
'    If af And WS_MAXIMIZEBOX Then s = s & "MaximizeBox "
'    If af And WS_MINIMIZEBOX Then s = s & "MinimizeBox "
'    If af And WS_POPUP Then s = s & "Popup "
'    If af And WS_SYSMENU Then s = s & "SysMenu "
'    If af And WS_TABSTOP Then s = s & "TabStop "
'    If af And WS_THICKFRAME Then s = s & "ThickFrame "
'    If af And WS_VSCROLL Then s = s & "VScroll "
'
'    ' Get extended style
'    af = GetWindowLong(hwnd, GWL_EXSTYLE)
'    If af And WS_EX_DLGMODALFRAME Then s = s & "DlgModalFrame "
'    If af And WS_EX_NOPARENTNOTIFY Then s = s & "NoParentNotify "
'    If af And WS_EX_TOPMOST Then
'        s = s & "Topmost "
'        Debug.Print af
'    End If
'    If af And WS_EX_ACCEPTFILES Then s = s & "AcceptFiles "
'    If af And WS_EX_TRANSPARENT Then s = s & "Transparent "
'    GetWndStyle = s
'End Function
