VERSION 5.00
Begin VB.Form frmClassSpy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Class Spy"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
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
   Icon            =   "frmClassSpy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsmart 
      Caption         =   "smart"
      Height          =   330
      Left            =   1635
      TabIndex        =   24
      Top             =   1245
      Width           =   1365
   End
   Begin VB.Timer tmrCheck 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4155
      Top             =   2340
   End
   Begin VB.TextBox txtStyle 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1080
      TabIndex        =   21
      Top             =   2700
      Width           =   3870
   End
   Begin VB.TextBox txtClass 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1080
      TabIndex        =   20
      Top             =   2430
      Width           =   3870
   End
   Begin VB.CommandButton cmdParent 
      Caption         =   "Get Parent"
      Enabled         =   0   'False
      Height          =   345
      Left            =   3810
      TabIndex        =   19
      Top             =   1785
      Width           =   1215
   End
   Begin VB.TextBox txtCaption 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1080
      TabIndex        =   18
      Top             =   2145
      Width           =   3870
   End
   Begin VB.TextBox txthWnd 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1080
      TabIndex        =   17
      Top             =   1860
      Width           =   2295
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   345
      Left            =   3810
      TabIndex        =   15
      Top             =   3165
      Width           =   1215
   End
   Begin VB.PictureBox picStart 
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
      Height          =   510
      Left            =   2445
      Picture         =   "frmClassSpy.frx":000C
      ScaleHeight     =   510
      ScaleWidth      =   540
      TabIndex        =   10
      Top             =   1080
      Width           =   540
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
      Begin VB.Label lblClass 
         BackStyle       =   0  'Transparent
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Top             =   660
         Width           =   3870
      End
      Begin VB.Label lblCaption 
         BackStyle       =   0  'Transparent
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   375
         Width           =   3870
      End
      Begin VB.Label lblhWnd 
         BackStyle       =   0  'Transparent
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   105
         Width           =   3870
      End
      Begin VB.Label lblStyle 
         BackStyle       =   0  'Transparent
         Height          =   420
         Left            =   1080
         TabIndex        =   5
         Top             =   945
         Width           =   3870
      End
      Begin VB.Shape shpBorder 
         BorderColor     =   &H00C0C0C0&
         Height          =   1410
         Left            =   0
         Top             =   0
         Width           =   5145
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
         TabIndex        =   4
         Top             =   105
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
         TabIndex        =   3
         Top             =   660
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
         TabIndex        =   2
         Top             =   375
         Width           =   825
      End
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
         TabIndex        =   1
         Top             =   945
         Width           =   825
      End
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Retreives the hWnd, Class and Style of a window"
      Height          =   255
      Left            =   300
      TabIndex        =   23
      Top             =   345
      Width           =   3825
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Class Spy"
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
      TabIndex        =   22
      Top             =   75
      Width           =   990
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   4470
      Picture         =   "frmClassSpy.frx":0316
      Top             =   165
      Width           =   480
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
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   4
      X1              =   120
      X2              =   5055
      Y1              =   3030
      Y2              =   3030
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      Caption         =   "&Help"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   105
      MouseIcon       =   "frmClassSpy.frx":0BE0
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   3195
      Width           =   360
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H80000018&
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
      Index           =   16
      Left            =   135
      TabIndex        =   14
      Top             =   2715
      Width           =   825
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H80000018&
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
      Index           =   15
      Left            =   135
      TabIndex        =   13
      Top             =   2145
      Width           =   825
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H80000018&
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
      Index           =   14
      Left            =   135
      TabIndex        =   12
      Top             =   2430
      Width           =   825
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H80000018&
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
      Index           =   13
      Left            =   135
      TabIndex        =   11
      Top             =   1875
      Width           =   825
   End
   Begin VB.Label lblLabel 
      Caption         =   "Drag the icon to start spying:"
      Height          =   255
      Index           =   0
      Left            =   135
      TabIndex        =   9
      Top             =   1215
      Width           =   2175
   End
   Begin VB.Image imgCur 
      Height          =   300
      Left            =   4815
      MouseIcon       =   "frmClassSpy.frx":0EEA
      Top             =   3825
      Width           =   375
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   5
      X1              =   105
      X2              =   5010
      Y1              =   3045
      Y2              =   3045
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
End
Attribute VB_Name = "frmClassSpy"
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
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Public cParent As clsClassSpy
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOP = 0
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

'Private Type POINTAPI
'        x As Long
'        y As Long
'End Type
Private Const GWL_EXSTYLE = (-20)
Private Const GWL_STYLE = (-16)

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Const WS_BORDER = &H800000
Private Const WS_CAPTION = &HC00000
Private Const WS_CHILD = &H40000000
Private Const WS_CHILDWINDOW = (WS_CHILD)
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_DISABLED = &H8000000
Private Const WS_DLGFRAME = &H400000
Private Const WS_EX_ACCEPTFILES = &H10&
Private Const WS_EX_DLGMODALFRAME = &H1&
Private Const WS_EX_NOPARENTNOTIFY = &H4&
Private Const WS_EX_TOPMOST = &H8&
Private Const WS_EX_TRANSPARENT = &H20&
Private Const WS_GROUP = &H20000
Private Const WS_HSCROLL = &H100000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_POPUP = &H80000000
Private Const WS_SYSMENU = &H80000
Private Const WS_TABSTOP = &H10000
Private Const WS_THICKFRAME = &H40000
Private Const WS_VSCROLL = &H200000
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_SHOW = 5
Private Const SW_HIDE = 0

Private Const WM_GETTEXT = &HD

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private bOnTop As Boolean

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long


Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long


Private Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long


Private Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)


Private Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer



Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function lSetFocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'Private Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
'Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
'Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
'Private Const WM_MOUSEMOVE = &H200
'Private Const MK_LBUTTON = &H1

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private lOldLeft As Long
Private bDragging As Boolean
Private hWndLast As Long
Private CurrenthWnd As Long
'Private tOldRect As RECT
'Implements ISubclass
Private Sub GetWindowInfo(hWndOver As Long)
    Dim tP As POINTAPI
    Dim sWindowText As String * 100
    Dim sClassName As String * 100
    Dim r As Long

    CurrenthWnd = hWndOver
    If hWndOver = 0 Then Exit Sub
    If hWndOver <> hWndLast Then
        hWndLast = hWndOver
        lblhWnd.Caption = hWndOver & " (&&H" & Hex(hWndOver) & ")"
        r = GetWindowText(hWndOver, sWindowText, 100)
        lblCaption.Caption = Left(sWindowText, r)
        r = GetClassName(hWndOver, sClassName, 100)
        lblClass.Caption = Left(sClassName, r)
        lblStyle.Caption = GetWndStyle(hWndOver)
    End If
End Sub
Private Sub GetWindowInfoEx()
    Dim tP As POINTAPI
    Dim sWindowText As String * 100
    Dim sClassName As String * 100
    Dim hWndOver As Long
    Dim r As Long
    
    If CurrenthWnd = 0 Then Exit Sub
    txthWnd.Text = CurrenthWnd & " (&&H" & Hex(CurrenthWnd) & ")"
    r = GetWindowText(CurrenthWnd, sWindowText, 100)
    txtCaption.Text = Left(sWindowText, r)
    r = GetClassName(CurrenthWnd, sClassName, 100)
    txtClass.Text = Left(sClassName, r)
    txtStyle.Text = GetWndStyle(CurrenthWnd)

    cmdParent.Enabled = True
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdParent_Click()
    Dim lhWnd As Long
    lhWnd = GetParent(CurrenthWnd)
    If lhWnd <> 0 Then
        CurrenthWnd = lhWnd
        GetWindowInfoEx
    Else
        cFunc.Dialogs.ShowWarning "This window does not have a parent"
    End If
End Sub

Private Sub cmdsmart_Click()
    frmSmartUpdate.Show vbModal
End Sub

Private Sub Form_Deactivate()
   ' Debug.Print "deactivate"
End Sub

'Private Sub Command1_Click()
'Dim hdc As Long
'Dim tWindowRect As RECT
'    hdc = GetDC(0&)
' GetWindowRect CurrenthWnd, tWindowRect
' 'erase previous
' DrawFocusRect hdc, tOldRect
'
'  DrawFocusRect hdc, tWindowRect
' ReleaseDC 0&, hdc
' LSet tOldRect = tWindowRect
'
'End Sub

Private Sub Form_Load()
   ' AttachMessage Me, picStart.hwnd, WM_MOUSEMOVE
End Sub

Private Sub Form_LostFocus()
 'Debug.Print "deactivate"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  '  DetachMessage Me, picStart.hwnd, WM_MOUSEMOVE
    SaveSetting App.Title, "Settings", "ClassSpyOnTop", bOnTop
End Sub

Private Sub lblHelp_Click()
    ShowHTMLHelpTopic 2, hwnd
End Sub

Private Sub picStart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton And bDragging = False Then
        bDragging = True
        SetCapture picStart.hwnd
        Set picStart.MouseIcon = imgCur.MouseIcon
        picStart.MousePointer = vbCustom
        ShowWindow cFunc.ApphWnd, SW_HIDE
        lOldLeft = Left
        Left = -10000
       ' Height = 0
        SetParent picSpy.hwnd, 0&
        MakeTopMost picSpy.hwnd
        picSpy.Visible = True
        tmrCheck.Enabled = True
      '  picStart_MouseMove vbLeftButton, -1, -1, -1
    End If
End Sub

Private Sub picStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bDragging And Button = vbLeftButton Then
        Dim tP As POINTAPI
        Dim tR As RECT
        Dim tScreen As RECT
        Dim hWndOver As Long
      '  Dim tScreen As RECT
        GetCursorPos tP
        hWndOver = WindowFromPointXY(tP.X, tP.Y)
        
        GetClientRect GetDesktopWindow(), tScreen
        GetWindowRect picSpy.hwnd, tR
     '   GetWindowRect 0&, tScreen
        tP.X = tP.X + 20
        tP.Y = tP.Y - 10 '+ 40
        If tP.X + (tR.Right - tR.Left) > (tScreen.Right - tScreen.Left) Then
            tP.X = tP.X - (tR.Right - tR.Left) - 40
        End If
        If tP.Y + (tR.Bottom - tR.Top) > (tScreen.Bottom - tScreen.Top) Then
            tP.Y = (Screen.Height / Screen.TwipsPerPixelX) - (tR.Bottom - tR.Top)
        ElseIf tP.Y < 0 Then
            tP.Y = 0
        End If
'        SetFocus
        MoveWindow picSpy.hwnd, tP.X, tP.Y, tR.Right - tR.Left, tR.Bottom - tR.Top, True
        


        GetWindowInfo hWndOver
    End If
End Sub

Private Sub picStart_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim hDC As Long
   ' Debug.Print "mouseup"
    If bDragging Then
        ReleaseCapture
        picStart.MousePointer = vbDefault
        'frmSpy.Hide
        SetParent picSpy.hwnd, hwnd
        MakeNormal picSpy.hwnd
        'MakeNormal picSpy.hwnd
        picSpy.Visible = False
        'cFunc.frmMain.WindowState = vbNormal
       ' cFunc.frmmain.Visible = True
        ShowWindow cFunc.ApphWnd, SW_SHOW
      '  Height = 3240
      Left = lOldLeft
      tmrCheck.Enabled = False
       ' SetFocus
        bDragging = False
        GetWindowInfoEx
'        hdc = GetDC(0&)
'        'erase previous
'        DrawFocusRect hdc, tOldRect
'        SetRectEmpty tOldRect
'        ReleaseDC 0&, hdc
    End If
End Sub
'Private Sub tmrUpdate_Timer()
'    WindowSPY
'End Sub
Function GetWndStyle(hwnd) As String
    Dim af As Long, s As String

    ' Get normal style
    af = GetWindowLong(hwnd, GWL_STYLE)
    
    If af And WS_BORDER Then s = s & "Border "
    If af And WS_CAPTION Then s = s & "Caption "
    If af And WS_CHILD Then s = s & "Child "
    If af And WS_CLIPCHILDREN Then s = s & "ClipChildren "
    If af And WS_CLIPSIBLINGS Then s = s & "ClipSiblings "
    If af And WS_DLGFRAME Then s = s & "DlgFrame "
    If af And WS_GROUP Then s = s & "Group "
    If af And WS_HSCROLL Then s = s & "HScroll "
    If af And WS_MAXIMIZEBOX Then s = s & "MaximizeBox "
    If af And WS_MINIMIZEBOX Then s = s & "MinimizeBox "
    If af And WS_POPUP Then s = s & "Popup "
    If af And WS_SYSMENU Then s = s & "SysMenu "
    If af And WS_TABSTOP Then s = s & "TabStop "
    If af And WS_THICKFRAME Then s = s & "ThickFrame "
    If af And WS_VSCROLL Then s = s & "VScroll "

    ' Get extended style
    af = GetWindowLong(hwnd, GWL_EXSTYLE)
    If af And WS_EX_DLGMODALFRAME Then s = s & "DlgModalFrame "
    If af And WS_EX_NOPARENTNOTIFY Then s = s & "NoParentNotify "
    If af And WS_EX_TOPMOST Then
        s = s & "Topmost "
       ' Debug.Print af
    End If
    If af And WS_EX_ACCEPTFILES Then s = s & "AcceptFiles "
    If af And WS_EX_TRANSPARENT Then s = s & "Transparent "
    
    GetWndStyle = s

End Function
Public Sub MakeNormal(hwnd As Long)
    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub
Public Sub MakeTopMost(hwnd As Long)
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Private Sub tmrCheck_Timer()
    If GetActiveWindow <> hwnd Then picStart_MouseUp 0, -1, -1, -1
End Sub
