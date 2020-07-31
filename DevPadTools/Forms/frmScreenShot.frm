VERSION 5.00
Object = "{1F5999A2-1D0B-11D4-82CF-004005AAE138}#5.2#0"; "VBWTBAR_6.OCX"
Begin VB.Form frmScreenShot 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Screen Shot"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   7830
   Icon            =   "frmScreenShot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   244
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   522
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picHolder 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   0
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   522
      TabIndex        =   4
      Top             =   0
      Width           =   7830
      Begin vbwTBar.cToolbar tbrMain 
         Left            =   -15
         Top             =   0
         _ExtentX        =   2011
         _ExtentY        =   582
      End
      Begin vbwTBar.cToolbarHost tbhMain 
         Height          =   315
         Left            =   3105
         TabIndex        =   5
         Top             =   75
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   556
      End
   End
   Begin VB.PictureBox picCorner 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6600
      Picture         =   "frmScreenShot.frx":000C
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   2700
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   100
      Left            =   0
      SmallChange     =   10
      TabIndex        =   2
      Top             =   2760
      Width           =   2415
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2535
      LargeChange     =   100
      Left            =   7320
      SmallChange     =   10
      TabIndex        =   1
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   15
      ScaleHeight     =   270
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   900
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.PictureBox vbalImgLst 
      Height          =   480
      Left            =   5190
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   6
      Top             =   900
      Width           =   1200
   End
End
Attribute VB_Name = "frmScreenShot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public cParent As clsScreenShot
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Const GWL_EXSTYLE = (-20)
Private Const GWL_STYLE = (-16)

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function lSetFocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetClassWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long


Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

'
Private fCapture As Boolean
Private Sub PointWindow()
    If Caption = "Screen Shot" Then
        fCapture = True
        Caption = "Screen Shot - Pointing... Right click to take a screen shot"
        Call SetCapture(hwnd)
    Else
        fCapture = False
        Caption = "Screen Shot"
        ReleaseCapture
    End If
End Sub

Private Sub Command1_Click()
    GetShot GetDesktopWindow
End Sub

Private Sub Form_Load()
    With tbrMain
        .ImageSource = CTBExternalImageList
        .SetImageList vbalImgLst, CTBImageListNormal

       .CreateToolbar 16, , True, True

       .Wrappable = True
       
       .AddButton "Save", 0, , , Empty, CTBNormal, "Save"
       .AddButton Empty, -1, , , , CTBSeparator
       .AddButton "Take a full screen shot", 1, , , Empty, , "ScreenShot"
       .AddButton "Point to a window using the mouse", 2, , , Empty, , "Point"
       .AddButton "Take a screen shot after the specified interval", 3, , , Empty, , "Timer"
       .AddButton Empty, -1, , , , CTBSeparator
       .AddButton "Screen Shot Preferences...", 4, , , Empty, , "Options"
       .ButtonEnabled("Timer") = False
       .ButtonEnabled("Options") = False
    End With
    '// Capture Main toolbar
    With tbhMain
         .BorderStyle = etbhBorderStyleNone
         .Width = ScaleWidth
         .Height = tbrMain.ToolbarHeight * Screen.TwipsPerPixelY
         .Capture tbrMain
    End With
    Form_Resize
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim pt As POINTAPI
    Dim lWindowhWnd As Long
    Dim idProc As Long
    If fCapture Then
        If Button = vbRightButton Then
            ' Set point and convert it to screen coordinates
            pt.X = X ' / Screen.TwipsPerPixelX
            pt.Y = Y ' / Screen.TwipsPerPixelY
            ClientToScreen hwnd, pt
            ' Find window under it
            lWindowhWnd = WindowFromPoint(pt.X, pt.Y)
'            Dim stext As String
'                stext = Space(255)
'                GetWindowText lwindowhwnd, stext, 255
'                chkParentCapture.Caption = stext
            'Call SetCapture(hwnd)
            If lWindowhWnd <> vbNull Then
                If lWindowhWnd <> Me.hwnd Then
                    Hide
                End If
                
                SetForegroundWindow lWindowhWnd
                lSetFocus lWindowhWnd
               ' If Abs(chkParentCapture.Value) Then
               '     lWindowhWnd = GetActiveWindow
               ' End If
                GetShot lWindowhWnd
                Show
                SetForegroundWindow hwnd
            End If
'            Button = -1
        End If
        Call PointWindow
    End If
End Sub


Public Sub GetShot(lWindowhWnd As Long)
    Dim lWindowhDC As Long
    Dim lParenthWnd As Long
    Dim lChildhWnd As Long

    Dim nLeft As Long
    Dim nTop As Long
    Dim nWidth As Long
    Dim nHeight As Long
    Dim rRect As RECT
    Screen.MousePointer = 11
    picScreen.Visible = True
    picScreen.Cls
    Set picScreen.Picture = Nothing
    DoEvents
    
    GetWindowRect lWindowhWnd, rRect
    lWindowhDC = GetWindowDC(lWindowhWnd)
    '// Get coordinates
    nLeft = 0
    nTop = 0
    nWidth = rRect.Right - rRect.Left
    nHeight = rRect.Bottom - rRect.Top
    picScreen.Width = nWidth
    picScreen.Height = nHeight
    '// Blt to frm.picScreen
    BitBlt picScreen.hDC, 0, 0, nWidth, nHeight, lWindowhDC, nLeft, nTop, vbSrcCopy
    '// Del DC
    ReleaseDC lWindowhWnd, lWindowhDC
    '// set picture
    picScreen.Picture = picScreen.Image
    
    VScroll1.Value = 0
    HScroll1.Value = 0
    Call Form_Resize
    Screen.MousePointer = 0
End Sub

Private Sub SaveImage()
   ' If cFunc.ShowCommonDialog(scd_save, "Save As", SCD_saveFLAGS, "Bitmaps|*.bmp", , hwnd) = False Then Exit Sub
    On Error Resume Next
    SavePicture picScreen.Picture, cFunc.cmdlg.FileName
    If Err Then MsgBox "Error " & Err & " : " & Error
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim pt As POINTAPI
'Dim lWindowhWnd As Long
'    ' Set point and convert it to screen coordinates
'            pt.X = X / Screen.TwipsPerPixelX
'            pt.Y = Y / Screen.TwipsPerPixelY
'            ClientToScreen hwnd, pt
'            ' Find window under it
'            lWindowhWnd = WindowFromPoint(pt.X, pt.Y)
'            'Dim stext As String
'             '   stext = Space(255)
'              '  GetWindowText lwindowhwnd, stext, 255
'                chkParentCapture.Caption = lWindowhWnd 'stext
End Sub

Private Sub Form_Resize()
On Error Resume Next
    VScroll1.Move ScaleWidth - VScroll1.Width, tbhMain.Height, 17, ScaleHeight - VScroll1.Top - 17
    VScroll1.Max = IIf(Sgn(picScreen.Height - VScroll1.Height) = -1, 0, picScreen.Height - VScroll1.Height)
    HScroll1.Move 0, ScaleHeight - HScroll1.Height, ScaleWidth - 17, 17
    HScroll1.Max = IIf(Sgn(picScreen.Width - HScroll1.Width) = -1, 0, picScreen.Width - HScroll1.Width)
    picCorner.Left = ScaleWidth - picCorner.Width
    picCorner.Top = ScaleHeight - picCorner.Width
    
    
    
    tbhMain.Move 0, 0, ScaleWidth, tbrMain.ToolbarHeight ' * Screen.TwipsPerPixelX
    
    picHolder.Height = tbrMain.ToolbarHeight
    picScreen.Left = 0
    picScreen.Top = tbhMain.Height
End Sub
Private Sub HScroll1_Change()
    picScreen.Left = (HScroll1.Value * -1)
End Sub

Private Sub HScroll1_GotFocus()
    If picScreen.Visible Then picScreen.SetFocus
End Sub



Private Sub tbrMain_ButtonClick(ByVal lButton As Long)
    Select Case lButton
    Case 0 ' "save"
        SaveImage
    Case 2 '"screenshot"
        GetShot GetDesktopWindow
    Case 3 '"point"
        PointWindow
    End Select
End Sub

Private Sub VScroll1_Change()
    picScreen.Top = 25 + (VScroll1.Value * -1)
End Sub

Private Sub VScroll1_GotFocus()
On Error Resume Next
    picScreen.SetFocus
End Sub
