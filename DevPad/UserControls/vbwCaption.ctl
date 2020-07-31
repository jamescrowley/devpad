VERSION 5.00
Begin VB.UserControl vbwCaption 
   Alignable       =   -1  'True
   BackColor       =   &H80000002&
   ClientHeight    =   2790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6300
   ControlContainer=   -1  'True
   ScaleHeight     =   2790
   ScaleWidth      =   6300
   Begin VB.Timer tmrUpdate 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "vbwCaption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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
' Thanks to VB Clarity for splitter graphics
' http://www.vbclarity.com/
' Thanks to VB Accelerator for title bar graphics
' http://www.vbaccelerator.com/
'
' WARNING: This code is very messy, as it is currently
' being exported to a seperate ActiveX Control project

Option Explicit

'function prototypes
'win32 forward declarations
'constants
'COMBINEREGION
Private Const RGN_AND = 1
Private Const RGN_OR = 2
Private Const RGN_XOR = 3
Private Const RGN_DIFF = 4
Private Const RGN_COPY = 5
Private Const WM_NCACTIVATE = &H86
Private Const WM_MOVE = &H3
Private Const WM_CLOSE = &H10
Private Const HTCAPTION = 2
Private Const DFC_CAPTION = 1
Private Const DFCS_CAPTIONCLOSE = &H0
Private Const DFCS_FLAT = &H4000
Private Const DFCS_INACTIVE = &H100
Private Const DFCS_PUSHED = &H200
Private Const SPI_GETNONCLIENTMETRICS = 41
' Sizer constants
Private Const HTRIGHT = 11
Private Const HTLEFT = 10
Private Const HTTOP = 12
Private Const HTBOTTOM = 15
Private Const HTBOTTOMRIGHT = 17

Private Enum DCFlags
   DC_ACTIVE = &H1&
   DC_SMALLCAP = &H2&
   DC_ICON = &H4&
   DC_TEXT = &H8&
   DC_INBUTTON = &H10&
   DC_GRADIENT = &H20&
End Enum
' GetSystemMetrics stuff
Private Const LF_FACESIZE = 32

' For some bizarre reason, maybe to do with byte
' alignment, the LOGFONT sucture we must apply
' to NONCLIENTMETRICS seems to require an LF_FACESIZE
' 4 bytes smaller than normal:
Private Type NMLOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE - 4) As Byte
End Type
Private Type NONCLIENTMETRICS
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As NMLOGFONT
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As NMLOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As NMLOGFONT
    lfStatusFont As NMLOGFONT
    lfMessageFont As NMLOGFONT
End Type

'
'Private Enum ESetWindowPosStyles
'    SWP_SHOWWINDOW = &H40
'    SWP_HIDEWINDOW = &H80
'    SWP_FRAMECHANGED = &H20 ' The frame changed: send WM_NCCALCSIZE
'    SWP_NOACTIVATE = &H10
'    SWP_NOCOPYBITS = &H100
'    SWP_NOMOVE = &H2
'    SWP_NOOWNERZORDER = &H200 ' Don't do owner Z ordering
'    SWP_NOREDRAW = &H8
'    SWP_NOREPOSITION = SWP_NOOWNERZORDER
'    SWP_NOSIZE = &H1
'    SWP_NOZORDER = &H4
'    SWP_DRAWFRAME = SWP_FRAMECHANGED
'    HWND_NOTOPMOST = -2
'End Enum

Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_OVERLAPPED = &H0&
Private Const WS_POPUP = &H80000000
Private Const WS_CHILD = &H40000000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_VISIBLE = &H10000000
Private Const WS_DISABLED = &H8000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_CAPTION = &HC00000
Private Const WS_BORDER = &H800000
Private Const WS_DLGFRAME = &H400000
Private Const WS_VSCROLL = &H200000
Private Const WS_HSCROLL = &H100000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const WS_GROUP = &H20000
Private Const WS_TABSTOP = &H10000
Private Const WS_SIZEBOX = WS_THICKFRAME
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

'function prototypes
'GDI32
Private Declare Function apiCombineRgn Lib "gdi32" Alias "CombineRgn" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function apiCreateBitmap Lib "gdi32" Alias "CreateBitmap" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function apiCreatePatternBrush Lib "gdi32" Alias "CreatePatternBrush" (ByVal hBitmap As Long) As Long
Private Declare Function apiCreateRectRgn Lib "gdi32" Alias "CreateRectRgn" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function apiCreateRectRgnIndirect Lib "gdi32" Alias "CreateRectRgnIndirect" (lpRect As RECT) As Long
Private Declare Function apiDeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As Long) As Long
Private Declare Function apiGetClipBox Lib "gdi32" Alias "GetClipBox" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function apiPatBlt Lib "gdi32" Alias "PatBlt" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function apiSelectClipRgn Lib "gdi32" Alias "SelectClipRgn" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function apiSelectObject Lib "gdi32" Alias "SelectObject" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function apiSetRectRgn Lib "gdi32" Alias "SetRectRgn" (ByVal hRgn As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'USER32
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function apiCopyRect Lib "user32" Alias "CopyRect" (lpDstRect As RECT, lpSrcRect As RECT) As Long
Private Declare Function apiFillRect Lib "user32" Alias "FillRect" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function apiInflateRect Lib "user32" Alias "InflateRect" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function apiIntersectRect Lib "user32" Alias "IntersectRect" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Private Declare Function apiIsRectEmpty Lib "user32" Alias "IsRectEmpty" (lpRect As RECT) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function SetWindowPosAPI Lib "user32" Alias "SetWindowPos" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function DrawCaption Lib "user32" (ByVal hWnd As Long, ByVal lhDC As Long, tR As RECT, ByVal eFlag As DCFlags) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long

Private m_hDC           As Long
Private m_hBmp          As Long
Private m_hBmpOld       As Long
Private m_lWidthDC      As Long
Private m_lHeightDC     As Long
Private m_tButtonR      As RECT
Private bSplitting      As Boolean
'private members
Private m_lngPower2(0 To 31) As Long

' client metrics (for caption height)
Private m_tNCM As NONCLIENTMETRICS
Private m_bButtonDown   As Boolean
Private m_bButtonOver   As Boolean
Private m_bMouseDown    As Boolean
Private m_bRunTime      As Boolean
Private m_sCaption      As String
Private m_hWnd          As Long
Private m_bDragging     As Boolean
Private m_lOldSize      As Long

Private tStartPoint         As POINTAPI
Private tPoint              As POINTAPI

Private tParentRect         As RECT
Private tClientRect         As RECT
Private tbox                As RECT
Private tOldRect            As RECT

Private TwipsPerPixelX      As Long
Private TwipsPerPixelY      As Long

Private bDocking            As Boolean
Private Docking             As AlignConstants
Private lOldStyle           As Long
'agghh!
Public WithEvents picHolder As PictureBox
Attribute picHolder.VB_VarHelpID = -1
Public WithEvents frmDockWindow As Form
Attribute frmDockWindow.VB_VarHelpID = -1
Public frmParent            As Form

Private m_bFloating         As Boolean
Private m_bDockWindowLoaded As Boolean
Private m_bIgnore           As Boolean
Private m_Docking           As AlignConstants
Private m_bUnloading        As Boolean
Private m_bStartUp          As Boolean
Private m_sAppName          As String
Private m_sSection          As String
Private m_sKey              As String
Private m_DefaultPos        As AlignConstants
Private m_DefaultWidth      As Single
Private m_DefaultHeight     As Single
Private DockStatus          As AlignConstants
Private vLastPointer        As MousePointerConstants

Public Event BeforeDock(ByRef bCancel As Boolean)
Public Event Dock(ByVal Pos As AlignConstants)
Public Event UnDock()
Public Event BeforeUnDock(ByRef bCancel As Boolean)

'*** VB Accelerator ***
' thanks to vbaccelerator.com for this!
Private Sub DrawTitleBar()
    Dim tR As RECT
    Dim lwidth As Long
    Dim lHeight As Long
    Dim lhDC As Long
    Dim lBarColor As Long
    Dim lCapColor As Long
    Dim lStyle As Long
    GetClientRect UserControl.hWnd, tR
    lwidth = Abs(tR.Right - tR.Left)
    lHeight = Abs(tR.Bottom - tR.Top)
    ' Memory DC for draw speed:
    If lwidth > m_lWidthDC Or lHeight > m_lHeightDC Then
        ' get a new dc
        pRebuildDC lwidth, lHeight
    End If
    lhDC = UserControl.hdc
    ' Draw gradient if possible
    lStyle = DC_SMALLCAP Or DC_TEXT Or DC_GRADIENT
    
    lStyle = lStyle Or DC_ACTIVE
    ' set colours
    lBarColor = (vbActiveTitleBar And &H1F&)
    lCapColor = (vbTitleBarText And &H1F&)
    ' draw the title bar caption
    DrawCaption m_hWnd, m_hDC, tR, lStyle
    ' draw button
    pDrawCloseButton m_hDC
    ' draw all at once
    BitBlt UserControl.hdc, 0, 0, lwidth, lHeight, m_hDC, 0, 0, vbSrcCopy
End Sub
'we only build the dc when we have to...
Private Sub pRebuildDC(ByVal lwidth As Long, ByVal lHeight As Long)
   If lwidth > m_lWidthDC Then
      m_lWidthDC = lwidth
   End If
   If lHeight > m_lHeightDC Then
      m_lHeightDC = lHeight
   End If
   If Not m_hBmpOld = 0 Then
      SelectObject m_hDC, m_hBmpOld
      m_hBmpOld = 0
   End If
   If Not m_hBmp = 0 Then
      DeleteObject m_hBmp
      m_hBmp = 0
   End If
   If m_hDC = 0 Then
      m_hDC = CreateCompatibleDC(UserControl.hdc)
   End If
   If Not m_hDC = 0 Then
      m_hBmp = CreateCompatibleBitmap(UserControl.hdc, m_lWidthDC, m_lHeightDC)
      If Not m_hBmp = 0 Then
         m_hBmpOld = SelectObject(m_hDC, m_hBmp)
      End If
   End If
   If m_hBmpOld = 0 Then
      If Not m_hBmp = 0 Then
         DeleteObject m_hBmp
         m_hBmp = 0
      End If
      If Not m_hDC = 0 Then
         DeleteDC m_hDC
         m_hDC = 0
      End If
      m_lWidthDC = 0
      m_lHeightDC = 0
   End If
End Sub
' thanks to vbaccelerator.com for this!
Private Sub pDrawCloseButton(ByVal lhDC As Long)
Dim lH As Long
Dim lStyle As Long
Dim lType As Long
Dim tR As RECT

    GetMetrics
    lH = m_tNCM.iCaptionHeight - 2
    GetClientRect UserControl.hWnd, tR
    UserControl.Height = m_tNCM.iCaptionHeight * TwipsPerPixelY
    tR.Left = tR.Right - lH - 2
    tR.Top = tR.Top + 2
    tR.Right = tR.Left + lH
    tR.Bottom = tR.Top + lH - 2
    LSet m_tButtonR = tR
    lType = DFC_CAPTION
    lStyle = DFCS_CAPTIONCLOSE
    If (m_bButtonDown And m_bButtonOver) Then
       lStyle = lStyle Or DFCS_PUSHED
    End If
    DrawFrameControl lhDC, tR, lType, lStyle
End Sub
Private Function GetMetrics() As Boolean
Dim lR As Long
    ' Get Non-client metrics:
    m_tNCM.cbSize = 340 ' LenB(m_tNCM) - why doesn't this go?
    lR = SystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, m_tNCM, 0)
End Function
'*** End VB Accelerator ***

Private Sub picHolder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And bSplitting Then
        'abort splitting
        bSplitting = False
        DrawSplitter
    End If
End Sub

Private Sub picHolder_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'update splitter bar
    If bSplitting Then
        DrawSplitter
    Else
        If MouseInSizer And picHolder.Visible = True Then
            If vLastPointer = -1 Or (Screen.MousePointer <> vbSizeWE And Screen.MousePointer <> vbSizeNS) Then 'save the pointer
                vLastPointer = Screen.MousePointer
            End If
            If picHolder.Align = vbAlignLeft Or picHolder.Align = vbAlignRight Then
                Screen.MousePointer = vbSizeWE
            Else
                Screen.MousePointer = vbSizeNS
            End If
        Else
            If vLastPointer <> -1 Then Screen.MousePointer = vLastPointer
            vLastPointer = -1
        End If
    End If
End Sub
Private Sub picHolder_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lVal As Long

    If bSplitting Then
        bSplitting = False
        'draw the splitter...
        'returns new width/height
        lVal = DrawSplitter
        Select Case picHolder.Align
        Case vbAlignLeft
            '
            picHolder.Width = lVal * TwipsPerPixelX + 75
        Case vbAlignRight
            picHolder.Width = picHolder.Width + ((lVal * -1) * TwipsPerPixelX)
        Case vbAlignTop
            picHolder.Height = lVal * TwipsPerPixelX + 75
        Case vbAlignBottom
            picHolder.Height = picHolder.Height + ((lVal * -1) * TwipsPerPixelX)
        End Select
    End If
End Sub

Private Sub UserControl_ExitFocus()
    'If m_bDragging = True Then AbortDrag
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And m_bDragging = True Then
        AbortDrag
    End If
End Sub
Private Sub AbortDrag()
    'abort docking...
    Docking = DockStatus
    m_bMouseDown = False
    'end the drag
    EndFRDrag -1, -1
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim tp As POINTAPI
    If (Button And vbLeftButton) = vbLeftButton Then
        m_bMouseDown = True
        GetCursorPos tp
        ' Move the form if we are floating:
        ScreenToClient m_hWnd, tp
        If PtInRect(m_tButtonR, tp.x, tp.y) <> 0 Then
           ' Over close button:
           m_bButtonDown = True
           m_bButtonOver = True
           pDrawCloseButton UserControl.hdc
        Else
           'over the caption...
           'begin dragging
           BeginFRDrag x, y
        End If
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim tp As POINTAPI
   If m_bMouseDown Then
        'get the current cursor pos
        GetCursorPos tp
        If m_bButtonDown Then
            'adjust values
            ScreenToClient m_hWnd, tp
            If PtInRect(m_tButtonR, tp.x, tp.y) <> 0 Then
                'mouse is over close button
                If Not m_bButtonOver Then
                    m_bButtonOver = True
                    pDrawCloseButton UserControl.hdc
                End If
            Else
                If m_bButtonOver Then
                    'mouse is no longer over button
                    m_bButtonOver = False
                    pDrawCloseButton UserControl.hdc
                End If
            End If
        ElseIf m_bDragging Then
            'update dragging rect
            DoFRDrag x, y
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim tp As POINTAPI
    If m_bMouseDown And m_bButtonDown Then
        'mouse is up...
        'update graphics
        m_bButtonDown = False
        m_bButtonOver = False
        'redraw close button
        pDrawCloseButton UserControl.hdc
        'get the cursor pos
        GetCursorPos tp
        ScreenToClient m_hWnd, tp
        If PtInRect(m_tButtonR, tp.x, tp.y) <> 0 Then
           ' we clicked the close button
           PostMessage m_hWnd, WM_CLOSE, 0, 0
        End If
    ElseIf m_bMouseDown And Button = vbLeftButton Then
        'end window dragging
        EndFRDrag x, y
    End If
    m_bMouseDown = False
End Sub
Private Sub UserControl_Paint()
    'draw the title bar
    DrawTitleBar
End Sub
Public Property Get Caption() As String
   Caption = m_sCaption
End Property
Public Property Let Caption(ByVal sCaption As String)
    m_sCaption = sCaption
    If m_bRunTime Then
        'update caption this way...
        'using VB methods re-shows the
        'standard title bar rather than our custom
        'draw one
        SetWindowText m_hWnd, sCaption
    End If
    PropertyChanged "Caption"
    DrawTitleBar
End Property
Private Sub UserControl_Resize()
    DrawTitleBar
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Caption = PropBag.ReadProperty("Caption", "")
    'init control
    pInitialise
End Sub
Private Sub pInitialise()
    m_bRunTime = UserControl.Ambient.UserMode
    m_hWnd = UserControl.Parent.hWnd
    tmrUpdate.Enabled = (m_bRunTime)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'save caption
    PropBag.WriteProperty "Caption", Caption, ""
End Sub

Private Sub BeginFRDrag(x As Single, y As Single)
Dim tp As POINTAPI
Dim m_Clienthwnd As Long
    'convert points to POINTAPI struct
    tPoint.x = x
    tPoint.y = y
    'get screen area of toolbar
    GetWindowRect Parent.hWnd, tParentRect
    m_Clienthwnd = FindWindowEx(frmParent.hWnd, 0, "MDICLIENT", vbNullString)
    GetWindowRect m_Clienthwnd, tClientRect
    If DockStatus = vbAlignLeft Then
        tClientRect.Left = tClientRect.Left - (picHolder.Width / Screen.TwipsPerPixelX)
    ElseIf DockStatus = vbAlignRight Then
        tClientRect.Right = tClientRect.Right + (picHolder.Width / Screen.TwipsPerPixelX)
    ElseIf DockStatus = vbAlignTop Then
        tClientRect.Top = tClientRect.Top - (picHolder.Height / Screen.TwipsPerPixelX)
    ElseIf DockStatus = vbAlignBottom Then
        tClientRect.Bottom = tClientRect.Bottom + (picHolder.Height / Screen.TwipsPerPixelX)
    End If
    'ClientToScreen frmParent.hWnd, tP
    'tClientRect.Left = tClientRect.Left + tP.x
    'get point of MouseDown in screen coordinates
    tStartPoint = tPoint
    ClientToScreen Parent.hWnd, tStartPoint
    Docking = DockStatus
    'erase old rect
    SetRectEmpty tOldRect
    'do the drag
    DoFRDrag x, y
End Sub
Private Sub DoFRDrag(x As Single, y As Single)
    Dim tDc As Long
    Dim sDC As Long
    Dim d As Long
    Dim lHeight As Long, lwidth As Long, ltop As Long, lLeft As Long
    Dim tp As POINTAPI
    Dim tRect As RECT
    
    
    tp.x = x
    tp.y = y

    ClientToScreen Parent.hWnd, tp
    ltop = (tParentRect.Top * TwipsPerPixelY + tp.y) - tStartPoint.y
    lLeft = (tParentRect.Left * TwipsPerPixelX + tp.x) - tStartPoint.x
    
    tRect.Right = (tParentRect.Right * TwipsPerPixelX + tp.x) - tStartPoint.x
    tRect.Bottom = (tParentRect.Bottom * TwipsPerPixelY + tp.y) - tStartPoint.y
    lHeight = (tRect.Bottom - ltop)
    lwidth = (tRect.Right - lLeft)

    CheckRect lLeft, ltop, lwidth, lHeight

    tRect.Left = lLeft / TwipsPerPixelY
    tRect.Top = ltop / TwipsPerPixelX
    tRect.Right = (tRect.Left + lwidth / TwipsPerPixelY)
    tRect.Bottom = (tRect.Top + lHeight / TwipsPerPixelX)

    
    sDC = GetDC(ByVal 0)
'    If DockStatus <> vbAlignNone Then
'
'        DrawFocusRect sDC, tOldRect
'        DrawFocusRect sDC, tRect
'    Else
        Dim hBrush As Long
        
        Dim lSize As Long
        hBrush = gdiCreateHalftoneBrush()
        lSize = IIf(Docking = vbAlignNone, 5, 1)
        If m_lOldSize = 0 Then m_lOldSize = lSize
        gdiDrawDragRectangle sDC, tRect.Left, tRect.Top, tRect.Right - tRect.Left, tRect.Bottom - tRect.Top, tOldRect.Left, tOldRect.Top, tOldRect.Right - tOldRect.Left, tOldRect.Bottom - tOldRect.Top, lSize, m_lOldSize, hBrush, 0
        m_lOldSize = lSize
        'clean up
        gdiDeleteBrush hBrush
        
    'End If
    d = ReleaseDC(0, sDC)
    tOldRect = tRect
m_bDragging = True
End Sub
Private Sub CheckRect(lLeft As Long, ltop As Long, lwidth As Long, lHeight As Long)
   'Exit Sub
   Dim lWidthHidden As Long
   Dim lHeightHidden As Long
   Dim tClientR As RECT
   
   Dim tp As POINTAPI

    With frmDockWindow
        lWidthHidden = 200
        lHeightHidden = .ScaleHeight / 2
    End With
    If DockStatus <> vbAlignNone Then
        lwidth = GetSetting(m_sAppName, m_sSection, m_sKey & "_Width", m_DefaultWidth)
        lHeight = GetSetting(m_sAppName, m_sSection, m_sKey & "_Height", m_DefaultHeight)
    End If
'    If DockStatus = vbAlignBottom Or DockStatus = vbAlignTop Then
'        lLeft = lLeft + lwidth
'    End If
    
'    If DockStatus <> vbAlignNone And Docking = vbAlignNone Then
'        GetCursorPos tP
'        lLeft = tP.x * TwipsPerPixelX - 100
'        lTop = tP.y * TwipsPerPixelY - 100
'    End If
    'If PtInRect(tClientRect, ltop / Screen.TwipsPerPixelY, lLeft / Screen.TwipsPerPixelX) Or PtInRect(tClientRect, (ltop + lHeight) / Screen.TwipsPerPixelY, (lLeft + lwidth) / Screen.TwipsPerPixelX) Then
    If ltop < frmParent.Top + frmParent.Height And _
        ltop < frmParent.Top + frmParent.Height And _
        lLeft < frmParent.Left + frmParent.Width And _
        lLeft + Width > frmParent.Left Then
        ' if we are within the parents lLeft
        ' and our lLeft is greater than the parents lLeft + 1/2 the forms width...
       ' Debug.Print lLeft
        If (lLeft) > (tClientRect.Left * TwipsPerPixelX) - (lwidth / 2) And lLeft < (tClientRect.Left * Screen.TwipsPerPixelX) Then
        'If (lLeft) < (frmParent.Left + lWidthHidden) And lLeft + (lwidth / 2) > (frmParent.Left - lWidthHidden) Then
            'frmDocking.ShowDockSymbol frmParent.Left, frmParent.Top, (.ScaleWidth / 2) - 10, frmParent.Height
           ' lTop = picHolder.Top + frmParent.Top
             '* TwipsPerPixelX
            If DockStatus <> vbAlignLeft Then
                lwidth = lwidth / 2
                lHeight = frmParent.ScaleHeight
                'ltop = tClientRect.Top * TwipsPerPixelY
            End If
            
            ' docking lLeft...
            Docking = vbAlignLeft
            bDocking = True
        ' if we are within the parents right
        ' and less than 1/2 the forms width away...
        ElseIf lLeft + lwidth > (frmParent.Width + frmParent.Left - lWidthHidden) And _
            lLeft < (frmParent.Width + frmParent.Left) + lWidthHidden Then
            'frmDocking.ShowDockSymbol frmParent.Left + frmParent.Width - (.ScaleWidth / 2) - 10, frmParent.Top, ScaleWidth / 2, frmParent.Height
           ' lHeight = picHolder.ScaleHeight * TwipsPerPixelX
            If DockStatus <> vbAlignRight Then
                lHeight = frmParent.ScaleHeight
                lwidth = lwidth / 2
                'If lLeft < frmParent.Width + frmParent.Left - lWidthHidden Then lLeft = lLeft + lWidth
            End If
            ' docking right
            Docking = vbAlignRight
            bDocking = True
        ' if we are within the parents lTop
        ' and are less than 1/2 the forms height away...
        ElseIf ltop < frmParent.Top And _
            ltop > frmParent.Top - lHeightHidden Then
            'frmDocking.ShowDockSymbol frmParent.Left, frmParent.Top, frmParent.Width, ScaleHeight / 2
            If DockStatus <> vbAlignTop Then
                lHeight = lHeight / 2
                lwidth = frmParent.ScaleWidth
            End If
            ' docking lTop
            Docking = vbAlignTop
            bDocking = True
        ' if we are within the parents bottom edge
        ' and we are less than 1/2 the forms height away...
        'ElseIf (ltop + lHeight) > (tClientRect.bottom * TwipsPerPixelX) And _
        '    (ltop + lHeight) < frmParent.Top + frmParent.Height + lHeightHidden Then
        ElseIf (ltop + lHeight - 400) > (tClientRect.Bottom * TwipsPerPixelX) And _
            (ltop + lHeight) < frmParent.Top + frmParent.Height + lHeightHidden Then
            'frmDocking.ShowDockSymbol frmParent.Left, frmParent.Top + frmParent.Height - (.ScaleHeight / 2) - 10, frmParent.Width, ScaleHeight / 2
            If DockStatus = vbAlignBottom Or DockStatus = vbAlignTop Then
                lwidth = frmDockWindow.ScaleWidth
                'lLeft = (tClientRect.Left * TwipsPerPixelX)
                'lleft=
                'lLeft = lLeft - lwidth
            Else
                lHeight = lHeight / 2
                lwidth = (tClientRect.Right - tClientRect.Left) * TwipsPerPixelX
            End If
            
            ' docking bottom
            Docking = vbAlignBottom
            bDocking = True
        Else
            Docking = vbAlignNone
        End If
    Else
        Docking = vbAlignNone
    End If ' phew...!
End Sub

'Private Sub CheckRect(lLeft As Long, ltop As Long, lwidth As Long, lHeight As Long)
'   'Exit Sub
'   Dim lWidthHidden As Long
'   Dim lHeightHidden As Long
'   Dim tClientR As Rect
'
'   Dim tP As POINTAPI
'
'    With frmDockWindow
'        lWidthHidden = 200
'        lHeightHidden = .ScaleHeight / 2
'    End With
'    If DockStatus <> vbAlignNone Then
'        lwidth = GetSetting(m_sAppName, m_sSection, m_sKey & "_Width", m_DefaultWidth)
'        lHeight = GetSetting(m_sAppName, m_sSection, m_sKey & "_Height", m_DefaultHeight)
'    End If
''    If DockStatus = vbAlignBottom Or DockStatus = vbAlignTop Then
''        lLeft = lLeft + lwidth
''    End If
'
''    If DockStatus <> vbAlignNone And Docking = vbAlignNone Then
''        GetCursorPos tP
''        lLeft = tP.x * TwipsPerPixelX - 100
''        lTop = tP.y * TwipsPerPixelY - 100
''    End If
'    If PtInRect(tClientRect, ltop / Screen.TwipsPerPixelY, lLeft / Screen.TwipsPerPixelX) Or PtInRect(tClientRect, (ltop + lHeight) / Screen.TwipsPerPixelY, (lLeft + lwidth) / Screen.TwipsPerPixelX) Then
''    If ltop < frmParent.Top + frmParent.Height And _
''        ltop < frmParent.Top + frmParent.Height And _
''        lLeft < frmParent.Left + frmParent.Width And _
''        lLeft + Width > frmParent.Left Then
'        ' if we are within the parents lLeft
'        ' and our lLeft is greater than the parents lLeft + 1/2 the forms width...
'       ' Debug.Print lLeft
'        If (lLeft) > (tClientRect.Left * TwipsPerPixelX) - (lwidth / 2) And lLeft < (tClientRect.Left * Screen.TwipsPerPixelX) Then
'        'If (lLeft) < (frmParent.Left + lWidthHidden) And lLeft + (lwidth / 2) > (frmParent.Left - lWidthHidden) Then
'            'frmDocking.ShowDockSymbol frmParent.Left, frmParent.Top, (.ScaleWidth / 2) - 10, frmParent.Height
'           ' lTop = picHolder.Top + frmParent.Top
'             '* TwipsPerPixelX
'            If DockStatus <> vbAlignLeft Then
'                lwidth = lwidth / 2
'                lHeight = frmParent.ScaleHeight
'                ltop = tClientRect.Top * TwipsPerPixelY
'            End If
'
'            ' docking lLeft...
'            Docking = vbAlignLeft
'            bDocking = True
'        ' if we are within the parents right
'        ' and less than 1/2 the forms width away...
'        ElseIf lLeft + lwidth > (frmParent.Width + frmParent.Left - lWidthHidden) And _
'            lLeft < (frmParent.Width + frmParent.Left) + lWidthHidden Then
'            'frmDocking.ShowDockSymbol frmParent.Left + frmParent.Width - (.ScaleWidth / 2) - 10, frmParent.Top, ScaleWidth / 2, frmParent.Height
'           ' lHeight = picHolder.ScaleHeight * TwipsPerPixelX
'            If DockStatus <> vbAlignRight Then
'                lHeight = frmParent.ScaleHeight
'                lwidth = lwidth / 2
'                'If lLeft < frmParent.Width + frmParent.Left - lWidthHidden Then lLeft = lLeft + lWidth
'            End If
'            ' docking right
'            Docking = vbAlignRight
'            bDocking = True
'        ' if we are within the parents lTop
'        ' and are less than 1/2 the forms height away...
'        ElseIf ltop < frmParent.Top And _
'            ltop > frmParent.Top - lHeightHidden Then
'            'frmDocking.ShowDockSymbol frmParent.Left, frmParent.Top, frmParent.Width, ScaleHeight / 2
'            If DockStatus <> vbAlignTop Then
'                lHeight = lHeight / 2
'                lwidth = frmParent.ScaleWidth
'            End If
'            ' docking lTop
'            Docking = vbAlignTop
'            bDocking = True
'        ' if we are within the parents bottom edge
'        ' and we are less than 1/2 the forms height away...
'        ElseIf (ltop + lHeight) > (tClientRect.bottom * TwipsPerPixelX) And _
'            (ltop + lHeight) < frmParent.Top + frmParent.Height + lHeightHidden Then
'            'frmDocking.ShowDockSymbol frmParent.Left, frmParent.Top + frmParent.Height - (.ScaleHeight / 2) - 10, frmParent.Width, ScaleHeight / 2
'            If DockStatus = vbAlignBottom Or DockStatus = vbAlignTop Then
'                lwidth = frmDockWindow.ScaleWidth
'                lLeft = (tClientRect.Left * TwipsPerPixelX)
'                'lleft=
'                'lLeft = lLeft - lwidth
'            Else
'                lLeft = tClientRect.Left * TwipsPerPixelX 'frmParent.Left 'picHolder.Left '(tParentRect.Left * TwipsPerPixelX)
'                lHeight = lHeight / 2
'                lwidth = (tClientRect.Right - tClientRect.Left) * TwipsPerPixelX
'            End If
'
'            ' docking bottom
'            Docking = vbAlignBottom
'            bDocking = True
'        Else
'            Docking = vbAlignNone
'        End If
'    Else
'        Docking = vbAlignNone
'    End If ' phew...!
'    If Docking = vbAlignNone Then
'
'
'        GetCursorPos tP
'        If (tP.x * Screen.TwipsPerPixelX) > lLeft + lwidth Or (tP.x * Screen.TwipsPerPixelX) < lLeft Then
'            lLeft = (tP.x - 10) * Screen.TwipsPerPixelX
'            ltop = (tP.y - 10) * Screen.TwipsPerPixelX
'           ' tPoint.x = tp.x * Screen.TwipsPerPixelX
'        End If
'
'    End If
'End Sub

Private Sub EndFRDrag(x As Single, y As Single)
    
    Dim tDc As Long
    Dim sDC As Long
    Dim d As Long
    
    Dim newleft As Single
    Dim newtop As Single
    Dim hBrush As Long
    Dim tp As POINTAPI
    GetCursorPos tp
    sDC = GetDC(ByVal 0)
    hBrush = gdiCreateHalftoneBrush()
    
    gdiDrawDragRectangle sDC, tOldRect.Left, tOldRect.Top, tOldRect.Right - tOldRect.Left, tOldRect.Bottom - tOldRect.Top, 0, 0, 0, 0, m_lOldSize, m_lOldSize, hBrush, 0

    'clean up
    gdiDeleteBrush hBrush
    d = ReleaseDC(0, sDC)
        
'    sDC = GetDC(ByVal 0)
'    DrawFocusRect sDC, tOldRect
'    d = ReleaseDC(0, sDC)

    m_lOldSize = 0
    'drag aborted
    m_bDragging = False
    
    If x = -1 And y = -1 Then Exit Sub
'    newleft = (tp.x - 10) * TwipsPerPixelX
'    newtop = (tp.y - 10) * TwipsPerPixelY
    newleft = x + tParentRect.Left * TwipsPerPixelX - tPoint.x
        newtop = y + tParentRect.Top * TwipsPerPixelY - tPoint.y
    
    If DockStatus <> vbAlignNone And Docking = vbAlignNone Then
        Call MoveWindow(frmDockWindow.hWnd, newleft / TwipsPerPixelX, newtop / TwipsPerPixelY, (GetSetting(m_sAppName, m_sSection, m_sKey & "_Width", m_DefaultWidth) / TwipsPerPixelX), (GetSetting(m_sAppName, m_sSection, m_sKey & "_Height", m_DefaultHeight) / TwipsPerPixelY), True)
        UndockWindow False
    ElseIf Docking <> vbAlignNone And Docking <> DockStatus Then
        Dock Docking
    ElseIf Docking = vbAlignNone Then

        frmDockWindow.Move newleft, newtop
    End If
End Sub
'/ VB Web Dockable Window Control
' Please visit www.vbweb.co.uk
'
' This Control can be used free of charge in any freeware project
' Please email james@vbweb.co.uk if you intend to use the Control
' in a commercial program. Credit would be appreciated!
'
' Please DO NOT make modifications to this Control
' If you can improve the code, please email
' james@vbweb.co.uk and it will be encorporated
' into the next release. Thanks.
'
' Parts of this code is from vbaccelerator.com



Public Property Let RegAppName(sAppName As String)
    m_sAppName = sAppName
End Property
Public Property Get RegAppName() As String
    RegAppName = m_sAppName
End Property
Public Property Let RegSection(sSection As String)
    m_sSection = sSection
End Property
Public Property Get RegSection() As String
    RegSection = m_sSection
End Property
Public Property Let RegKey(sKey As String)
    m_sKey = sKey
End Property
Public Property Get RegKey() As String
    RegKey = m_sKey
End Property
Public Property Let DefaultPos(Pos As AlignConstants)
    m_DefaultPos = Pos
End Property
Public Property Get DefaultPos() As AlignConstants
    DefaultPos = m_DefaultPos
End Property
Public Property Let DefaultWidth(Width As Single)
    m_DefaultWidth = Width
End Property
Public Property Get DefaultWidth() As Single
    DefaultWidth = m_DefaultWidth
End Property
Public Property Let DefaultHeight(Height As Single)
    m_DefaultHeight = Height
End Property
Public Property Get DefaultHeight() As Single
    DefaultHeight = m_DefaultHeight
End Property

Private Sub UserControl_Initialize()
    RegSection = "WindowSettings"
    DefaultPos = vbAlignLeft
    DefaultWidth = 2000
    DefaultHeight = 3000
    'screen RECT of toolbar
    TwipsPerPixelX = Screen.TwipsPerPixelX
    TwipsPerPixelY = Screen.TwipsPerPixelY
    vLastPointer = -1
End Sub
Private Sub frmDockWindow_Activate()
    Call SendMessage(frmParent.hWnd, WM_NCACTIVATE, 1, 0)
End Sub
Private Sub frmDockWindow_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then m_bUnloading = True
End Sub
Private Sub picHolder_Resize()
    If picHolder.Visible = False And m_bStartUp = False Then Exit Sub
    If m_bIgnore = True Then Exit Sub
    ' update height
    'm_bIgnore = True
   ' Call MoveWindow(frmDockWindow.hwnd, -4, -4, (picHolder.Width / TwipsPerPixelX) + 8, picHolder.Height / TwipsPerPixelY + 8, True)
   ' Call MoveWindow(frmDockWindow.hwnd, 0, 0, (picHolder.Width / TwipsPerPixelX), frmDockWindow.Height / TwipsPerPixelY + 8, True)
   ' m_bIgnore = False
    
    Select Case DockStatus
    Case vbAlignLeft
        Call MoveWindow(frmDockWindow.hWnd, 0, 0, (picHolder.Width / TwipsPerPixelX) - 5, picHolder.Height / TwipsPerPixelY, True)
        'MoveWindow picSizer.hwnd, (picHolder.Width / TwipsPerPixelX) - 5, 0, 5, (picHolder.Height / TwipsPerPixelY), True
    Case vbAlignRight
        Call MoveWindow(frmDockWindow.hWnd, 5, 0, (picHolder.Width / TwipsPerPixelX) - 5, picHolder.Height / TwipsPerPixelY, True)
        'MoveWindow picSizer.hwnd, 0, 0, 5, (picHolder.Height / TwipsPerPixelY), True
    Case vbAlignTop
        Call MoveWindow(frmDockWindow.hWnd, 0, 0, (picHolder.Width / TwipsPerPixelX), (picHolder.Height / TwipsPerPixelY) - 5, True)
        'MoveWindow picSizer.hwnd, 0, (picHolder.Height / TwipsPerPixelX) - 5, (picHolder.Width / TwipsPerPixelY), 5, True
    Case vbAlignBottom
        Call MoveWindow(frmDockWindow.hWnd, 0, 5, (picHolder.Width / TwipsPerPixelX), (picHolder.Height / TwipsPerPixelY) - 5, True)
        'MoveWindow picSizer.hwnd, 0, 0, (picHolder.Width / TwipsPerPixelY), 5, True
    End Select
'
'    If DockStatus = vbAlignRight Or DockStatus = vbAlignLeft Then
'        'SetParent picSizer.hwnd, picHolder.hwnd
'        MoveWindow picSizer.hwnd, (picHolder.Width / TwipsPerPixelX) - 5, 0, 5, (picHolder.Height / TwipsPerPixelY), True
'       ' picSizer.Left = picHolder.Width - 5
'    End If
    Call UpdateWindow(frmDockWindow.hWnd)
    'Call UpdateWindow(picSizer.hwnd)
End Sub
Public Sub Initialize()
   ' AttachMessages
    If RegKey = "" Then RegKey = frmDockWindow.Name
    picHolder.Visible = False
    m_bDockWindowLoaded = True
End Sub
Private Sub DockWindow()
    frmDockWindow.Hide
    SaveWindowDimen
    m_bFloating = False
    Call SetParent(frmDockWindow.hWnd, picHolder.hWnd)
    frmDockWindow.Tag = "DOCKED"
    frmDockWindow.BorderStyle = 4
    picHolder.Visible = True
    DoEvents
    frmDockWindow.Show
End Sub
Private Sub SaveWindowDimen()
    If m_bStartUp Then Exit Sub
    ' align left or right (3 or 4)
    If DockStatus > 2 Or m_bFloating Then
        '
        SaveSetting m_sAppName, m_sSection, m_sKey & "_Width", frmDockWindow.Width
    End If
    ' align top or bottom (1 or 2)
    If DockStatus <= 2 And DockStatus <> 0 Or m_bFloating Then
        SaveSetting m_sAppName, m_sSection, m_sKey & "_Height", frmDockWindow.Height
    End If
    If m_bFloating Then
        SaveSetting m_sAppName, m_sSection, m_sKey & "_Left", frmDockWindow.Left
        SaveSetting m_sAppName, m_sSection, m_sKey & "_Top", frmDockWindow.Top
    End If
End Sub
'Private Sub AttachMessages()
'    AttachMessage Me, frmDockWindow.hwnd, WM_SYSCOMMAND
'    AttachMessage Me, frmDockWindow.hwnd, WM_MOVE
'    AttachMessage Me, frmDockWindow.hwnd, WM_NCLBUTTONDOWN
'    AttachMessage Me, frmDockWindow.hwnd, WM_EXITSIZEMOVE
'End Sub
'Private Sub DetachMessages()
'    DetachMessage Me, frmDockWindow.hwnd, WM_SYSCOMMAND
'    DetachMessage Me, frmDockWindow.hwnd, WM_MOVE
'    DetachMessage Me, frmDockWindow.hwnd, WM_NCLBUTTONDOWN
'    DetachMessage Me, frmDockWindow.hwnd, WM_EXITSIZEMOVE
'End Sub
Private Sub UserControl_Terminate()
    If Not picHolder Is Nothing Then
        SaveWindowPos
        picHolder.Visible = False
        'DetachMessages
    End If
End Sub
'Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
'
'End Property
'Private Property Get ISubclass_MsgResponse() As EMsgResponse
'    ISubclass_MsgResponse = emrPostProcess
'End Property
'Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'    On Error Resume Next
'    Static bButtonDown As Boolean
'    Static lWidth As Long
'    Static lHeight As Long
'    Dim lWidthHidden As Long
'    Dim lHeightHidden As Long
'    Dim bDocking As Boolean
''    With frmDockWindow
''        Select Case iMsg
''        Case WM_SYSCOMMAND
''            If wParam = 61458 Then
''                If .Tag = "DOCKED" Then
''                    UndockWindow
''                End If
''            End If
''        End Select
''    End With
'End Function
Public Sub SetNoBorder(lhWnd As Long)
Dim lStyle As Long
    If lOldStyle = 0 Then
        lOldStyle = GetWindowLong(lhWnd, GWL_STYLE)
        lStyle = lOldStyle And Not (WS_BORDER Or WS_DLGFRAME Or WS_CAPTION Or WS_SIZEBOX Or WS_THICKFRAME)
        'Apply changes
        SetWindowLong lhWnd, GWL_STYLE, lStyle
    End If
End Sub
Public Sub RestoreBorder(lhWnd As Long)
Dim lStyle As Long
Dim cRect As RECT
    GetWindowRect lhWnd, cRect
    If lOldStyle <> 0 Then
        'undo no-border changes
        SetWindowLong lhWnd, GWL_STYLE, lOldStyle
        'Apply changes... why can't we do this without specifying the
        'window location?!
        'SetWindowPosAPI lhWnd, 0, cRect.Left, cRect.Top, cRect.Right, cRect.Bottom, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
        FrameChanged lhWnd
        lOldStyle = 0
    End If
End Sub
Public Sub Dock(Align As AlignConstants, Optional bNoCancel As Boolean = False)
Dim bCancel As Boolean
    If bNoCancel = False Then RaiseEvent BeforeDock(bCancel)
    If bCancel = True Then Exit Sub
    picHolder.Visible = False
    m_bIgnore = True
    If Align = vbAlignNone Then
        UndockWindow True
    Else
        'picSizer.Visible = True
        With frmDockWindow
            picHolder.Align = Align
            If Align = vbAlignLeft Or Align = vbAlignRight Then
                ' Docking to the sides
                picHolder.Width = .Width + (5 * TwipsPerPixelX)
                'picSizer.MousePointer = vbSizeWE
                If picHolder.ScaleWidth > (frmParent.Width / 2) Then
                    picHolder.Width = frmParent.Width / 2
                    .Width = picHolder.ScaleWidth
                End If
            ElseIf Align = vbAlignTop Or Align = vbAlignBottom Then
                ' Docking to the top or bottom
                picHolder.Height = .Height + (5 * TwipsPerPixelY)
                'picSizer.MousePointer = vbSizeNS
                
                If picHolder.ScaleHeight > (frmParent.ScaleHeight / 2) Then
                    picHolder.Height = frmParent.ScaleHeight / 2
                    .Height = picHolder.ScaleHeight
                End If
            End If
            m_bIgnore = False
             SetNoBorder frmDockWindow.hWnd
            If .Tag <> "DOCKED" Then DockWindow
           
            picHolder.Visible = True
          '  Unload frmDocking
        End With
        DockStatus = picHolder.Align
        SaveWindowDimen
    End If
    m_bIgnore = False
    picHolder_Resize
    RaiseEvent Dock(Align)
End Sub

Private Sub tmrUpdate_Timer()
    If m_bRunTime = False Then Exit Sub
    Dim lActive As Long
    lActive = GetActiveWindow
    If lActive <> frmDockWindow.hWnd And m_bDragging Then
        AbortDrag
    ElseIf lActive <> frmParent.hWnd And bSplitting Then
        'abort splitting
        bSplitting = False
        DrawSplitter
    End If
    If MouseInSizer And picHolder.Visible = True Then
    ElseIf bSplitting = False Then
        If vLastPointer <> -1 Then Screen.MousePointer = vLastPointer
        vLastPointer = -1
    End If
End Sub
Public Sub UndockWindow(Optional bMoveToLast As Boolean = False, Optional bNoCancel As Boolean = False)
Dim bCancel As Boolean
    If bNoCancel = False Then RaiseEvent BeforeUnDock(bCancel)
    If bCancel Then Exit Sub
    m_bFloating = True
    DockStatus = vbAlignNone
    ' sort out the sizer pic box
  '  picSizerTop.MousePointer = vbDefault
  '  picSizerTop.Visible = False
  '  picSizer.Visible = False
    Dim p As POINTAPI
    With frmDockWindow
        RestoreBorder frmDockWindow.hWnd
        ' hide the parent container + form
        picHolder.Visible = False
        .Hide
        SetParent .hWnd, 0
       ' Call GetCursorPos(p)
        ' set form to resizable (doesn't work in VB5, sorry!)
       ' .BorderStyle = 5
        ' save the forms width or height
        ' (whichever side is not extended because of being docked
        If picHolder.Align = vbAlignLeft Or picHolder.Align = vbAlignRight Then
            SaveSetting m_sAppName, m_sSection, m_sKey & "_Width", frmDockWindow.Width
        ElseIf picHolder.Align = vbAlignBottom Or picHolder.Align = vbAlignTop Then
            SaveSetting m_sAppName, m_sSection, m_sKey & "_Height", frmDockWindow.Height
        End If
        ' move the window to where the cursor is
        ' so we can continue moving
        'Call MoveWindow(.hwnd, p.x - 10, p.y - 10, (GetSetting(m_sAppName, m_sSection, m_sKey & "_Width", m_DefaultWidth) / TwipsPerPixelX), (GetSetting(m_sAppName, m_sSection, m_sKey & "_Height", m_DefaultHeight) / TwipsPerPixelY), True)
        'Call UpdateWindow(frmParent.hwnd)
        .Tag = Empty
        ' restore last pos
        If bMoveToLast Then
            .Left = GetSetting(m_sAppName, m_sSection, m_sKey & "_Left", .Left)
            .Top = GetSetting(m_sAppName, m_sSection, m_sKey & "_Top", .Top)
        End If
        .Show , frmParent
        Call UpdateWindow(.hWnd)
        Call SendMessage(frmParent.hWnd, WM_NCACTIVATE, 1, 0)
    End With
    RaiseEvent UnDock
End Sub
Public Sub ShowWindow(bInitialPos As Boolean)
    Dim bWindowShowing As Boolean
    If frmDockWindow.Visible Then Exit Sub
    If m_bDockWindowLoaded = False Then Initialize
    m_bStartUp = bInitialPos
    bWindowShowing = SetWindowPos(bInitialPos)
    m_bStartUp = False
    ' if the window has been positioned, and needs
    ' to be shown, show it!
    If bWindowShowing Then frmDockWindow.Show
End Sub
Private Function SetWindowPos(bInitialPos As Boolean) As Boolean
On Error Resume Next
    Dim intPos As AlignConstants
    SetWindowPos = True
    With frmDockWindow
        If bInitialPos Then
            intPos = Val(GetSetting(m_sAppName, m_sSection, m_sKey & "_Pos", m_DefaultPos))
        Else
            intPos = Val(GetSetting(m_sAppName, m_sSection, m_sKey & "_OldPos", m_DefaultPos))
        End If
        Select Case intPos
        Case 3, 4 ' Left, Right
            .Width = GetSetting(m_sAppName, m_sSection, m_sKey & "_Width", m_DefaultWidth)
            Dock intPos, True
        Case 1, 2 ' top, bottom
            .Height = GetSetting(m_sAppName, m_sSection, m_sKey & "_Height", m_DefaultHeight)
            Dock intPos, True
        Case 0 ' floating
            .Width = GetSetting(m_sAppName, m_sSection, m_sKey & "_Width", m_DefaultWidth)
            .Height = GetSetting(m_sAppName, m_sSection, m_sKey & "_Height", m_DefaultHeight)
            UndockWindow True, True
        Case 5 ' hidden, switch to default
            If bInitialPos = False Then
                .Width = GetSetting(m_sAppName, m_sSection, m_sKey & "_Width", m_DefaultWidth)
                Dock 3, True
            Else
                SetWindowPos = False
            End If
        End Select
    End With
End Function
'Private Sub picSizer_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    ' Sizing ideas for m_Docking Window from
'    ' Peter Siebler [p.siebler@cable.vol.at]
'    If Button = vbLeftButton Then
'        ' As a simple alternative to using a splitter control we can resize
'        ' frmDocking.Picture1 by sending a message that will make windows draw a
'        ' resizing border for us.
'        Dim nSide As Integer
'        If picHolder.Visible Then
'            ' need to do this or SendMessage fails
'            ReleaseCapture
'            ' Send message to start resizing frmDocking.Picture1
'            Select Case picHolder.Align
'            Case vbAlignLeft
'                nSide = HTRIGHT
'            Case vbAlignRight
'                nSide = HTLEFT
'            Case vbAlignTop
'                nSide = HTBOTTOM
'            Case vbAlignBottom
'                nSide = HTTOP
'            End Select
'            SendMessage picHolder.hwnd, WM_NCLBUTTONDOWN, nSide, ByVal &O0
'            If picHolder.Width < 500 Then picHolder.Width = 500
'            If picHolder.Height < 500 Then picHolder.Height = 500
'            picHolder_Resize
'        End If
'    End If
'End Sub

Private Sub picHolder_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Sizing ideas for m_Docking Window from
    ' Peter Siebler [p.siebler@cable.vol.at]
    If MouseInSizer And Button = vbLeftButton Then
        ' As a simple alternative to using a splitter control we can resize
        ' frmDocking.Picture1 by sending a message that will make windows draw a
        ' resizing border for us.
        Dim nSide As Integer
        If picHolder.Visible Then
            ' need to do this or SendMessage fails
           ' ReleaseCapture
            ' Send message to start resizing frmDocking.Picture1
            Select Case picHolder.Align
            Case vbAlignLeft
                nSide = HTRIGHT
            Case vbAlignRight
                nSide = HTLEFT
            Case vbAlignTop
                nSide = HTBOTTOM
            Case vbAlignBottom
                nSide = HTTOP
            End Select
            'start dragging
            bSplitting = True
            DrawSplitter
'            SendMessage picHolder.hwnd, WM_NCLBUTTONDOWN, nSide, ByVal &O0
'            If picHolder.Width < 500 Then picHolder.Width = 500
'            If picHolder.Height < 500 Then picHolder.Height = 500
'            picHolder_Resize
        End If
    End If
End Sub
Private Function DrawSplitter() As Long
    Dim tp As POINTAPI
    Dim tR As RECT
    Dim tParent As RECT
    Dim hdc As Long
    Dim hBrush As Long
    Dim lMin As Long
    Dim tOwner As RECT
    Static tOldRect As RECT
    
    GetCursorPos tp
    GetWindowRect picHolder.hWnd, tParent
    GetWindowRect GetWindow(frmParent.hWnd, GW_CHILD), tOwner
    If picHolder.Align = vbAlignLeft Then
        If tp.x < tParent.Left + 30 Then
            tp.x = tParent.Left + 30
        ElseIf tp.x > tOwner.Right - 30 Then
            tp.x = tOwner.Right - 30
        End If
    ElseIf picHolder.Align = vbAlignRight Then
        If tp.x > tParent.Right - 30 Then
            tp.x = tParent.Right - 30
        ElseIf tp.x < tOwner.Left + 30 Then
            tp.x = tOwner.Left + 30
        End If
    ElseIf picHolder.Align = vbAlignTop Then
        If tp.y > tOwner.Bottom - 30 Then
            tp.y = tOwner.Bottom - 30
        ElseIf tp.y < tParent.Top + 30 Then
            tp.y = tParent.Top + 30
        End If
    ElseIf picHolder.Align = vbAlignBottom Then
        If tp.y < tOwner.Top + 30 Then
            tp.y = tOwner.Top + 30
        ElseIf tp.y > tParent.Bottom - 30 Then
            tp.y = tParent.Bottom - 30
        End If
    End If

    If bSplitting Then
        If picHolder.Align = vbAlignLeft Or picHolder.Align = vbAlignRight Then
            tR.Left = tp.x
            tR.Top = tParent.Top
            tR.Bottom = tParent.Bottom
            tR.Right = tR.Left + 5
        Else
            tR.Left = tParent.Left
            tR.Top = tp.y
            tR.Bottom = tR.Top + 5
            tR.Right = tParent.Right
        End If
    End If
    
    hdc = GetDC(ByVal 0)
    
    'create halftone brush
    hBrush = gdiCreateHalftoneBrush()
    
    'DrawFocusRect hDC, tOldRect
    'gdiDrawSolidRectangleBrush hDC, tOldRect.Left, tOldRect.Top, tOldRect.Right - tOldRect.Left, tOldRect.Bottom - tOldRect.Top, hBrush
    'If bSplitting Then DrawFocusRect hDC, tR
    
    If bSplitting Then
        gdiDrawDragRectangle hdc, tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, tOldRect.Left, tOldRect.Top, tOldRect.Right - tOldRect.Left, tOldRect.Bottom - tOldRect.Top, 5, 5, hBrush, 0
    Else
        gdiDrawDragRectangle hdc, tOldRect.Left, tOldRect.Top, tOldRect.Right - tOldRect.Left, tOldRect.Bottom - tOldRect.Top, 0, 0, 0, 0, 5, 5, hBrush, 0
    End If
    'clean up
    gdiDeleteBrush hBrush
    
    ReleaseDC 0, hdc
    LSet tOldRect = tR
    ScreenToClient picHolder.hWnd, tp
    If picHolder.Align = vbAlignLeft Or picHolder.Align = vbAlignRight Then
        DrawSplitter = tp.x
    Else
        DrawSplitter = tp.y
    End If
End Function

    
Private Function MouseInSizer() As Boolean
    Dim tR As RECT
    Dim tp As POINTAPI
    Dim tParent As RECT
    If m_bRunTime = False Then Exit Function
    
    GetWindowRect picHolder.hWnd, tParent
    LSet tR = tParent
    Select Case picHolder.Align
    Case vbAlignLeft
        tR.Left = tParent.Right - 5
        tR.Right = tR.Left + 5
    Case vbAlignRight
        tR.Left = tParent.Left ' - 5
        tR.Right = tR.Left + 5
    Case vbAlignTop
        tR.Top = tParent.Bottom - 5
        tR.Bottom = tR.Top + 5
    Case vbAlignBottom
        tR.Top = tParent.Top
        tR.Bottom = tR.Top + 5
    End Select
    GetCursorPos tp
    MouseInSizer = PtInRect(tR, tp.x, tp.y)
End Function
Private Sub SaveWindowPos()
    If picHolder Is Nothing Then Exit Sub
    'If blnClosing Then Exit Sub
    Dim intPos As Integer
    Dim iPos As Integer
    Dim iOldPos As Integer
    iOldPos = -1
    If m_bUnloading Then
       
        If picHolder.Visible = False Then
            iOldPos = 0 ' floating (vbAlignNone)
        Else
            iOldPos = picHolder.Align  ' Docked left, right, top or bottom
        End If
     
       ' SaveSetting m_sAppName, m_sSection, m_sKey & "_OldPos", intPos
        iPos = 5 ' hidden
    ElseIf picHolder.Visible = False Then
        iPos = 0 ' floating (vbAlignNone)
    Else
        iPos = picHolder.Align  ' Docked left, right, top or bottom
    End If
    If iOldPos = -1 Then iOldPos = iPos
    
    SaveSetting m_sAppName, m_sSection, m_sKey & "_OldPos", iOldPos
    
    SaveSetting m_sAppName, m_sSection, m_sKey & "_Pos", iPos
    
    SaveWindowDimen
End Sub

















Public Sub gdiDrawDragRectangle(ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal DX As Long, ByVal DY As Long, ByVal LastX As Long, ByVal LastY As Long, ByVal LastDX As Long, ByVal LastDY As Long, ByVal Size As Long, ByVal LastSize As Long, ByVal hBrush As Long, ByVal hBrushLast As Long, Optional ByVal hClipRgn As Long)
'draw drag rectangle, erasing old as needed

    Dim tRect As RECT, tRectLast As RECT
    Dim hOutsideRgn As Long, hInsideRgn As Long
    Dim hLastRgn As Long, hUpdateRgn As Long
    Dim hNewRgn As Long
    Dim tNewRect As RECT
    Dim hOldBrush As Long
    
    Debug.Assert hBrush <> 0
    
    'init vars
    With tRect
        .Left = x
        .Top = y
        .Right = x + DX
        .Bottom = y + DY
        
    End With
    With tRectLast
        .Left = LastX
        .Top = LastY
        .Right = LastX + LastDX
        .Bottom = LastY + LastDY
    
    End With
    
    'determine update region and select it
    hOutsideRgn = apiCreateRectRgnIndirect(tRect)
    apiCopyRect tNewRect, tRect
    apiInflateRect tNewRect, -Size, -Size
    apiIntersectRect tNewRect, tNewRect, tRect
    hInsideRgn = apiCreateRectRgnIndirect(tNewRect)
    hNewRgn = apiCreateRectRgn(0, 0, 0, 0)
    apiCombineRgn hNewRgn, hOutsideRgn, hInsideRgn, RGN_XOR
    If hBrushLast = 0 Then
        hBrushLast = hBrush
        
    End If
    If apiIsRectEmpty(tRectLast) = 0 Then
        'find difference between new region and old region
        hLastRgn = apiCreateRectRgn(0, 0, 0, 0)
        apiSetRectRgn hOutsideRgn, tRectLast.Left, tRectLast.Top, tRectLast.Right, tRectLast.Bottom
        apiCopyRect tNewRect, tRectLast
        apiInflateRect tNewRect, -LastSize, -LastSize
        apiIntersectRect tNewRect, tNewRect, tRectLast
        apiSetRectRgn hInsideRgn, tNewRect.Left, tNewRect.Top, tNewRect.Right, tNewRect.Bottom
        apiCombineRgn hLastRgn, hOutsideRgn, hInsideRgn, RGN_XOR
        'only diff them if brushes are the same
        If (hBrush = hBrushLast) Then
            hUpdateRgn = apiCreateRectRgn(0, 0, 0, 0)
            apiCombineRgn hUpdateRgn, hLastRgn, hNewRgn, RGN_XOR
            
        End If
    
    End If
    If (hBrush <> hBrushLast) And (apiIsRectEmpty(tRectLast) = 0) Then
        'brushes are different -- erase old region first
        apiSelectClipRgn hdc, hLastRgn
        apiGetClipBox hdc, tNewRect
        hOldBrush = apiSelectObject(hdc, hBrushLast)
        apiPatBlt hdc, tNewRect.Left, tNewRect.Top, tNewRect.Right - tNewRect.Left, tNewRect.Bottom - tNewRect.Top, vbPatInvert
        apiSelectObject hdc, hOldBrush
        apiDeleteObject hOldBrush
        
    End If
    
    'draw into the update/new region
    If hUpdateRgn <> 0 Then
        If hClipRgn <> 0 Then
            apiCombineRgn hUpdateRgn, hClipRgn, hUpdateRgn, RGN_AND
        
        End If
        apiSelectClipRgn hdc, hUpdateRgn
        
    Else
        If hClipRgn <> 0 Then
            apiCombineRgn hNewRgn, hClipRgn, hNewRgn, RGN_AND
        
        End If
        apiSelectClipRgn hdc, hNewRgn
    
    End If
    apiGetClipBox hdc, tNewRect
    hOldBrush = apiSelectObject(hdc, hBrush)
    apiPatBlt hdc, tNewRect.Left, tNewRect.Top, tNewRect.Right - tNewRect.Left, tNewRect.Bottom - tNewRect.Top, vbPatInvert

    'cleanup DC
    If (hOldBrush <> 0) Then
        apiSelectObject hdc, hOldBrush
        
    End If
    apiSelectClipRgn hdc, 0&
    
    'delete objects
    If hUpdateRgn <> 0 Then
        apiDeleteObject hUpdateRgn
        
    End If
    If hLastRgn <> 0 Then
        apiDeleteObject hLastRgn
        
    End If
    If hNewRgn <> 0 Then
        apiDeleteObject hNewRgn
        
    End If
    If hInsideRgn <> 0 Then
        apiDeleteObject hInsideRgn
        
    End If
    If hOutsideRgn <> 0 Then
        apiDeleteObject hOutsideRgn
        
    End If
End Sub
Public Sub gdiDrawSolidRectangleBrush(ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal DX As Long, ByVal DY As Long, ByVal hBrush As Long)
'draw filled rectangle using supplied brush

    Dim tRect As RECT
    
    'build RECT
    gdiSetRect tRect, x, y, DX, DY
    
    'draw filled rect
    apiFillRect hdc, tRect, hBrush
End Sub
Private Sub gdiSetRect(tRect As RECT, ByVal x As Long, ByVal y As Long, ByVal DX As Long, ByVal DY As Long)
    
    'build RECT
    With tRect
        .Left = x
        .Top = y
        .Right = x + DX
        .Bottom = y + DY
        
    End With
End Sub
Public Sub gdiDeleteBrush(ByVal hBrush As Long)
'delete brush object

    apiDeleteObject hBrush
End Sub
Public Function gdiCreateHalftoneBrush() As Long
'create a halftoned brush object

    Dim nGreyPattern(8) As Integer
    Dim nBitCount As Integer
    Dim hBitmap As Long
    Dim lngResult As Long
    
    'create greyed bitmap
    For nBitCount = 0 To 7
        nGreyPattern(nBitCount) = LShiftWord(&H5555, (nBitCount And 1))
        
    Next nBitCount
    hBitmap = apiCreateBitmap(8, 8, 1, 1, nGreyPattern(0))
    
    'create halftone pattern brush
    gdiCreateHalftoneBrush = apiCreatePatternBrush(hBitmap)
    
    'delete bitmap
    apiDeleteObject hBitmap
End Function
Private Function LShiftWord(ByVal nWord As Integer, ByVal nBitCount As Integer) As Integer
'left shift dword value lngDWord by nBitCount
    
    Dim lngDWord As Long

    Debug.Assert (nBitCount >= 0 And nBitCount <= 15)   'design-time only
    If Not (nBitCount >= 0 And nBitCount <= 15) Then Exit Function
    
    lngDWord = nWord * Power2(nBitCount)
    If lngDWord And &H8000& Then
        LShiftWord = CInt(lngDWord And &H7FFF&) Or &H8000
        
    Else
        LShiftWord = lngDWord And &HFFFF&
        
    End If
End Function
Private Function Power2(ByVal nIndex As Integer) As Long
'build lookup table for bitshifting operations

    If Not (nIndex >= 0 And nIndex <= 31) Then Exit Function
    
    If m_lngPower2(0) = 0 Then 'only set array if not yet been used
        m_lngPower2(0) = &H1&
        m_lngPower2(1) = &H2&
        m_lngPower2(2) = &H4&
        m_lngPower2(3) = &H8&
        m_lngPower2(4) = &H10&
        m_lngPower2(5) = &H20&
        m_lngPower2(6) = &H40&
        m_lngPower2(7) = &H80&
        m_lngPower2(8) = &H100&
        m_lngPower2(9) = &H200&
        m_lngPower2(10) = &H400&
        m_lngPower2(11) = &H800&
        m_lngPower2(12) = &H1000&
        m_lngPower2(13) = &H2000&
        m_lngPower2(14) = &H4000&
        m_lngPower2(15) = &H8000&
        m_lngPower2(16) = &H10000
        m_lngPower2(17) = &H20000
        m_lngPower2(18) = &H40000
        m_lngPower2(19) = &H80000
        m_lngPower2(20) = &H100000
        m_lngPower2(21) = &H200000
        m_lngPower2(22) = &H400000
        m_lngPower2(23) = &H800000
        m_lngPower2(24) = &H1000000
        m_lngPower2(25) = &H2000000
        m_lngPower2(26) = &H4000000
        m_lngPower2(27) = &H8000000
        m_lngPower2(28) = &H10000000
        m_lngPower2(29) = &H20000000
        m_lngPower2(30) = &H40000000
        m_lngPower2(31) = &H80000000
        
    End If
    Power2 = m_lngPower2(nIndex)
End Function

