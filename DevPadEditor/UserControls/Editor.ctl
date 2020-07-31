VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.UserControl Editor 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   6630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7290
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   6630
   ScaleWidth      =   7290
   Begin VB.Timer tmrInsert 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1440
      Top             =   855
   End
   Begin VB.PictureBox picLines 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1455
      ScaleWidth      =   735
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   345
      Width           =   735
   End
   Begin VB.PictureBox picTB 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   7290
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   7290
      Begin VB.ComboBox cboLanguage 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "Editor.ctx":0000
         Left            =   285
         List            =   "Editor.ctx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   15
         Width           =   1440
      End
      Begin VB.ComboBox cboProcedures 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "Editor.ctx":0004
         Left            =   2115
         List            =   "Editor.ctx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   15
         Width           =   1995
      End
      Begin VB.Line LineSep 
         BorderColor     =   &H80000010&
         X1              =   300
         X2              =   6360
         Y1              =   345
         Y2              =   345
      End
      Begin VB.Image imgPic 
         Height          =   240
         Index           =   0
         Left            =   1815
         Picture         =   "Editor.ctx":0027
         ToolTipText     =   "Jump to..."
         Top             =   45
         Width           =   240
      End
      Begin VB.Image imgPic 
         Height          =   240
         Index           =   1
         Left            =   30
         Picture         =   "Editor.ctx":0171
         ToolTipText     =   "Language"
         Top             =   45
         Width           =   240
      End
   End
   Begin VB.Timer tmrColor 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2985
      Top             =   1635
   End
   Begin VB.PictureBox picFlag 
      AutoRedraw      =   -1  'True
      Height          =   285
      Left            =   4185
      Picture         =   "Editor.ctx":02BB
      ScaleHeight     =   225
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   990
      Visible         =   0   'False
      Width           =   315
   End
   Begin RichTextLib.RichTextBox rtfMain 
      Height          =   2880
      Left            =   2610
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   990
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   5080
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      Appearance      =   0
      RightMargin     =   1e6
      OLEDragMode     =   0
      OLEDropMode     =   1
      TextRTF         =   $"Editor.ctx":0405
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtfTemp 
      Height          =   2880
      Left            =   4200
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   5080
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      Appearance      =   0
      RightMargin     =   1e6
      TextRTF         =   $"Editor.ctx":0485
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait...."
      Height          =   255
      Left            =   780
      TabIndex        =   8
      Top             =   480
      Width           =   2385
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1073"
      Height          =   240
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   960
   End
End
Attribute VB_Name = "Editor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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

Implements ISubclass
Private Const REG_KEY = "DevPad"
Public Enum ModifyTypes
    AddText = 0
    DeleteText = 1
    ReplaceText = 2
    CutText = 3
    PasteText = 4
    IndentText = 5
    OutdentText = 6
    CommentText = 7
    UnCommentText = 8
End Enum
Public Enum ERECViewModes
    ercDefault = 0
    ercWordWrap = 1
    ercWYSIWYG = 2
End Enum
Public Enum ColourStatus
    InTag = 0
    OutTag = 1
    InComment = 2
    OutComment = 3
    InScript = 4
    InHTMLExtension = 5
End Enum
'*** Constants ***
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_COMMAND = &H111
Private Const WM_USER = &H400
Private Const WM_SETREDRAW = &HB
Private Const WM_VSCROLL = &H115
Private Const CB_SETDROPPEDWIDTH = &H160&

Private Const PS_SOLID = 0
Private Const WM_KEYUP = &H101
Private Const WM_KEYDOWN = &H100
Private Const CB_SHOWDROPDOWN = &H14F
Private Const EM_EXGETSEL = (WM_USER + 52)
Private Const EM_GETSCROLLPOS = (WM_USER + 221)
Private Const EM_SETSCROLLPOS = (WM_USER + 222)
Private Const EM_SETEVENTMASK = (WM_USER + 69)
Private Const CB_GETDROPPEDSTATE = &H157
Private Const EM_CANPASTE = (WM_USER + 50)
Private Const WM_COPY = &H301
Private Const WM_CUT = &H300
Private Const WM_PASTE = &H302
Private Const WM_MOUSEWHEEL = &H20A '#define WM_MOUSEWHEEL                   0x020A
' Text Constants
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_LINELENGTH = &HC1
Private Const EM_LINEINDEX = &HBB
Private Const EM_GETTEXTRANGE = (WM_USER + 75)
Private Const EM_SETTEXTMODE = (WM_USER + 89)
Private Const WM_GETTEXTLENGTH = &HE

Private Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const EM_POSFROMCHAR = &HD6&
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_CHARFROMPOS = &HD7&
Private Const EM_EXSETSEL = (WM_USER + 55)
Private Const EM_SETTARGETDEVICE = (WM_USER + 72)

'Drawing...
Private Const DT_CALCRECT = &H400
Private Const DT_RIGHT = &H2

'*** Types ***
Private Type CHARRANGE
  cpMin             As Long      ' First character of range (0 for start of doc)
  cpMax             As Long      ' Last character of range (-1 for end of doc)
End Type
Private Type TEXTRANGE
    chrg As CHARRANGE
    lpstrText As Long    ' /* allocated by caller, zero terminated by RichEdit */
End Type
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Type Undo_Item
    sAddText     As String
    sDelText     As String
    lStart       As Long
    lAddTextLen  As Long
    lDelTextLen  As Long
    bAddTextPlain As Boolean
    bDelTextPlain As Boolean
'    bAddTextRTF  As Boolean
'    bDelTextRTF  As Boolean
    ModifyType   As ModifyTypes
End Type
'*** API ***
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

'gdi
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

'user32
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function InvalidateRectAsNull Lib "user32" Alias "InvalidateRect" (ByVal hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Private Declare Function GetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long) As Long
Private Declare Function SetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As Any) As Long  ' lpPoint As POINTAPI) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

'*** Private Variables ***
Private bInHTMLExCode     As Boolean 'are we in server-side code?
Private lLastLine         As Long 'for detecting if we need to colourize a new area
Private cLangCombo(1)     As clsFlatCombo
Private bTempInHTMLExCode As Boolean 'value saved to this if we don't want to affect the global one
Private m_blnBusy         As Boolean 'are we busy?
Private bMainLocked       As Boolean 'have we already locked the main window?
Private bIgnoreEvent      As Boolean 'ignore events triggered by code
Private lLockCount        As Long 'count the # of times we have tried to lock main window
Private lBusyCount        As Long
Private m_lLastLine       As Long 'last line we were on
Private m_lLastLinePos    As Long 'the index of the last line (used for colourizing line by line)
Private Status            As ColourStatus 'current status pos (ie InComment)

Private lLastGoLine       As Long 'preserve the last line we jumped to (using go back/forward)
Private cLinesBack        As Collection 'line back history
Private cLinesForward     As Collection 'line forward history

Private lDragStart        As Long 'start pos of drag selection
Private lDragLen          As Long 'length of drag selection
Private bDragging         As Boolean 'are we doing a drag & drop?
Private bDraggingCopy     As Boolean
Private bIgnore           As Boolean 'ignore triggered events

Private m_lFlag()        As Long
Private m_lFlagCount      As Long

Private UndoStack         As Collection 'undo collection
Private RedoStack         As Collection 'redo collection
Private bRedoing          As Boolean 'are we redoing?
Private bUndoing          As Boolean 'are we undoing?
Private lPixelsPerLine    As Long
Private lMouseDownParam   As Long

Private m_rtfPos_tP       As POINTAPI 'saved scroll pos
Private m_rtfPos_tCR      As CHARRANGE 'saved selection

Private bResizingTriggered  As Boolean 'has colourizing been triggered by a resize?
Private bDeletingComment    As Boolean 'are we deleting a comment?
Private bDeletingEndComment As Boolean 'are we deleting an end comment?
Private lStartOfCommentCode As Long 'start of the comment

'*** Properties ***
Private m_cGlobal         As clsGlobalEditor
Private m_ViewMode        As ERECViewModes
Private m_bChanged        As Boolean 'has the document changed?
Private m_bLineChanged    As Boolean 'has the current line changed?
Private m_bCancelColour   As Boolean 'aborts colourizing
Private m_bSaved          As Boolean 'has the document been saved?
Private m_sFileName       As String 'the file we have saved to
Private m_bShowLines      As Boolean 'display the lines?
Private m_bNoEnterFocus   As Boolean 'don't set the focus to rtf on enterfocus
Private m_bLoadingFile    As Boolean
Private m_bRunTime        As Boolean
'*** Syntax Declarations ***
Private vSyntaxInfo         As Syntax_Info
Private sCurIndent          As String
Private m_lMaxCommentLen    As Long

'*** Event Declarations ***
Public Event SelChange()
Public Event FlagsChanged()
Public Event LoadFile(ByVal sFileName As String)
Public Event ShowPopup(x As Single, y As Single)
Public Event Error(ByVal Err As Long, ByVal Source As String, ByVal Error As String)

Private Const EM_GETTEXTLENGTHEX = (WM_USER + 95)
Private Type GETTEXTLENGTHEX
    flags As Long
    codepage As Integer
End Type
'*** Edit Commands ***
Public Sub Cut()
    'add to undo event list
    AddUndoDeleteEvent 0, True, CutText
    SendMessage rtfMain.hwnd, WM_CUT, 0&, 0&
End Sub
Public Sub Copy()
Dim sText As String
    If m_blnBusy Then
        'copy the previously selected text
        sText = TextInRange(m_rtfPos_tCR.cpMin, m_rtfPos_tCR.cpMax - m_rtfPos_tCR.cpMin)
        Clipboard.SetText sText
    Else
        'tell richedit to handle the copy
        SendMessage rtfMain.hwnd, WM_COPY, 0&, 0&
    End If
End Sub
Public Sub AssignGlobalEditor(GlobalEditor As clsGlobalEditor)
    Set m_cGlobal = GlobalEditor
    'call init
    pInitialise
End Sub
Public Sub Paste()
    'if it's not text, get out of here!
    If Clipboard.GetFormat(vbCFText) = False Then Exit Sub ' Not text, get out of here
    'save to textbox first...
    rtfTemp.Text = Clipboard.GetText
    'and then insert it...
    'this solves problem with copying from Visual InterDev,
    'which otherwise leaves us with characters we don't like!
    InsertCode rtfTemp.Text, True
End Sub
Public Sub Append()
    Dim sClipboard As String
    ' Get the clipboard data
    If Clipboard.GetFormat(vbCFText) Then
        'append the selected text to the clipboard
        'save current clipbard
        sClipboard = Clipboard.GetText(vbCFText)
        'erase clipboard
        Clipboard.Clear
        'set new text
        Clipboard.SetText sClipboard & rtfMain.SelText, vbCFText
    Else
        'can't append to an invalid format, such as picture!
        RaiseEvent Error(vbObjectError + 2222, "DevPadEditor.Editor:Append", "Invalid Clipboard Format")
    End If
End Sub
Public Sub SelectAll()
    'select all the text
    SetSelection 0, -1
End Sub
Public Sub Clear()
    'clear the selection
    rtfMain.Text = ""
End Sub


'*** RichTextBox Commands ***

Public Sub Refresh()
    InvalidateRectAsNull rtfMain.hwnd, 0&, 0
End Sub
'*** Load/Save ***
Public Function SaveFile(sFileName As String, Optional bIgnoreSave As Boolean = False) As Boolean
Dim lAttrib As Long
    On Error GoTo ErrHandler
    'abort if no filename
    Debug.Assert sFileName <> ""
    If sFileName = "" Then Exit Function
    'richtextbox can't save to a hidden file,
    'so check if file is hidden
    If Dir(sFileName) <> "" Then
        lAttrib = GetAttr(sFileName)
        If lAttrib = lAttrib And vbHidden Then
            'it is hidden... reset attribute
            SetAttr sFileName, vbNormal ' vbHidden
        Else
            'don't bother restoring... we haven't changed anything
            lAttrib = 0
        End If
    End If
    'save the file. Save in RTF if the current mode is that
    rtfMain.SaveFile sFileName, IIf(vSyntaxInfo.bRTF, rtfRTF, rtfText)
    If lAttrib <> 0 Then
        'restore attributes if we need to
        SetAttr sFileName, lAttrib
    End If
    If bIgnoreSave = False Then
        'if we don't want to ignore the save, then
        'set correct attributes
        m_bSaved = True
        m_bChanged = False
        m_sFileName = sFileName
    End If
    'successful
    SaveFile = True
    Exit Function
ErrHandler:
    'error!
    Err.Raise Err, "DevPadEditor.Editor:SaveFile", Error
End Function
Public Function LoadFile(ByVal sFileName As String, ByVal bRTF As Boolean, Optional ByVal bForceText As Boolean = False) As Boolean
    On Error GoTo ErrHandler
    m_bLoadingFile = True
    'load the file
    rtfMain.LoadFile sFileName, IIf(bRTF, rtfRTF, rtfText)
    
    If TextInRange(1, 4) = "ÿþi" Then
        'unicode fix...
        Debug.Print "Implementing Unicode Fix..."
        rtfMain.Text = Replace(Right$(rtfMain.Text, Len(rtfMain.Text) - 2), Chr$(0), "")
    End If
    
    If bForceText Then
        'ignore file extension, open as text
        ChangeMode "text.stx"
    Else
        'open the file using default syntax file
        SetMode sFileName
    End If
    
    'set attributes
    m_bSaved = True 'it has been loaded from a file
    m_bChanged = False
    m_sFileName = sFileName
    'redraw the lines
    DrawLines
    RaiseEvent LoadFile(sFileName)
    'suceeded!
    m_bLoadingFile = False
    LoadFile = True

    Exit Function
ErrHandler:
    'error!
    If Err = 75 Then
        'access error
        Err.Raise Err, "DevPadEditor.Editor:LoadFile", "The file, '" & sFileName & "', does not exist, or cannot be accessed." & vbCrLf & Error  ', "Colour.LoadFile", Error
    Else
        Err.Raise Err, "DevPadEditor.Editor:LoadFile", Error
    End If
End Function
'sets the syntax mode
Private Sub SetMode(sFileName As String)
Dim sExtension As String
Dim i As Long
Dim bFound As Boolean
    'we have been given the current file
    'select correct syntax file from its extension
    sExtension = GetExtension(sFileName)
    If sExtension <> "" Then
        With m_cGlobal
            For i = 1 To .SyntaxFilesCount
                If InStr(1, .SyntaxFile(i).sFilter, "." & sExtension, vbTextCompare) Then
                    ChangeMode .SyntaxFile(i).sFile
                    bFound = True
                    Exit For
                End If
            Next
        End With
    End If
    If bFound = False Then ChangeMode "text.stx"
End Sub

' Loads a template without prompts, or setting filename properties
Public Function LoadTemplate(sFileName As String)
    Dim vTemplate As TemplateInfo
    On Error Resume Next
    'retreive the info for the template
    vTemplate = m_cGlobal.GetTemplateInfo(sFileName)
    'load the template
    rtfMain.LoadFile m_cGlobal.TemplatePath & sFileName, rtfText
    'change mode to the correct one
    ChangeMode vTemplate.sSyntax
    'not modified
    Modified = False
    'set cursor sel
    rtfMain.SelStart = vTemplate.lSelStart
    rtfMain.SelLength = vTemplate.lSelLen
End Function
' Inserts Text
Private Sub InsertCode(sCode As String, bReplaceSelection As Boolean, Optional bIndent As Boolean = False)
    Dim cUndo           As Undo_Item
    Dim lTempStart      As Long
    Dim lTempEnd        As Long

    'init undo
    With cUndo
        'save cursor pos
        .lStart = rtfMain.SelStart
        'save the text we are deleting
        .sDelText = rtfMain.SelRTF
        'and its length
        .lDelTextLen = rtfMain.SelLength
        'set undo type
        .ModifyType = IIf(rtfMain.SelLength = 0, AddText, ReplaceText)
        'save the text we are adding
        .sAddText = sCode
        .bAddTextPlain = True
        'add to the undo
        UndoStack.Add cUndo
    End With
    
    If vSyntaxInfo.bCode = False Then
        'item is not code... we can skip lots!
        'clear selected text
        rtfMain.SelText = ""
        'set the text
        rtfMain.SelText = sCode
        m_bChanged = True
        Exit Sub
    End If
    
    'if we are already busy, don't bother colourizing
    If m_blnBusy = False Then
        'we are busy!
        Busy = True
        'save positions
        lTempStart = rtfMain.SelStart + 1
        lTempEnd = lTempStart + Len(sCode)
        ReColourRange sCode, lTempStart, lTempEnd
        Busy = False
    Else
        'set the text
        rtfMain.SelText = sCode
    End If
    
    'doc has changed
    rtfMain_Change
    'trigger sel change
    rtfMain_SelChange
End Sub
Private Sub ReColourRange(sCode As String, lStart As Long, lEnd As Long)
    Dim vSyntax         As ColourStatus
    'insert the text
    rtfMain.SelText = sCode
    If IsBiggerThanScreen(lEnd) Then
        'bigger than visible screen...
        're-colour the whole page
        ParseVisibleRange
    Else
        LockMain
        SaveCursorPos
        If vSyntaxInfo.bHTML Then
            'get status... we might be in html extension code
            vSyntax = GetStatus(lStart)
            'parse the html
            ParseHTMLRange lStart + 1, lEnd + 1, , (bTempInHTMLExCode)
        Else
            'parse code
            ParseRange lStart + 1, lEnd
        End If
        RestoreCursorPos
        UnlockMain
    End If
End Sub
'*** Control Properties ***
Public Property Let CancelColour(ByVal New_CancelColour As Boolean)
    'update value
    m_bCancelColour = New_CancelColour
    PropertyChanged "CancelColour"
End Property
Public Property Get Saved() As Boolean
    Saved = m_bSaved
End Property
Public Property Let Saved(ByVal bNew As Boolean)
    m_bSaved = bNew
    PropertyChanged "Saved"
End Property
Public Property Let LoadingFile(ByVal bNew As Boolean)
    'skips some code...
    m_bLoadingFile = bNew
End Property
Public Property Get Modified() As Boolean
    Modified = m_bChanged
End Property
Public Property Let Modified(ByVal bNew As Boolean)
    m_bChanged = bNew
    PropertyChanged "Modified"
End Property
'Public Property Get NoEnterFocus() As Boolean
'    NoEnterFocus = m_bNoEnterFocus
'End Property
'Public Property Let NoEnterFocus(bNew As Boolean)
'    m_bNoEnterFocus = bNew
'End Property

Public Property Get CanPaste() As Boolean
    'are we allowed to paste?
    CanPaste = SendMessage(rtfMain.hwnd, EM_CANPASTE, 0&, 0&)
End Property
Public Property Get CanCopy() As Boolean
    'is there a selection to copy?
    CanCopy = (rtfMain.SelLength > 0)
End Property

Public Property Get ShowLines() As Boolean
    ShowLines = m_bShowLines
End Property
Public Property Let ShowLines(ByVal bNew As Boolean)
    'hide/show the lines
    m_bShowLines = bNew
    picLines.Width = IIf(bNew, 735, 315)
    
    PropertyChanged "ShowLines"
    If rtfMain.Visible Then UserControl_Resize
End Property
Public Property Get Text() As String
    Text = rtfMain.Text
End Property
Public Property Let Text(ByVal New_Text As String)
Dim cUndo As Undo_Item
    'set the new richtextbox text
    'we need to do an undo event...
    With cUndo
        'set add text
        .bAddTextPlain = True
        .sAddText = New_Text
        'prevent updates
        If rtfMain.Text <> "" Then
            LockMain
            'we only want the RTF of the document,
            'not some extra tags added by TextRTF property
            SetSelection 0, -1
            .sDelText = rtfMain.SelRTF
            .lDelTextLen = rtfMain.SelLength
            UnlockMain
        End If
        .ModifyType = ReplaceText
        UndoStack.Add cUndo
    End With
    rtfMain.Text = New_Text
    're-colourize visible
    ParseVisibleRange
    PropertyChanged "Text"
End Property
Public Property Get SelLength() As Long
    SelLength = rtfMain.SelLength
End Property
Public Property Let SelLength(ByVal lNew As Long)
    rtfMain.SelLength = lNew
End Property
Public Property Get SelStart() As Long
    SelStart = rtfMain.SelStart
End Property
Public Property Let SelStart(ByVal lNew As Long)
    rtfMain.SelStart = lNew
End Property
Public Property Let SelText(ByVal sNew As String)
    'replace the selection
    InsertCode sNew, True
End Property
Public Property Get SelText() As String
    SelText = rtfMain.SelText
End Property
Public Property Get hwnd() As Long
    hwnd = rtfMain.hwnd
End Property
Public Property Get Font() As IFontDisp
    Set Font = rtfMain.Font
End Property
'only allow to be set privately...!
Private Property Set Font(New_Font As IFontDisp)
    'updates the font...
    Set rtfMain.Font = New_Font
    Set rtfTemp.Font = New_Font
    Set picLines.Font = New_Font
    lPixelsPerLine = picLines.TextHeight("90980&^,") / Screen.TwipsPerPixelX
    PropertyChanged "Font"
End Property
Public Sub SetFont(sFontName As String, nFontSize As Integer)
    Dim fFont As IFontDisp
    Set fFont = New StdFont
    fFont.Name = sFontName
    fFont.Size = nFontSize
    Set Font = fFont
End Sub
Public Property Let Redraw(bNew As Boolean)
    If bNew Then
        'start redrawing
        UnlockMain
        rtfMain_SelChange
    Else
        'no redraw
        LockMain
    End If
    'set busy flag too...
    Busy = Not bNew
End Property
Public Property Get FileName() As String
    'return the current filename
    FileName = m_sFileName
End Property
Public Property Let FileName(ByVal New_FileName As String)
    m_sFileName = New_FileName
    PropertyChanged "FileName"
End Property
'*** Status Properties ***
Public Property Get Mode() As String
    Mode = cboLanguage.Text
End Property
Public Property Let Mode(ByVal New_Mode As String)
    If IsNumeric(New_Mode) Then Exit Property
    cboLanguage.Text = New_Mode
End Property
Public Property Get LineCount() As Long
    LineCount = SendMessage(rtfMain.hwnd, EM_GETLINECOUNT, 0&, 0&)
End Property
Public Property Get GetFirstLineVisible() As Long
    GetFirstLineVisible = SendMessage(rtfMain.hwnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
End Property
Public Property Get CurrentColumn() As Long
    Dim lCurLine As Long
    ' Current Line
    lCurLine = 1 + rtfMain.GetLineFromChar(rtfMain.SelStart)
    ' Column
    CurrentColumn = SendMessage(rtfMain.hwnd, EM_LINEINDEX, ByVal lCurLine - 1, 0&)
    CurrentColumn = (rtfMain.SelStart) - CurrentColumn
End Property
Public Property Get CurrentLine() As Long
    CurrentLine = 1 + rtfMain.GetLineFromChar(rtfMain.SelStart)
End Property
Public Property Get Border() As Boolean
    Border = BorderStyle
End Property
Public Property Let Border(ByVal bState As Boolean)
    BorderStyle() = Abs(bState)
    PropertyChanged "Border"
End Property

Public Property Let ViewMode(ByVal eViewMode As ERECViewModes)
Dim lParam As Long
Dim wParam As Long
    'sets the View Mode
    Select Case eViewMode
    Case ercWYSIWYG
        wParam = Printer.hdc
        lParam = Printer.Width
    Case ercWordWrap
        'wParam = 0
        'lParam = 0
    Case ercDefault
        'wParam = 0
        lParam = 1
    End Select
    SendMessageLong rtfMain.hwnd, EM_SETTARGETDEVICE, wParam, lParam
    SendMessageLong rtfTemp.hwnd, EM_SETTARGETDEVICE, wParam, lParam
    m_ViewMode = eViewMode
End Property
Public Property Get ViewMode() As ERECViewModes
    ViewMode = m_ViewMode
End Property

'*** Cursor Position ***
Private Sub SaveCursorPos()
Dim f As GETTEXTLENGTHEX
Const GTL_PRECISE = 2
    'To get the parsing range (character range),
    'we use some messages that are available in
    'RichEd 3.0 dll to tell us where the scroll
    'currently is...By getting the current scroll
    'point, we can set things back to how they
    'were before we modified the text...
    'if you don't have RichEd 3.0 we are in trouble!
    
    f.flags = GTL_PRECISE
    Debug.Print "StartVal:" & SendMessage(rtfMain.hwnd, EM_GETTEXTLENGTHEX, VarPtr(f), 0)
    
    'save the sel
    On Error Resume Next
    SendMessage rtfMain.hwnd, EM_EXGETSEL, 0, m_rtfPos_tCR
    'Get beginning point...
    SendMessage rtfMain.hwnd, EM_GETSCROLLPOS, 0, m_rtfPos_tP
    
End Sub
Private Sub RestoreCursorPos()
Dim f As GETTEXTLENGTHEX
Const GTL_PRECISE = 2
    'restore the selection
    SendMessage rtfMain.hwnd, EM_EXSETSEL, 0, m_rtfPos_tCR
    'send the focus to the richedit, unless told not to
    'restore the scroll position
    SendMessage rtfMain.hwnd, EM_SETSCROLLPOS, 0, m_rtfPos_tP
    f.flags = GTL_PRECISE
    Debug.Print "EndVal:" & SendMessage(rtfMain.hwnd, EM_GETTEXTLENGTHEX, VarPtr(f), 0)
End Sub
'*** Redrawing ***
Private Sub LockMain()
    If bMainLocked Then
        'we have already locked the richedit...
        'increment lock count
        lLockCount = lLockCount + 1 'why not lLockCount++ ?!
        Exit Sub
    End If
    'lock the text box to prevent changes
    SendMessage rtfMain.hwnd, WM_SETREDRAW, False, 0&
    
    rtfMain.Locked = True
    bMainLocked = True
    'set lock count
    lLockCount = 1
End Sub
Private Sub UnlockMain()
    'richedit is not locked... abort
    If Not bMainLocked Then Exit Sub
    If lLockCount > 1 Then
        'Lock Main was called more than once...
        'Don't unlock yet
        lLockCount = lLockCount - 1
        Exit Sub
    End If
    'reset lock count
    lLockCount = 0
    'unlock richedit
    rtfMain.Locked = False
    'tell the richedit that it can redraw
    SendMessage rtfMain.hwnd, WM_SETREDRAW, True, 0&
    'force it to redraw
    InvalidateRectAsNull rtfMain.hwnd, 0, 0 '1
    'richedit is not locked
    bMainLocked = False
    Call rtfMain_SelChange
    'DoEvents
End Sub
Public Property Let Busy(bNew As Boolean)
    If m_blnBusy = True And bNew Then
        'add another to count
        lBusyCount = lBusyCount + 1
    ElseIf m_blnBusy = True And bNew = False Then
        If lBusyCount > 1 Then
            'called more than once...
            lBusyCount = lBusyCount - 1
        Else
            'reset flag
            m_blnBusy = False
            lBusyCount = 0
        End If
    ElseIf m_blnBusy = False And bNew = True Then
        'set to busy
        m_blnBusy = True
        lBusyCount = 1
    End If
End Property
Private Function GetLinePos(lStart As Long, lLen As Long)
    'returns the start and length of the current line
    Dim lCurrLine       As Long
    Dim lStartPos       As Long
    Dim lEndLen         As Long
    Dim lSelLen         As Long
    Dim sSearch         As String
    Dim sText           As String
    Dim lCommentStart   As Long
    Dim lCommentEnd     As Long

    'Get current line
    lCurrLine = SendMessage(rtfMain.hwnd, EM_LINEFROMCHAR, rtfMain.SelStart, 0&)
    If vSyntaxInfo.bMultiComment = False Then
        'not a multicomment file... easy!
        'Set the start pos at the beginning of the line
        lStart = SendMessage(rtfMain.hwnd, EM_LINEINDEX, lCurrLine, 0&)
        'Get the length of the line
        lLen = SendMessage(rtfMain.hwnd, EM_LINELENGTH, rtfMain.SelStart, 0&)
    Else
        'not so easy...
        sText = rtfMain.Text
        'get the last occurance of a comment
        lCommentStart = InStrRev(Left$(sText, rtfMain.SelStart), vSyntaxInfo.sMultiCommentStart)
        If lCommentStart <> 0 Then
            'get its end tag
            lCommentEnd = InStr(lCommentStart, sText, vSyntaxInfo.sMultiCommentEnd)
        Else
            lCommentEnd = -1
        End If
        Debug.Assert lCommentEnd <> 0
        If lCommentEnd = 0 Then
            'this shouldn't happen ... this proc shouldn't have been called!
            
            'we are in a comment
            'set the start pos at the beginning of the comment
'            lStart = lCommentStart - 1
'            If lCommentEnd = lCommentStart + 1 Then
'                lLen = Len(sText) - rtfMain.SelStart
'            Else
'                ' Get the length of the line
'                lLen = lCommentEnd + Len(vSyntaxInfo.sMultiCommentEnd) - 1 - rtfMain.SelStart
'            End If
        Else
            ' Set the start pos at the beginning of the line
            lStart = SendMessage(rtfMain.hwnd, EM_LINEINDEX, lCurrLine, 0&)
            If lStart < lCommentEnd Then
                lStart = lCommentEnd + 1
            End If
            ' Get the length of the line
            lLen = SendMessage(rtfMain.hwnd, EM_LINELENGTH, rtfMain.SelStart, 0&)
        End If
    End If
End Function

Private Sub ColourCurLine()
    Dim lEnd As Long
    Dim lStart As Long
    If vSyntaxInfo.bHTMLExtension And bInHTMLExCode = False Then Exit Sub
    'get the line range
    GetLinePos lStart, lEnd
    'values returned by GetLinePos are compatible with richtextbox, not Mid
    ParseRange lStart + 1, lStart + 1 + lEnd, (vSyntaxInfo.bHTMLExtension And bInHTMLExCode)
End Sub
Private Function pStripIgnoreChars(ByVal sString As String) As String
Dim lPos As Long
    
'    'find sString in line...
'    lResult = GetFirstCharFromTxt(, lTxtLen + 1, sString)
'    If lResult >= 0 Then 'is before any normal characters
'        'select the found text
'        rtfTemp.SelStart = lStart + lResult - 1
'        rtfTemp.SelLength = Len(sString)
'        'remove it
'        rtfTemp.SelText = ""
'        'adjust the text length
'        lTxtLen = lTxtLen - Len(sString)
'    End If
'
End Function
Private Sub RemoveBlock(ByVal sString As String)
    'removes a string from the selected block
    On Error GoTo ErrHandler
    Dim lLineCount      As Long
    Dim lResult         As Long
    Dim lStart          As Long
    Dim lEnd            As Long
    Dim lLine           As Long
    Dim lTxtLen         As Long
    Dim sTemp           As String
    Dim cUndo           As Undo_Item
    Dim bHTMLComment    As Boolean
    Dim cCursor         As New clsCursor
    Dim lPos            As Long
    'get the current status
    bHTMLComment = (vSyntaxInfo.bHTML And sString = "") Or (bInHTMLExCode = False And vSyntaxInfo.bHTMLExtension And sString = vSyntaxInfo.sSingleComment)
    'this is a HTML comment... we can't undo these yet
    If bHTMLComment Then Exit Sub
    'is there a selection?
    lPos = -1
    If rtfMain.SelLength = 0 Then lPos = rtfMain.SelStart
    'we are busy
    Busy = True
    cCursor.SetCursor vbHourglass
    'prevent redrawing
    LockMain
    'select the whole block (ie no partial lines)
    SelectBlock
    'save selected text to variable
    sTemp = rtfMain.SelRTF
    rtfTemp.TextRTF = sTemp

    ' undo stuff
    'set the delete text
    cUndo.sDelText = sTemp
    cUndo.lDelTextLen = rtfMain.SelLength
    cUndo.lStart = rtfMain.SelStart
    'get the number of lines
    lLineCount = SendMessage(rtfTemp.hwnd, EM_GETLINECOUNT, 0&, 0&)
    'and its length
    lTxtLen = TempCharacterCount

    For lLine = 0 To lLineCount - 1
        'loop through each line...
        'get the line's start point
        lStart = SendMessage(rtfTemp.hwnd, EM_LINEINDEX, lLine, 0&)
        'find sString in line...
        lResult = GetFirstCharFromTxt(lStart + 1, lTxtLen + 1, sString)
        If lResult >= 0 Then 'is before any normal characters
            'select the found text
            rtfTemp.SelStart = lStart + lResult - 1
            rtfTemp.SelLength = Len(sString)
            'remove it
            rtfTemp.SelText = ""
            'adjust the text length
            lTxtLen = lTxtLen - Len(sString)
            lPos = lPos - Len(sString)
        End If
    Next
    'select everything
    'SetTempSelection 0, -1 doesn't work!
    SetTempSelection 0, TempCharacterCount
    sTemp = rtfTemp.SelRTF
    'save the cursor pos
    lStart = rtfMain.SelStart
    'do the undo stuff...
    cUndo.sAddText = sTemp
    cUndo.lAddTextLen = rtfTemp.SelLength
    If sString = vSyntaxInfo.sSingleComment Then
        cUndo.ModifyType = UnCommentText
    Else
        cUndo.ModifyType = OutdentText
    End If
    UndoStack.Add cUndo
    'update rtfmain with new text
    rtfMain.SelRTF = sTemp
    'calculate the end position
    lEnd = rtfMain.SelStart - lStart
    're-select the text
    rtfMain.SelStart = lStart
    rtfMain.SelLength = lEnd
    If ((vSyntaxInfo.bHTML = False) Or (vSyntaxInfo.bHTMLExtension = True And bInHTMLExCode = True)) And (sString <> sCurIndent) Then
        SaveCursorPos
        'If lStart = 0 Then lStart = 1
        ParseRange lStart + 1, lStart + lEnd + 1, , True
        RestoreCursorPos
    End If
    'clear temporary textbox
    rtfTemp.Text = ""
    Busy = False
    If lPos > -1 Then
        rtfMain.SelStart = lPos
        rtfMain.SelLength = 0
    End If
    'allow redrawing
    UnlockMain
    m_bChanged = True
    Exit Sub
ErrHandler:
    RaiseEvent Error(Err, "DevPadEditor.Editor:RemoveBlock", Error)
End Sub
'uncomment & comment code blocks
Public Sub CommentBlock()
    GetStatus
    If vSyntaxInfo.bCode Then InsertBlock vSyntaxInfo.sSingleComment
End Sub
Public Sub UncommentBlock()
    GetStatus
    If vSyntaxInfo.bCode Then RemoveBlock vSyntaxInfo.sSingleComment
End Sub
'indent & outdent code blocks
Public Sub Indent()
    GetStatus
    InsertBlock sCurIndent
End Sub
Public Sub Outdent()
    GetStatus
    RemoveBlock sCurIndent
End Sub
Private Sub SelectBlock()
    Dim lStart As Long
    Dim lLen As Long
    Dim lNewLen As Long
    Dim lNewStart As Long
    Dim lTempEnd As Long
    With rtfMain
        ' ensure we have selected the whole line
        lLen = .SelLength
        lStart = .SelStart
        'if there is a vbCrLf at the end... don't include it
        If TextInRange(rtfMain.SelLength + rtfMain.SelStart - 1, 2) = vbCrLf Then lLen = lLen - 1
        ' go to beginning of line
        lNewStart = SendMessage(.hwnd, EM_LINEINDEX, .GetLineFromChar(lStart), 0&)
        'extend selection as needed
        'get the beginning of the last line...
        lTempEnd = SendMessage(.hwnd, EM_LINEINDEX, .GetLineFromChar(lStart + lLen), 0&) '+ SendMessage(.hWnd, EM_LINELENGTH, .GetLineFromChar(.SelStart + lEnd), 0&)  '(lStart - .SelStart) + lEnd
        'add line length
        lTempEnd = lTempEnd + SendMessage(.hwnd, EM_LINELENGTH, lStart + lLen, 0&)
        'we now have the new end...
        
        'adjust length for spanning to beginning of line
        lNewLen = lLen + (lStart - lNewStart)
        'adjust length for spanning to end of line
        lNewLen = lNewLen + lTempEnd - (lNewStart + lNewLen)
        'goto end of line

        .SelStart = lNewStart
        If lNewLen >= 0 Then .SelLength = lNewLen

        DoEvents
    End With
End Sub
Public Function InsertBlock(sString As String)
    On Error GoTo ErrHandler
    Dim lLineCount      As Long
    Dim lEnd            As Long
    Dim lStart          As Long
    Dim lPos            As Long
    Dim i               As Long
    Dim sTemp           As String
    Dim cUndo           As Undo_Item
    Dim bNoSel          As Boolean
    Dim bHTMLComment    As Boolean
    Dim cCursor         As New clsCursor
    
    'set the cursor
    cCursor.SetCursor vbHourglass
    'no redrawing...
    LockMain
    Busy = True
    'is there a selection?
    If rtfMain.SelLength = 0 Then
        lPos = rtfMain.SelStart + Len(sString)
        bNoSel = True
    End If
    bHTMLComment = (vSyntaxInfo.bHTML And sString = "") Or (bInHTMLExCode = False And vSyntaxInfo.bHTMLExtension And sString = vSyntaxInfo.sSingleComment)
    
    'select to end of lines if not a HTML comment
    If bHTMLComment = False Then SelectBlock
    'save variable
    sTemp = rtfMain.SelRTF
    
    rtfTemp.TextRTF = sTemp
    ' undo stuff
    cUndo.sDelText = sTemp
    cUndo.lDelTextLen = rtfMain.SelLength
    cUndo.lStart = rtfMain.SelStart
    With rtfTemp
        If bHTMLComment Then
            'is a html comment...simple!
            'add <!-- at beginning
            .SelStart = 0
            .SelText = "<!--"
            'add --> at end
            .SelStart = TempCharacterCount
            .SelText = "-->"
        Else
            'loop through all the lines
            lLineCount = SendMessage(.hwnd, EM_GETLINECOUNT, 0&, 0&)
            For i = 0 To lLineCount - 1
                .SelStart = SendMessage(.hwnd, EM_LINEINDEX, i, 0&)
                .SelText = sString
            Next
        End If
        'select everything...
        SetRichTextSelection rtfTemp.hwnd, 0, TempCharacterCount
        'save the sel start
        lStart = rtfMain.SelStart
        'update the text
        rtfMain.SelRTF = .SelRTF
    End With
    With rtfMain
        'adjust end position
        lEnd = rtfMain.SelStart - lStart
        'restore the selection
        rtfMain.SelStart = lStart
        rtfMain.SelLength = lEnd
        ' comment
        If bHTMLComment Then
            'colour as HTML comment
            rtfMain.SelColor = vSyntaxInfo.vClr_HTMLComment
        ElseIf sString = vSyntaxInfo.sSingleComment Then
            'colour as comment
            rtfMain.SelColor = vSyntaxInfo.vClr_Comment
        End If
        'save selected rtf
        cUndo.sAddText = rtfMain.SelRTF
        cUndo.lAddTextLen = rtfMain.SelLength
        'set undo type
        If sString = vSyntaxInfo.sSingleComment Then
            cUndo.ModifyType = CommentText
        Else
            cUndo.ModifyType = IndentText
        End If
        UndoStack.Add cUndo
    End With
    'clear temp rtfbox
    rtfTemp.Text = ""
    'richedit has changed
    m_bChanged = True
    Busy = False
    'redraw again
    UnlockMain
    
    'trigger sel_change event if there was a selection
    If bNoSel = False Then
        rtfMain_SelChange
    Else
        rtfMain.SelStart = lPos
        rtfMain.SelLength = 0
    End If
    Exit Function
ErrHandler:
    RaiseEvent Error(Err, "DevPadEditor.Editor:InsertBlock", Error)
End Function
Private Function GetFirstCharFromTxt(ByVal lStart As Long, ByVal lEnd As Long, ByVal sString As String) As Long
    Dim lLen        As Long
    Dim i           As Long
    Dim sResult     As String
    Dim sTemp       As String
    sTemp = TextInRange(lStart, lEnd - lStart, rtfTemp.hwnd)
    'pass on to other proc
    GetFirstCharFromTxt = GetFirstCharFromString(sString, sTemp)
End Function

Private Function GetFirstCharFromString(ByVal sFindString As String, ByVal sText As String) As Long
    Dim lLen        As Long
    Dim i           As Long
    Dim sResult     As String
    lLen = Len(sFindString)
    If lLen = 0 Then lLen = 1
    'loop through each char
    For i = 1 To Len(sText)
        sResult = Mid$(sText, i, lLen)
        Select Case sResult
        Case sFindString
            'that character has been found... return its position
            GetFirstCharFromString = i
            Exit Function
        Case " ", vbTab
            'ignore these chars
        Case Else
            'another character has been found
            GetFirstCharFromString = -1
            Exit Function
        End Select
    Next
    'nothing of interest has been found
    GetFirstCharFromString = -2
End Function

Public Sub SetFocus()
    'sets the focus
    rtfMain.SetFocus
End Sub
Public Property Get IndentString() As String
    'return the current indent string
    IndentString = sCurIndent
End Property
Public Function GetIndent(sLine As String) As String
    'get the indent of text
    Dim i As Long
    For i = 1 To Len(sLine)
        Select Case Mid$(sLine, i, 1)
        Case " ", Chr$(vbKeyTab)
            'add to list...
            GetIndent = GetIndent & Mid$(sLine, i, 1)
        Case Else
            'another char... exit
            Exit For
        End Select
    Next
End Function
Public Property Get LineText(Optional lLine As Long = -1) As String
    ' Get current line?
    If lLine = -1 Then lLine = CurrentLine
    LineText = GetLineTextByLine(lLine)
End Property
Private Function GetLineTextByLine(lLine As Long) As String
    Dim lLen As Long
    Dim lLineStart As Long
    ' Set the start pos at the beginning of the line
    lLineStart = SendMessage(rtfMain.hwnd, EM_LINEINDEX, lLine - 1, 0&) + 1
    ' Get the length of the line
    lLen = SendMessage(rtfMain.hwnd, EM_LINELENGTH, lLineStart, 0&)
    ' Select the line
    GetLineTextByLine = TextInRange(lLineStart, lLen, rtfMain.hwnd)
End Function
Private Function GetLineTextByIndex(lLineStart As Long) As String
    Dim lLen As Long
    ' Get the length of the line
    lLen = SendMessage(rtfMain.hwnd, EM_LINELENGTH, lLineStart, 0&)
    ' Select the line
    GetLineTextByIndex = TextInRange(lLineStart, lLen, rtfMain.hwnd)
End Function
Private Function GetStatus(Optional lStart As Long = -1) As ColourStatus
    Dim bStatusOnly             As Boolean
    Dim CurrentStatus           As ColourStatus
    Dim sText                   As String
    Dim lLastOpenCommentTag     As Long
    Dim lLastCloseCommentTag    As Long
    Dim lLastOpenTag            As Long
    Dim lLastCloseTag           As Long
    Dim lLastOpenScriptTag      As Long
    Dim lLastCloseScriptTag     As Long
    Dim sQuote                  As String
    Dim bInQuote                As Boolean
    Dim i                       As Long
    Dim lResult                 As Long
    Dim lLen                    As Long
    Dim sFind                   As String
    'this code is terrible!
    'if anyone has a better way, please email
    'james@vbweb.co.uk... thanks!
    
    'reset flag
    bTempInHTMLExCode = False
    If lStart = -1 Then
        'reset flag
        bInHTMLExCode = False
        'get start pos
        lStart = rtfMain.SelStart
    Else
        'we want to return the status, not
        'set the global variables
        bStatusOnly = True
    End If
    If vSyntaxInfo.bHTML Then
        'is a HTML file
        sText = TextInRange(1, lStart)
        'save text length
        lLen = CharacterCount
        
        If vSyntaxInfo.bHTMLExtension Then
            'we are a HTML extension (ie ASP)
            'get last script tags
            lLastOpenScriptTag = InStrRev(sText, vSyntaxInfo.sHTMLExtensionStart)
            lLastCloseScriptTag = InStrRev(sText, vSyntaxInfo.sHTMLExtensionEnd)
            ' check for script first
            If lLastOpenScriptTag > lLastCloseScriptTag Then
                'we are in a HTML extension tag
                'now check for a comment...
                lLastOpenCommentTag = InStrRev(sText, "<!--")
                lLastCloseCommentTag = InStrRev(sText, "-->")
                If lLastOpenCommentTag > lLastCloseCommentTag Then
                    'we are in a comment
                    CurrentStatus = InComment
                    'set current text colour to correct value
                    If rtfMain.SelLength = 0 And bStatusOnly = False Then rtfMain.SelColor = vSyntaxInfo.vClr_HTMLComment
                    GoTo TheEnd
                End If
                'we are in HTML Ex code
                bTempInHTMLExCode = True
                'set current text colour accordingly
                If rtfMain.SelLength = 0 And bStatusOnly = False Then rtfMain.SelColor = vSyntaxInfo.vClr_Text
                GoTo CheckMultiComment
            End If
        End If
        
        ' get positions of last HTML tags
        lLastOpenTag = InStrRev(sText, "<")
        lLastCloseTag = InStrRev(sText, ">")

        If lLastOpenTag = 0 Then
            'we have no open tag... therefore we are not in any tag
            CurrentStatus = OutTag
            If rtfMain.SelLength = 0 And bStatusOnly = False Then rtfMain.SelColor = vSyntaxInfo.vClr_Text
            GoTo TheEnd
        End If
        
        'get script tag positions
        'we have to check for script before <!-- --> comments, because
        'they are used within script tags too!
        lLastOpenScriptTag = InStrRev(sText, "<script", , vbTextCompare)
        lLastCloseScriptTag = InStrRev(sText, "</script>", , vbTextCompare)
        'are we in a script tag?
        If lLastOpenScriptTag > lLastCloseScriptTag Then
            CurrentStatus = InScript
            If rtfMain.SelLength = 0 And bStatusOnly = False Then rtfMain.SelColor = vSyntaxInfo.vClr_HTMLScript
            GoTo TheEnd
        End If
        
        If lLastOpenCommentTag = 0 Then
            'check for a comment
            '(we may have already checked in server-side script check)
            lLastOpenCommentTag = InStrRev(sText, "<!--")
            lLastCloseCommentTag = InStrRev(sText, "-->")
            
            'If lLastCloseCommentTag = 0 Then lLastCloseCommentTag = lLen
            If lLastOpenCommentTag > lLastCloseCommentTag Then
                'we are in a comment
                CurrentStatus = InComment
                'set colour....
                If rtfMain.SelLength = 0 And bStatusOnly = False Then rtfMain.SelColor = vSyntaxInfo.vClr_HTMLComment
                GoTo TheEnd
            End If
        End If
        
        'finally, check if we are in a normal tag
        If lLastOpenTag >= lLastCloseTag Then
            'we are in a tag
            CurrentStatus = InTag
            If rtfMain.SelLength = 0 And bStatusOnly = False Then
                'we need to set the cursor colour... so here we go
                'now we need to see if we are in a quote
                '(this is a pain!)
                'get the text of the tag
                sQuote = TextInRange(lLastOpenTag, lStart + 1 - lLastOpenTag)
                'assume not...
                bInQuote = False
                For i = 0 To 1
                    'we need to check for " and '
                    sFind = IIf(i = 0, """", "'")
                    Do
                        'find next position of "
                        lResult = InStr(lResult + 1, sQuote, """")
                        If lResult = 0 Then Exit Do
                        bInQuote = Not bInQuote
                    Loop
                    If bInQuote Then Exit For 'that's ok...
                    'otherwise, check for other quote too
                Next
                If bInQuote Then
                    'set current colour accordingly
                    rtfMain.SelColor = vSyntaxInfo.vClr_Text 'vbBlack
                Else
                    rtfMain.SelColor = vSyntaxInfo.vClr_HTMLTag '&HBF0000
                End If
            End If
        Else
            'we are not in a tag
            CurrentStatus = OutTag
            If rtfMain.SelLength = 0 And bStatusOnly = False Then rtfMain.SelColor = vSyntaxInfo.vClr_Text 'vbBlack
        End If
    ElseIf vSyntaxInfo.bMultiComment Then
CheckMultiComment:
        'checks for multi-comments
        If sText = "" Then sText = TextInRange(1, lStart + 1)
        lLastOpenCommentTag = InStrRev(sText, vSyntaxInfo.sMultiCommentStart)
        lLastCloseCommentTag = InStrRev(sText, vSyntaxInfo.sMultiCommentEnd)
        ' in comment?
        If lLastOpenCommentTag > lLastCloseCommentTag Then
            CurrentStatus = InComment
            If rtfMain.SelLength = 0 And bStatusOnly = False And rtfMain.SelColor <> vSyntaxInfo.vClr_Comment Then rtfMain.SelColor = vSyntaxInfo.vClr_Comment
        Else
            'not in a comment!
            CurrentStatus = OutComment
            If rtfMain.SelLength = 0 And bStatusOnly = False And rtfMain.SelColor <> vSyntaxInfo.vClr_Text Then rtfMain.SelColor = vSyntaxInfo.vClr_Text
        End If
    End If
TheEnd:
    If bStatusOnly = False Then
        bInHTMLExCode = bTempInHTMLExCode
        If bInHTMLExCode = False And vSyntaxInfo.bHTMLExtension Then
            'only used for html extensions...
            sCurIndent = vSyntaxInfo.sHTMLIndent
        Else
            'default to standard indent
            sCurIndent = vSyntaxInfo.sIndent
        End If
        Status = CurrentStatus
    Else
        'just return the status... don't set global variables
        GetStatus = CurrentStatus
    End If
End Function

Private Sub cboLanguage_Click()
    'ignore click event...
    If bIgnoreEvent Then Exit Sub
    'save the syntax type
    SaveSetting REG_KEY, "Settings", "DefaultSyntaxType", cboLanguage.ListIndex
    'update the mode
    ChangeMode ""
End Sub
Public Sub ChangeMode(sFile As String)
    'if no file specified... use item selected in cboLanguage
    If sFile = "" Then sFile = m_cGlobal.SyntaxFile(cboLanguage.Text).sFile
    'use specified file
    LoadSyntaxFile sFile

    If vSyntaxInfo.bCode = False And vSyntaxInfo.bRTF = False Then
        'we are now plain-text... remove colourizing
        LockMain
        SaveCursorPos
        ResetColour 0, -1 'all
        RestoreCursorPos
        UnlockMain
    Else
        'update visible range
        ParseVisibleRange
    End If
End Sub

'*** Procedures Selection ***
Private Sub cboProcedures_Click()
    If bIgnore Or cboProcedures.ListIndex = -1 Then Exit Sub
    'if we have a valid item, and it is not dropped, then...
    If SendMessage(cboProcedures.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 Then
        'goto the position
        rtfMain.SelStart = cboProcedures.ItemData(cboProcedures.ListIndex)
        'set the focus
        If rtfMain.Visible Then rtfMain.SetFocus
    End If
End Sub

Private Sub cboProcedures_DropDown()
    Dim i           As Long
    Dim lPos        As Long
    Dim lRes(1)     As Long
    Dim sProcs()    As String
    Dim sText       As String
    Dim lIndex      As Long
    Dim lLine       As Long
    Dim lLen        As Long
    Dim sLastText   As String
    Dim cCursor     As New clsCursor
    Dim lPos2       As Long
    On Error GoTo ErrHandler
    'split the procedures into an array
    sProcs = Split(vSyntaxInfo.sProcedures, "*")
    'set cursor
    cCursor.SetCursor vbHourglass
    'no redraw
    SendMessage cboProcedures.hwnd, WM_SETREDRAW, False, 0
    'save the text
    sLastText = cboProcedures.Text
    'clear combo
    cboProcedures.Clear
    cboProcedures.AddItem "(Header Section)"

    For i = 0 To UBound(sProcs) - 1
        If sProcs(i) <> "" Then
            Do
                'find next item
                lPos = rtfMain.Find(sProcs(i) & " ", lPos + 1, , rtfNoHighlight)
                If lPos <> -1 Then
                    'item found...
                    'ignore vbCrLf
                    lPos = lPos + 2
                    'get the line
                    lLine = SendMessage(rtfMain.hwnd, EM_LINEFROMCHAR, lPos, 0&)
                    'get its index
                    lIndex = SendMessage(rtfMain.hwnd, EM_LINEINDEX, lLine, 0&)
                    'get the lines text
                    sText = GetLineTextByIndex(lIndex + 1)
                    lLen = Len(sText)
                    lRes(0) = InStr(1, sText, "(")
                    'keyword inside ( ... ignore
                    If (lPos - lIndex) + 2 < lRes(0) Then
                        'remove text upto procedure keyword...
                        sText = Right$(sText, Len(sText) - (lPos - lIndex) + 2)
                        lRes(0) = InStr(1, sText, "(")
                        lRes(1) = InStr(1, sText, ")")
                        'if (lPos - lIndex) + 2
                        If lRes(0) < lRes(1) Or (lRes(1) = 0 And lRes(0) = 0) Or (lRes(1) = 0) And Right$(sText, 1) <> ";" Then
                            '( has to occur before ), and we mustn't have more than one
                            'bracket
                            If InStr(lRes(0) + 1, sText, "(") = 0 And InStr(lRes(1) + 1, sText, ")") = 0 Then
                                lPos = lPos + lLen
                                If lRes(0) <> 0 Then
                                    If lRes(1) <> 0 Then
                                        'remove text after )
                                        sText = Left$(sText, lRes(1))
                                    End If
                                    'remove text between ()
                                    sText = Left$(sText, lRes(0) - 1) & Right$(sText, Len(sText) - lRes(1))
                                    'remove text other than proc name
                                    lRes(1) = InStrRev(Left$(sText, lRes(0)), " ")
                                    If lRes(1) = lRes(0) - 1 Then lRes(1) = InStrRev(Left$(sText, lRes(0) - 2), " ")
                                    If lRes(1) <> 0 Then sText = Right$(sText, Len(sText) - lRes(1))
                                End If
                                If Left$(sText, 1) <> """" Then
                                    'add the item
                                    cboProcedures.AddItem Trim$(sText)
                                    'set its data
                                    cboProcedures.ItemData(cboProcedures.NewIndex) = lIndex
                                End If
                            End If
                        End If
                    End If
                End If
            Loop While lPos <> -1
        End If
    Next i
    'allow redraw
    SendMessage cboProcedures.hwnd, WM_SETREDRAW, True, 0
    ''add item if there are none
    'If cboProcedures.ListCount = 0 Then cboProcedures.AddItem "<No Sections>"
    'redraw combo
    cboProcedures.Refresh
    If sLastText <> "" Then
        'ignore the click event
        bIgnore = True
        On Error Resume Next
        'restore text if possible
        cboProcedures.Text = sLastText
        'restore flag
        bIgnore = False
    End If
    Exit Sub
ErrHandler:
    RaiseEvent Error(Err, "DevPadEditor.Editor:ListProcedures", Error)
End Sub
Private Sub cboProcedures_KeyUp(KeyCode As Integer, Shift As Integer)
    'do default action
    If KeyCode = vbKeyReturn Then cboProcedures_Click
End Sub
Private Sub imgPic_Click(Index As Integer)
    'make the combo drop if the icon has been clicked
    If Index = 1 Then
        SendMessage cboLanguage.hwnd, CB_SHOWDROPDOWN, True, 0
    Else
        SendMessage cboProcedures.hwnd, CB_SHOWDROPDOWN, True, 0
    End If
End Sub

Private Sub picLines_DblClick()
    Dim lPos As Long
    Dim lLine As Long
    Dim tp As POINTAPI
    GetCursorPos tp
    ScreenToClient picLines.hwnd, tp
    lPos = CharFromPos(1, tp.y)
    lLine = LineFromChar(lPos)
    'Debug.Print lLine
    AddFlag lLine + 1
End Sub

Private Sub picLines_Resize()
    UserControl_Resize
End Sub

'*** RTF Events ***

Private Sub rtfMain_Change()
    'ignore?
    If m_blnBusy Then Exit Sub
    'rtf has changed
    m_bChanged = True
    m_bLineChanged = True
    If bRedoing Or bUndoing Then
        'if we are undoing, or redoing, prevent further redos...
        bRedoing = False
        bUndoing = False
        ClearStack RedoStack
        'raise selchange event
        RaiseEvent SelChange
    End If
End Sub

Private Sub rtfMain_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrHandler

    Dim sIndent             As String
    Dim sLine               As String
    Dim iKeyCode            As Integer
    Dim lLineStart          As Long
    Dim bNoReset            As Boolean
    Dim sSurroundingText    As String
    Dim lCommentOpen        As Long
    Dim lCommentClose       As Long
    Dim lStartSearch        As Long
    Dim sNewText            As String
    Dim lCurPos             As Long

    'exit this procedure if we are doing something
    If m_blnBusy Then Exit Sub
    'don't bother processing control keys
    If KeyCode = vbKeyControl Or KeyCode = vbKeyShift Or KeyCode = vbKeyMenu Then Exit Sub
   
    'save the original keycode, in case we change it
    iKeyCode = KeyCode
    
    If KeyCode = vbKeyF5 Then
        'hotkey to update colourizing..
        '(just in case it gets itself in a mess!)
        ParseVisibleRange
        Exit Sub
    End If
    
    If vSyntaxInfo.bHTML Then
        'we are a HTML file... we always need to know the current status
        'get pos status
        GetStatus
    End If

    If vSyntaxInfo.bHTML And bInHTMLExCode = False Then
        'we are a HTML file, and not in HTML server-side code
        If Shift = vbShiftMask Then
            'check for key events... if we are pressing < or a "
            'we want to update the current colour before they are
            'processed...
            
            'check for <
            If Status = OutTag And iKeyCode = 188 Then
                rtfMain.SelColor = vSyntaxInfo.vClr_HTMLTag
            'check for "
            ElseIf Status = InTag And iKeyCode = 50 Then
                rtfMain.SelColor = vSyntaxInfo.vClr_Text
            End If
        End If
    ElseIf vSyntaxInfo.bMultiComment Then
        'clear flags
        bDeletingComment = False
        bDeletingEndComment = False
        lStartOfCommentCode = 0

        Select Case KeyCode
        Case vbKeyBack, vbKeyDelete
            ' detect if we are deleting a comment
            ' deleting text from the right
            
            ' get the surrounding Text on the right
            If rtfMain.SelStart > Len(vSyntaxInfo.sMultiCommentStart) Then
                lStartSearch = rtfMain.SelStart - (Len(vSyntaxInfo.sMultiCommentStart)) + IIf(KeyCode = vbKeyDelete, 2, 1)
                sSurroundingText = TextInRange(lStartSearch, (Len(vSyntaxInfo.sMultiCommentStart) * 2 - 1), rtfMain.hwnd)
            Else
                lStartSearch = 1
                sSurroundingText = TextInRange(lStartSearch, (Len(vSyntaxInfo.sMultiCommentStart)) + rtfMain.SelStart, rtfMain.hwnd)
            End If
            If Len(sSurroundingText) > 1 Then
                'get the open/close positions
                lCommentOpen = InStr(1, sSurroundingText, vSyntaxInfo.sMultiCommentStart)
                lCommentClose = InStr(1, sSurroundingText, vSyntaxInfo.sMultiCommentEnd)
        
                If (lCommentClose = 0 Or lCommentClose < lCommentOpen) And lCommentOpen <> 0 Then
                    ' only an open tag
                    bDeletingComment = True
                    lStartOfCommentCode = lCommentOpen + lStartSearch - 2
                ElseIf lCommentClose <> 0 Then
                    ' a close tag
                    bDeletingEndComment = True
                    bDeletingComment = True
                    lStartOfCommentCode = lCommentOpen + lStartSearch
                Else
                    'no comment...
                    DoEvents
                    'check if we are creating a comment by deleting a character
                    'calculate the cursor pos in sSurroundingText
                    lCurPos = rtfMain.SelStart - lStartSearch + IIf(KeyCode = vbKeyDelete, 1, 0)
                    If lCurPos > 0 Then
                        'get the text that will be created after deletion
                        sNewText = Left$(sSurroundingText, lCurPos)
                        sNewText = sNewText & Right$(sSurroundingText, Len(sSurroundingText) - lCurPos - 1)
                        sNewText = Trim$(sNewText)
                        If InStr(1, sNewText, vSyntaxInfo.sMultiCommentEnd) Or InStr(1, sNewText, vSyntaxInfo.sMultiCommentStart) Then
                            'is a valid comment...
                            bDeletingComment = True
                            'is an end comment?
                            bDeletingEndComment = CBool(InStr(1, sNewText, vSyntaxInfo.sMultiCommentStart))
                            'set start pos
                            lStartOfCommentCode = lStartSearch + IIf(sNewText = vSyntaxInfo.sMultiCommentEnd, 1, -1)
                        End If
                    End If
                End If
            End If
        Case Else
            sSurroundingText = TextInRange(rtfMain.SelStart, 2, rtfMain.hwnd)
            ' Detect if we are seperating the two comment characters
            If sSurroundingText = vSyntaxInfo.sMultiCommentStart Or sSurroundingText = vSyntaxInfo.sMultiCommentEnd Then
                bDeletingComment = True
                lStartOfCommentCode = rtfMain.SelStart - 1
                'are we deleting the end part of a comment?
                bDeletingEndComment = (sSurroundingText = vSyntaxInfo.sMultiCommentEnd)
            End If
        End Select
    Else
    End If
    
    
    Select Case Shift
    Case 0
        Select Case iKeyCode
        Case vbKeyReturn
            ' AutoIndent Code
            If vSyntaxInfo.bAutoIndent = True And m_blnBusy = False Then
                Busy = True
                ' get the current text
                sLine = LineText()
                If sLine <> "" Then
                    If vSyntaxInfo.sAutoIndentChar <> "" Then
                        lCurPos = GetFirstCharFromString(vSyntaxInfo.sAutoOutdentChar, sLine)
                        If lCurPos = -1 Then
                            'only an indent
                            Busy = False
                            bNoReset = True
                        End If
                        If lCurPos > 0 And m_bLineChanged Then
                            '} is present...
                            'outdent the current line
                            Outdent
                            'get line text again...
                            sLine = LineText()
                        End If
                    End If
                    'get the indent text
                    sIndent = GetIndent(sLine)
                    'get the start of the line
                    lLineStart = SendMessage(rtfMain.hwnd, EM_LINEINDEX, CurrentLine - 1, 0)
                    If Len(sIndent) > rtfMain.SelStart - lLineStart Then
                        sIndent = Left$(sIndent, rtfMain.SelStart - lLineStart)
                    End If
                    'current line ends in {, and indent hasn't been changed by
                    'a } being present... increase the indent
                    If Right$(sLine, Len(vSyntaxInfo.sAutoIndentChar)) = vSyntaxInfo.sAutoIndentChar And vSyntaxInfo.sAutoIndentChar <> "" And m_bLineChanged Then sIndent = sIndent & sCurIndent
                    'add undo event
                    AddUndoEvent rtfMain.SelStart, vbCrLf & sIndent
                    'add the indent
                    rtfMain.SelText = vbCrLf & sIndent
                    'cancel the return key
                    KeyCode = 0
                End If
                If bNoReset = False Then Busy = False
            End If
            pAdjustFlagPos CurrentLine, 1
        Case vbKeyBack
            If vSyntaxInfo.bDelIndent And Len(sCurIndent) > 1 Then
                Dim sBehind As String
                'If rtfMain.SelStart - vSyntaxInfo.sIndent + 1 <= 0 Then Exit Sub
                'sText = GetLineTextByLine(rtfMain, lStart)
                'If rtfMain.SelStart + 1 - lStart - vSyntaxInfo.sIndent + 1 <= 0 Then Exit Sub
                sBehind = TextInRange(rtfMain.SelStart + 1 - Len(sCurIndent), Len(sCurIndent))
                If sBehind = sCurIndent Then
                    Busy = True
                    LockMain
                    'add a delete event
                    AddUndoDeleteEvent -(Len(sCurIndent)), False, DeleteText
                    
                    rtfMain.SelStart = rtfMain.SelStart - Len(sCurIndent)
                    rtfMain.SelLength = Len(sCurIndent)
                    rtfMain.SelText = ""

                    UnlockMain
                    KeyCode = 0
                    Busy = False
                    DoEvents
                End If
            End If
        Case vbKeyTab
            ' Indent code
            If rtfMain.SelLength > 0 And vSyntaxInfo.bTabIndent Then
                Indent
                KeyCode = 0
            Else
                'add undo event
                AddUndoEvent rtfMain.SelStart, sCurIndent
                'sim tab using specified character
                rtfMain.SelText = sCurIndent
                KeyCode = 0
            End If
        
        End Select
    Case vbShiftMask
        If iKeyCode = vbKeyTab Then
            ' outdent code
            If rtfMain.SelLength > 0 And vSyntaxInfo.bTabIndent Then
                Outdent
                KeyCode = 0
            End If
        End If
    End Select
    If KeyCode = vbKeyHome And Shift = 0 Then
        'save current pos
        lStartSearch = rtfMain.SelStart
        lLineStart = rtfMain.SelLength
        'get the beginning of the line (ignoring the indent)
        lCurPos = LineIndex(CurrentLine) + Len(GetIndent(LineText))
        If rtfMain.SelStart = lCurPos Then
            'already there, go to beginning of line
            rtfMain.SelStart = LineIndex(CurrentLine)
        Else
            rtfMain.SelStart = lCurPos
        End If
        'keep selection...
        'If Shift = vbShiftMask Then SetSelection rtfMain.SelStart, rtfMain.SelStart + 200 'rtfMain.SelLength = (lStartSearch - rtfMain.SelStart) + lLineStart
        
        KeyCode = 0
    End If
    ' undo/redo Code
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        'add a delete event
        AddUndoDeleteEvent IIf(KeyCode = vbKeyBack, -1, 0), False, DeleteText    'lAmount
        DoEvents
    End If
    'Simulate selection change, as we have cancelled the key
    If KeyCode = 0 Then rtfMain_SelChange
    
    If KeyCode = vbKeyDelete And m_blnBusy = False Then
        ' Ugly hack for Delete Key, so that the control
        ' will have acted on the key before we call the
        ' keypress event
        KeyCode = 0
        'ignore...!
        Busy = True
        SendMessage rtfMain.hwnd, WM_KEYDOWN, vbKeyDelete, 0
        Busy = False
        'something has still changed...
        Call rtfMain_Change
        Call rtfMain_SelChange
        ' keypress event does not occur for the Del key, so trigger it!
        Call rtfMain_KeyPress(0)
    End If
    Exit Sub
ErrHandler:
    RaiseEvent Error(Err, "DevPadEditor.Editor:KeyDown", Error)
End Sub
Private Sub rtfMain_KeyPress(KeyAscii As Integer)
    'abort if busy
    If m_blnBusy Then Exit Sub
    Dim nOldAscii       As Integer
    Dim lNextOpenTag    As Long
    Dim lLastCloseTag   As Long
    Dim lNextCloseTag   As Long
    Dim lLastOpenTag    As Long
    
    ' Undo/Redo Code
    If KeyAscii >= 32 Or KeyAscii = vbKeyReturn Then
        'add an undo event
        AddUndoEvent rtfMain.SelStart, IIf(KeyAscii = 13, vbCrLf, Chr(KeyAscii))
    End If

    If vSyntaxInfo.bHTML And bInHTMLExCode = False Then
        'do nothing...!
    ElseIf vSyntaxInfo.bMultiComment Then
        'get the current status
        GetStatus
        'detect comment creation as needed
        If Status = InComment Then
            DetectCommentCreation KeyAscii, False
        Else
            DetectCommentCreation KeyAscii, True
        End If
    End If
    'save event
    nOldAscii = KeyAscii
    If IsMoveKey(KeyAscii) = False And KeyAscii >= 32 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then
        ' for the rest of the code, we need the
        ' new character to be there!
        rtfMain.SelText = Chr$(KeyAscii)
        ' cancel key
        KeyAscii = 0
    End If
    'now process creation of comments
    
    If vSyntaxInfo.bHTML And bInHTMLExCode = False Then
    ElseIf vSyntaxInfo.bMultiComment Then
        If bDeletingComment Then
            'set comment start to beginning
            If lStartOfCommentCode = -1 Then lStartOfCommentCode = 0
            'find the next close tag
            lNextCloseTag = FindFirst(lStartOfCommentCode + 1, vSyntaxInfo.sMultiCommentEnd) + 1
            If bDeletingEndComment = False Then
                'find the next open tag
                lNextOpenTag = FindFirst(lStartOfCommentCode + 1, vSyntaxInfo.sMultiCommentStart) - 1
                'if the open tag is before the close tag, and it is found (ie != -1)
                If lNextOpenTag < lNextCloseTag And lNextOpenTag <> -1 Then
                    lNextCloseTag = lNextOpenTag
                End If
            End If
            'close tag not found...set it to end of doc
            If lNextCloseTag = 1 Then lNextCloseTag = CharacterCount + 1
            'get ready to rumble!
            Busy = True
            SaveCursorPos
            LockMain
            ' we are deleting a comment
            If bDeletingEndComment Then
                ' we are deleting the end of a comment
                ' therefore, colour the rest of the text
                ' upto the next */ or the end of the document
                ' green
                rtfMain.SelStart = lStartOfCommentCode
                rtfMain.SelLength = lNextCloseTag - rtfMain.SelStart
                rtfMain.SelColor = vSyntaxInfo.vClr_Comment
            Else
                GetStatus
                If Status = OutComment Then
                    ' colourize from the start of the ex-comment
                    ' to the end of the comment
                    If IsBiggerThanScreen(lNextCloseTag) Then
                        ParseVisibleRange , True
                    Else
                        ParseRange lStartOfCommentCode, lNextCloseTag + 2, True
                    End If
                End If
            End If
            'restore changes...
            RestoreCursorPos
            UnlockMain
            Busy = False
        End If
    Else
    End If
End Sub
Private Sub DetectCommentCreation(KeyAscii As Integer, bStart As Boolean)
    Dim lAmount  As Long
    Dim sChar(1) As String
    Dim sNewChar As String
    Dim lAltChar As Long
    Dim i        As Long
    sNewChar = Chr$(KeyAscii)
    'get chars
    If bStart Then
        sChar(0) = Left$(vSyntaxInfo.sMultiCommentStart, 1) '"/"
        sChar(1) = Right$(vSyntaxInfo.sMultiCommentStart, 1) '"*"
    Else
        sChar(0) = Left$(vSyntaxInfo.sMultiCommentEnd, 1)
        sChar(1) = Right$(vSyntaxInfo.sMultiCommentEnd, 1)
    End If
    For i = 0 To 1
        If sChar(i) = sNewChar Then
            If i = 1 Then 'end of end/start comment
                lAmount = 0
                lAltChar = i - 1
            Else
                lAmount = 1
                lAltChar = i + 1
            End If
            If TextInRange(rtfMain.SelStart + lAmount, 1, rtfMain.hwnd) = sChar(lAltChar) Then
                ' char nextdoor is a /
    
                ' trick my code into thinking that
                ' we are deleting the start of a comment
                ' and therefore colours from lStartOfCommentCode to the next end comment
                ' or end of the text
                bDeletingComment = True
                bDeletingEndComment = bStart 'False
                If bStart Then
                    lStartOfCommentCode = rtfMain.SelStart - 1
                Else
                    lStartOfCommentCode = rtfMain.SelStart + 2 'IIf(i < 2, 1, 2)
                End If
            End If
            Exit For
        End If
    Next
'    Select Case sNewChar
'    Case sChar(1), sChar(3)
'        lAmount = IIf(bStart, 0, 1)
'        If TextInRange(rtfMain.SelStart + lAmount, 1, rtfMain.hWnd) = sChar(0) Then
'            ' char nextdoor is a /
'
'            ' trick my code into thinking that
'            ' we are deleting the start of a comment
'            ' and therefore colours from lStartOfCommentCode to the next end comment
'            ' or end of the text
'            bDeletingComment = True
'            bDeletingEndComment = bStart 'False
'            If bStart Then
'                lStartOfCommentCode = rtfMain.SelStart - 1
'            Else
'                lStartOfCommentCode = rtfMain.SelStart + 2
'            End If
'            'GetStatus
'        End If
'    Case sChar(0), sChar(2)
'        lAmount = Abs(bStart)
'        If TextInRange(rtfMain.SelStart + lAmount, 1, rtfMain.hWnd) = sChar(1) Then
'            ' char to the left is a /
'
'            ' trick my code into thinking that
'            ' we are deleting the start of a comment
'            ' and therefore colours from lStartOfCommentCode to the next end comment
'            ' or end of the text
'            bDeletingComment = True
'            bDeletingEndComment = bStart
'            If bStart Then
'                lStartOfCommentCode = rtfMain.SelStart - 1
'            Else
'                lStartOfCommentCode = rtfMain.SelStart + 1 ' + 3
'            End If
'        End If
'    End Select
End Sub

Private Sub rtfMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Debug.Print "mousedown"
On Error GoTo ErrHandler
    Dim lPos        As Long
    'display the popup menu?
    If Button = vbRightButton Then
        'yes
        RaiseEvent ShowPopup(x, y)
    End If
    Exit Sub
ErrHandler:
    RaiseEvent Error(Err, "DevPadEditor.Editor:MouseDown", Error)
End Sub

Private Sub rtfMain_SelChange()
    On Error GoTo ErrHandler

    Dim bTemp   As Boolean
    Dim lLine   As Long
    'abort if busy or we are about to insert
    If m_blnBusy Or tmrInsert.Enabled = True Then Exit Sub
    'we are not drop & dragging...
    bDragging = False
    bResizingTriggered = False
    'update the lines...
    DrawLines
    'raise selchange event
    RaiseEvent SelChange
    'get the current line...
    lLine = SendMessage(rtfMain.hwnd, EM_LINEFROMCHAR, rtfMain.SelStart + rtfMain.SelLength, 0&)
    If lLine <> lLastGoLine Then
        'if it has changed, add it to lines list
        cLinesBack.Add lLine
        lLastGoLine = lLine
    End If
    'line not changed...
    If m_lLastLine = lLine Then Exit Sub
    'are we dealing with code? (probably!)
    If vSyntaxInfo.bCode Then
        'colour any new lines if pos has changed
        ParseVisibleRange True
        'want to colour by line, and the line has changed
        If vSyntaxInfo.bColourByLine Then
            'abort if we are a server-side script, but not in server-side code
            If vSyntaxInfo.bHTMLExtension And bInHTMLExCode = False Then Exit Sub
            If m_bLineChanged = False Then
                'line has not changed...
                m_lLastLine = lLine
                m_lLastLinePos = rtfMain.SelStart + rtfMain.SelLength
                Exit Sub
            End If
            
            If LastLineEmpty(m_lLastLine) = False Then
                'last line did not just contain spaces...
                'we are busy
                Busy = True
                LockMain
                'save cursor
                SaveCursorPos
                ' go back to the last line we were on
                rtfMain.SelStart = m_lLastLinePos
                rtfMain.SelLength = 0
                'get the current status
                GetStatus
                If Status <> InComment Or vSyntaxInfo.bMultiComment = False Then
                    'we are not in a comment
                    'colour it
                    ColourCurLine
                End If
                'restore the cursor
                RestoreCursorPos
                UnlockMain
                'we are not busy
                Busy = False
                'the line has not changed
                m_bLineChanged = False
            End If
            m_lLastLine = SendMessage(rtfMain.hwnd, EM_LINEFROMCHAR, rtfMain.SelStart + rtfMain.SelLength, 0&)
            m_lLastLinePos = rtfMain.SelStart + rtfMain.SelLength
        End If
    End If
    Exit Sub
ErrHandler:
    RaiseEvent Error(Err, "DevPadEditor.Editor:SelChange", Error)
End Sub
'*** Jump Forward/Back ***
Public Sub Back()
    SetPos True
End Sub
Public Sub Forward()
    SetPos False
End Sub
Public Property Get CanGoBack() As Boolean
    CanGoBack = (cLinesBack.Count <> 0)
End Property
Public Property Get CanGoForward() As Boolean
    CanGoForward = (cLinesForward.Count <> 0)
End Property

Public Sub SetPos(bLinesBack As Boolean)
    Static bRunning     As Boolean
    Dim lPos            As Long
    Dim cCollection     As Collection
    Dim cAltCollection  As Collection
    Dim lItem           As Long
    'if we are running, then don't allow another
    If bRunning Then Exit Sub
    bRunning = True
    'set collections correctly...
    If bLinesBack Then
        'we are going backwards
        Set cCollection = cLinesBack
        Set cAltCollection = cLinesForward
    Else
        'we are going forwards
        Set cCollection = cLinesForward
        Set cAltCollection = cLinesBack
    End If
    'only bother if there are items in the collection
    If cCollection.Count <> 0 Then
        'we are busy
        Busy = True
        'save collection count
        lItem = cCollection.Count
        'if the last item matches where we are now...
        If cCollection.Item(lItem) = SendMessage(rtfMain.hwnd, EM_LINEFROMCHAR, rtfMain.SelStart + rtfMain.SelLength, 0&) Then
            'tranfer across collection
            cAltCollection.Add cCollection.Item(lItem)
            cCollection.Remove (lItem)
            'and update count
            lItem = cCollection.Count
        End If
        If lItem <> 0 Then
            'get the last 'go' line
            lLastGoLine = cCollection.Item(lItem)
            'get its index
            lPos = SendMessage(rtfMain.hwnd, EM_LINEINDEX, lLastGoLine, 0&)
            'restore the cursor position
            If lPos <> -1 Then rtfMain.SelStart = lPos
            'transfer across to other collection
            cAltCollection.Add lLastGoLine
            cCollection.Remove (lItem)
        End If
        'we are not busy
        Busy = False
    End If
    'trigger selchange event
    rtfMain_SelChange
    'we are not running...
    bRunning = False
End Sub

Private Function IsMoveKey(KeyCode As Integer) As Boolean
    Select Case KeyCode
    Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyPageUp, vbKeyPageDown, vbKeyHome, vbKeyEnd ', vbKeyShift, vbKeyControl
        IsMoveKey = True
    End Select
End Function
Private Function LastLineEmpty(lLine As Long) As Boolean
    'returns if the last line didn't have any 'real' text (excludes spaces & tabs)
    If TextInRange(SendMessage(rtfMain.hwnd, EM_LINEINDEX, lLine, 0&) + 1, 2) = vbCrLf Then
        LastLineEmpty = True
    ElseIf GetFirstCharFromString("", GetLineTextByLine(lLine + 1)) <> -1 Then
        LastLineEmpty = True
    End If
End Function

Private Sub tmrColor_Timer()
    'this timer triggers for colourizing when scrolling or resizing...
    'abort if we are not enabled, or we are loading a file
    If tmrColor.Enabled = False Or m_bLoadingFile = True Then Exit Sub
    'reset flag
    m_bCancelColour = False
    'disable the timer
    tmrColor.Enabled = False
    'request visible area to be colourized
    
    ParseVisibleRange Not bResizingTriggered
    'reset flag
    bResizingTriggered = False
End Sub
Private Sub tmrInsert_Timer()
    'inserts code after dropping & dragging...
    'no good in rtfMain_OLEDrop() event, as it causes a GPF
Dim sInsertText     As String
Dim cUndo           As Undo_Item
Dim lStart          As Long
Dim lLen            As Long
Dim bNoChange       As Boolean
    'disable the timer
    tmrInsert.Enabled = False
    'do the usual
    LockMain
    Busy = True
    'save the selstart
    lStart = rtfMain.SelStart
    'restore the cursor pos to drag pos
    rtfMain.SelStart = lDragStart
    rtfMain.SelLength = lDragLen
    'get its selected text
    sInsertText = rtfMain.SelText
 '   Debug.Print sInsertText
    
    If bDraggingCopy = False Then
        'we are moving the text
        If (lStart > lDragStart) And lDragStart = lStart - Len(sInsertText) Then
            'no change!
            bNoChange = True
        Else
            'build up undo struct...
            With cUndo
                .lStart = rtfMain.SelStart
                .sDelText = rtfMain.SelRTF
                .lDelTextLen = rtfMain.SelLength
                'half a vbcrlf... remove it
                 If Len(rtfMain.SelText) <> rtfMain.SelLength Then
                     rtfMain.SelLength = rtfMain.SelLength - 2
                 End If
                .ModifyType = DeleteText
                UndoStack.Add cUndo
            End With
            'erase text
            rtfMain.SelText = ""
            If lStart > rtfMain.SelStart Then
                'adjust drop position according to the text
                'we have just deleted
                lStart = lStart - Len(sInsertText)
                If lStart < 0 Then lStart = 0
            End If
            'restore cursor
            rtfMain.SelStart = lStart
        End If
    Else
        'restore cursor
        rtfMain.SelStart = lStart
    End If
    'not busy...
    Busy = False
    'insert the text in the new location
    If bNoChange = False Then InsertCode sInsertText, True
    'unlock
    UnlockMain
End Sub

'*** Subclassing ***
Private Property Let ISubclass_MsgResponse(ByVal RHS As vbwSubClass.EMsgResponse)
End Property

Private Property Get ISubclass_MsgResponse() As vbwSubClass.EMsgResponse
    Select Case CurrentMessage
    Case WM_VSCROLL, WM_MOUSEWHEEL
        ISubclass_MsgResponse = emrConsume
    Case Else
        ISubclass_MsgResponse = emrPreprocess
    End Select
End Property
Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim iScrollPos As Long
Dim bScrolling As Boolean
Dim tp As POINTAPI
    ' scroll events
    Select Case iMsg
    Case WM_MOUSEWHEEL
        'go up or down 3 lines...
        SendMessage hwnd, EM_GETSCROLLPOS, 0, tp
        If wParam > 0 Then
            tp.y = tp.y - (lPixelsPerLine * 5)
        Else
            tp.y = tp.y + (lPixelsPerLine * 5)
        End If
        SendMessage hwnd, EM_SETSCROLLPOS, 0, tp
        DrawLines
        'reset timer
        pResetTimer
    Case WM_VSCROLL
        'are we scrolling, or scrolled?
        bScrolling = (LoWord(wParam) = 4 Or LoWord(wParam) = 5)
        If bScrolling Then
            'get info using wparam
            iScrollPos = HiWord(wParam)
        Else
            'call api
            iScrollPos = GetScrollPos(hwnd, 1)
        End If
        
        'update lines
        If Int(iScrollPos / lPixelsPerLine) <> (iScrollPos / lPixelsPerLine) Then
            'not divisible by 16
            'we are showing half a line...
            If bScrolling Then
                wParam = MakeDWord(LoWord(wParam), Int(iScrollPos / lPixelsPerLine) * lPixelsPerLine)
            Else
                'hmm...
            End If
        End If
        'send message to rtf for processing
        ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
        'update the lines
        DrawLines
        'reset flags
        bResizingTriggered = False
        If wParam = 8 Then '8=SB_END... scroll end
            tmrColor.Enabled = False
            'update colourizing
            ParseVisibleRange True
        ElseIf wParam > 1 Then 'not up/down button
            'just reset timer
            'pResetTimer
        End If
    Case WM_LBUTTONDOWN
        Debug.Print lParam
    End Select
End Function
Private Sub pAttachMessages()
    On Error Resume Next
    AttachMessage Me, rtfMain.hwnd, WM_VSCROLL
    AttachMessage Me, rtfMain.hwnd, WM_MOUSEWHEEL
    AttachMessage Me, rtfMain.hwnd, WM_LBUTTONDOWN
End Sub
Private Sub pDetachMessages()
    On Error Resume Next
    DetachMessage Me, rtfMain.hwnd, WM_VSCROLL
    DetachMessage Me, rtfMain.hwnd, WM_MOUSEWHEEL
    DetachMessage Me, rtfMain.hwnd, WM_LBUTTONDOWN
End Sub

Private Sub UserControl_Click()
    Debug.Print "click"
End Sub

'*** User Control ***
Private Sub UserControl_EnterFocus()
    On Error Resume Next
    UserControl_Resize
End Sub

Private Sub pInitialise()
    Dim lStyle      As Long
    Dim lwidth      As Long
    Dim sItems()    As String
    Dim i           As Long

    'init stuff
    If m_bRunTime Then
        'init variables, so we know this is the first time
        '(compared to them being at 0)
        lLastLine = -1
        lLastGoLine = -1
        'init collections
        Set UndoStack = New Collection
        Set RedoStack = New Collection
        Set cLinesBack = New Collection
        Set cLinesForward = New Collection
        'set the timer interval
        tmrColor.Interval = GetSetting(REG_KEY, "Settings", "ScrollDelay", 500)
        'start subclassing
        pAttachMessages
        'attach flat appearances
        Set cLangCombo(0) = New clsFlatCombo
        Set cLangCombo(1) = New clsFlatCombo
        cLangCombo(0).Attach cboLanguage.hwnd
        cLangCombo(1).Attach cboProcedures.hwnd
        'set the drop height
        SetDropHeight cboLanguage, 200
        'list the available languages
        m_cGlobal.GetSyntaxFiles sItems()
        For i = 1 To UBound(sItems)
            cboLanguage.AddItem sItems(i)
            'save max width
            If TextWidth(sItems(i)) > lwidth Then lwidth = TextWidth(sItems(i))
        Next
        'adjust drop width to max needed
        SendMessage cboLanguage.hwnd, CB_SETDROPPEDWIDTH, (lwidth / Screen.TwipsPerPixelX) + 30, 0&
        'set combo width
        cboLanguage.Width = lwidth + 400
        'adjust drop height
        SetDropHeight cboProcedures, 200
        'widen procedures drop
        SendMessage cboProcedures.hwnd, CB_SETDROPPEDWIDTH, 250, 0&
        'move to correct position
        imgPic(0).Left = cboLanguage.Left + cboLanguage.Width + 45
        cboProcedures.Left = cboLanguage.Left + cboLanguage.Width + 345
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Debug.Print "down"
End Sub

Private Sub UserControl_Show()
    'if we are loading a file, or are not at runtime, abort
    If m_bLoadingFile Or m_bRunTime = False Then Exit Sub
    'restore last language?
    If cboLanguage.ListIndex = -1 Then
        On Error Resume Next
        cboLanguage.ListIndex = GetSetting(REG_KEY, "Settings", "DefaultSyntaxType", 0)
        If Err Then cboLanguage.ListIndex = 0
    End If
    'update lines
    DrawLines True
End Sub
Private Sub UserControl_Terminate()
    'remove subclassing
    pDetachMessages
End Sub
Private Sub UserControl_Resize()
    On Error Resume Next
    'If m_bShowLines Then
        'resize richtextbox with lines
        rtfMain.Move picLines.Width, picTB.Height + 15, ScaleWidth - picLines.Width, ScaleHeight - picTB.Height - 15
        picLines.Move 0, picTB.Height, picLines.Width, ScaleHeight
        'move 3d line accordingly to line visiblity
       ' LineSep.X1 = 20 * Screen.TwipsPerPixelX  'picLines.Width - 15
   ' Else
        'resize richtextbox without lines
    '    rtfMain.Move 0, picTB.Height + 15, ScaleWidth, ScaleHeight - picTB.Height - 15
       ' LineSep.X1 = 0
    'End If
    'update line numbering
    DrawLines
    'reset lastline flag (for ParseVisibleRange)
    'lLastLine = -1
    'resize 3d line...
    LineSep.X2 = picTB.ScaleWidth

    If rtfMain.Visible = True And m_bLoadingFile = False Then
        'is triggered by resize...
        bResizingTriggered = True
        'reset timer
        pResetTimer
    End If
End Sub

Private Sub UserControl_InitProperties()
    'init default font
    Set Font = Ambient.Font
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'restore properties
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    ShowLines = PropBag.ReadProperty("ShowLines", True)
    'are we at runtime?
    m_bRunTime = UserControl.Ambient.UserMode

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'save properties
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("ShowLines", m_bShowLines, True)
End Sub
'returns the line from its char number
Public Property Get LineFromChar(lChar As Long) As Long
    LineFromChar = rtfMain.GetLineFromChar(lChar)
End Property
'returns the char from xy pos
Public Property Get CharFromPos(ByVal xPixels As Long, ByVal yPixels As Long) As Long
Dim tp As POINTAPI
   tp.x = xPixels
   tp.y = yPixels
   CharFromPos = SendMessage(rtfMain.hwnd, EM_CHARFROMPOS, 0, tp)
End Property
'sets the cursor selection...
Public Sub SetSelection(ByVal lStart As Long, ByVal lEnd As Long)
    SetRichTextSelection rtfMain.hwnd, lStart, lEnd
End Sub
Private Sub SetTempSelection(ByVal lStart As Long, ByVal lEnd As Long)
    SetRichTextSelection rtfTemp.hwnd, lStart, lEnd
End Sub
Private Sub SetRichTextSelection(ByVal hwnd As Long, ByVal lStart As Long, ByVal lEnd As Long)
    Dim tCR As CHARRANGE
    tCR.cpMin = lStart
    tCR.cpMax = lEnd
    SendMessage hwnd, EM_EXSETSEL, 0, tCR
End Sub
'returns the character count...
Public Function CharacterCount() As Long
    CharacterCount = SendMessageLong(rtfMain.hwnd, WM_GETTEXTLENGTH, 0, 0)
End Function
Private Function TempCharacterCount() As Long
    TempCharacterCount = SendMessageLong(rtfTemp.hwnd, WM_GETTEXTLENGTH, 0, 0)
End Function
Private Sub rtfMain_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' check the format of the data that is being dropped
    If Data.GetFormat(vbCFFiles) = True Then
        Effect = vbDropEffectNone
        'frmMainForm.DropFiles Data
    ElseIf Data.GetFormat(vbCFText) Or Data.GetFormat(vbCFRTF) Then
        Effect = vbDropEffectNone
        tmrInsert.Enabled = True
    End If
End Sub

Private Sub rtfMain_OLEDragOver(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    ' check the data to see if it is what we will allow. if not so "no drop"
    If Data.GetFormat(vbCFFiles) Then
        Effect = vbDropEffectCopy And Effect
    ElseIf Data.GetFormat(vbCFText) And Button = vbLeftButton Then
        If Shift = vbCtrlMask Then
            Effect = vbDropEffectCopy
        Else
            Effect = vbDropEffectMove
        End If
        bDraggingCopy = (Shift = vbCtrlMask)
        'save the selection
        If bDragging = False Then
            lDragStart = rtfMain.SelStart
            lDragLen = rtfMain.SelLength
            bDragging = True
        End If
        Busy = True
        rtfMain.SelStart = CharFromPos(ScaleX(x, vbTwips, vbPixels), ScaleY(y, vbTwips, vbPixels) - 1)
        Busy = False
    Else
        Effect = vbDropEffectNone
    End If
End Sub
'*** Undo/Redo ***
Public Sub Redo()
On Error GoTo ErrHandler
    Dim cRedo As Undo_Item
    'abort if busy
    If m_blnBusy Then Exit Sub
    'abort if no items in redo stack
    If RedoStack.Count = 0 Then Exit Sub
    'reset flags
    bRedoing = False
    bUndoing = False
    'get redo item
    cRedo = RedoStack(RedoStack.Count)
    'move to undo stack
    UndoStack.Add cRedo
    'remove from redo stack
    RedoStack.Remove (RedoStack.Count)
    'prevent redrawing
    LockMain
    Busy = True
    'select the text
    rtfMain.SelStart = cRedo.lStart
    'select the text that was deleted
    'undo and created it, and we are going to
    'delete it again!
    If cRedo.bDelTextPlain Then
        rtfMain.SelLength = Len(cRedo.sDelText)
    Else
        rtfMain.SelLength = cRedo.lDelTextLen
    End If
    'delete any text that was deleted
    rtfMain.SelText = ""
    'redo the text edit
    If cRedo.bAddTextPlain Then
        If cRedo.sAddText <> "" Then
            'recolour the current range with the text to add
            ReColourRange cRedo.sAddText, cRedo.lStart, cRedo.lStart + Len(cRedo.sAddText)
        End If
    Else
        'replace with the text that was added
        rtfMain.SelRTF = cRedo.sAddText
        'we need to update if we are uncommenting text
        If cRedo.ModifyType = UnCommentText Then ParseVisibleRange
    End If
    'restore flags
    Busy = False
    UnlockMain
    'trigger change event
    rtfMain_Change
    'we are redoing
    bRedoing = True
    Exit Sub
ErrHandler:
    RaiseEvent Error(Err, "DevPadEditor.Editor:Redo", Error)
End Sub

Public Sub Undo()
On Error GoTo ErrHandler
    Dim cUndo As Undo_Item
    Dim cUndoNew As Undo_Item
    'abort if busy
    If m_blnBusy Then Exit Sub
    'abort if stack is empty
    If UndoStack.Count = 0 Then Exit Sub
    'reset flags
    bUndoing = False
    bRedoing = False
    'retreive undo item
    cUndo = UndoStack(UndoStack.Count)
    'save original
    cUndoNew = cUndo
    'remove from undo stack
    UndoStack.Remove (UndoStack.Count)
    'prevent redrawing
    LockMain
    Busy = True
    'goto pos
    rtfMain.SelStart = cUndo.lStart
    If cUndo.bAddTextPlain Then
        'select the text that was added
        rtfMain.SelLength = Len(cUndo.sAddText)
        If cUndo.ModifyType <> CommentText Then
            'convert an plaintext add to RTF (as we will have colourized it)
            cUndoNew.bAddTextPlain = False
            cUndoNew.sAddText = rtfMain.SelRTF
            cUndoNew.lAddTextLen = rtfMain.SelLength
        End If
        'delete the text that was added
        rtfMain.SelText = ""
        If cUndo.bDelTextPlain Then
            If cUndo.sDelText <> "" Then
                're-colour range with the deleted text (which will be restored)
                ReColourRange cUndo.sDelText, cUndo.lStart, cUndo.lStart + Len(cUndo.sDelText)
            End If
        Else
            'restore RTF text
            rtfMain.SelRTF = cUndo.sDelText
        End If
    Else
        ' select the previous text
        rtfMain.SelLength = cUndo.lAddTextLen
        ' delete it
        rtfMain.SelText = ""
        If cUndo.bDelTextPlain Then
            If cUndo.sDelText <> "" Then
                ReColourRange cUndo.sDelText, cUndo.lStart, cUndo.lStart + Len(cUndo.sDelText)
            End If
        Else
            ' replace with the text that was deleted
            rtfMain.SelRTF = cUndo.sDelText
        End If
    End If
    'add to redo stack
    RedoStack.Add cUndoNew
    'reset flags
    Busy = False
    UnlockMain
    'trigger sel change
    rtfMain_SelChange
    'we are undoing...
    bUndoing = True
    Exit Sub
ErrHandler:
    RaiseEvent Error(Err, "DevPadEditor.Editor:Undo", Error)
End Sub
'can we undo?
Public Property Get CanUndo() As Boolean
    CanUndo = Not (UndoStack.Count = 0)
End Property
Public Property Get CanRedo() As Boolean
    CanRedo = Not (RedoStack.Count = 0)
End Property
Public Property Get UndoText() As String
    On Error Resume Next
    If UndoStack.Count <> 0 Then UndoText = GetActionText(UndoStack(UndoStack.Count).ModifyType)
End Property
Public Property Get RedoText() As String
    On Error Resume Next
    If RedoStack.Count <> 0 Then RedoText = GetActionText(RedoStack(RedoStack.Count).ModifyType)
End Property
'retreive the action text
Private Function GetActionText(ModifyType As ModifyTypes) As String
    Select Case ModifyType
    Case DeleteText
        GetActionText = "Delete Text"
    Case AddText
        GetActionText = "Add Text"
    Case ReplaceText
        GetActionText = "Replace Text"
    Case PasteText
        GetActionText = "Paste Text"
    Case CutText
        GetActionText = "Cut Text"
    Case CommentText
        GetActionText = "Comment Block"
    Case UnCommentText
        GetActionText = "Uncomment Block"
    Case IndentText
        GetActionText = "Indent Block"
    Case OutdentText
        GetActionText = "Outdent Block"
    End Select
End Function
Public Sub Delete()
On Error GoTo ErrHandler
    'add delete event...
    AddUndoDeleteEvent 0, True, DeleteText
    rtfMain.SelText = ""
    Exit Sub
ErrHandler:
    RaiseEvent Error(Err, "DevPadEditor.Editor:Delete", Error)
End Sub
Private Sub AddUndoDeleteEvent(lAmount As Long, bIgnoreIfNoSel As Boolean, UndoType As ModifyTypes)
    Dim cUndo As Undo_Item
    
    With cUndo
        If rtfMain.SelLength = 0 And bIgnoreIfNoSel = False Then
            If lAmount = -1 And rtfMain.SelStart = 0 Then
                ' we aren't going anywhere!
                Exit Sub
            End If
            .lStart = IIf(rtfMain.SelStart = 0, 0, rtfMain.SelStart + lAmount)
            ' see what we are going to delete
            .sDelText = TextInRange(rtfMain.SelStart + lAmount + 1, IIf(lAmount < -1, Abs(lAmount), 1))
            'abort?
            If .sDelText = "" Then Exit Sub
            ' if there is part of vbCrLf selected
            ' set the length to 2 instead
            If InStr(1, Chr(10) & Chr(13), .sDelText) Then
                If .sDelText = Chr(10) Then
                    ' deleting end of CrLf
                    .lStart = .lStart - 1
                End If
                .sDelText = vbCrLf
            End If
            .bDelTextPlain = True
        Else
            'there is a selection... save what we are about to delete
            .lStart = rtfMain.SelStart
            If Len(rtfMain.SelText) <> rtfMain.SelLength And rtfMain.SelLength > 1 Then
                'we have half a vbcrlf... deselect it
                LockMain
                rtfMain.SelLength = rtfMain.SelLength - 2
                UnlockMain
            End If
            .sDelText = rtfMain.SelRTF
            .lDelTextLen = rtfMain.SelLength
        End If
        .ModifyType = UndoType
        'save undo event
        UndoStack.Add cUndo
        If .sDelText = vbCrLf Then
            pAdjustFlagPos CurrentLine, -1
        End If
    End With
End Sub
Private Sub AddUndoEvent(lStart As Long, sAdd As String)
    Dim cUndo As Undo_Item
    'save the char that is being inserted
    With cUndo
        .lStart = lStart
        .sAddText = sAdd
        .sDelText = rtfMain.SelText
        .bDelTextPlain = True
        .bAddTextPlain = True
        .ModifyType = IIf(.sDelText = "", AddText, ReplaceText)
    End With
    UndoStack.Add cUndo
End Sub
Public Property Get TextInRange(ByVal lStart As Long, ByVal lLen As Long, Optional lhWnd As Long = 0)
    Dim tR As TEXTRANGE
    Dim lR As Long
    Dim sText As String
    Dim b() As Byte
    Dim lEnd As Long
    
    If lhWnd = 0 Then lhWnd = rtfMain.hwnd
    'starts from 0, not 1
    lStart = lStart - 1
    'stupid!!!
    If lStart = -1 Then lStart = 0
    lEnd = lStart + lLen
    'set min/max values
    tR.chrg.cpMin = lStart
    tR.chrg.cpMax = lEnd
    'fill the buffer
    sText = String$(lEnd - lStart + 1, 0)
    'convert from Unicode
    b = StrConv(sText, vbFromUnicode)
    ' VB won't do the terminating null for you!
    ReDim Preserve b(0 To UBound(b) + 1) As Byte
    b(UBound(b)) = 0
    tR.lpstrText = VarPtr(b(0))
    lR = SendMessage(lhWnd, EM_GETTEXTRANGE, 0, tR)
    If (lR > 0) Then
        'return the text (in Unicode)
        sText = StrConv(b, vbUnicode)
        TextInRange = Left$(sText, lR)
    End If
End Property
'*** Find functions ***
Public Function FindFirst(lStart As Long, sFind As String) As Long
    ' + 1 so that it is like InStr
    FindFirst = rtfMain.Find(sFind, lStart - 1, , rtfNoHighlight) + 1
End Function
Public Function Find(ByVal sString As String, Optional ByVal vStart As Variant, Optional ByVal vEnd As Variant, Optional ByVal vOptions As Variant) As Long
    If IsMissing(vEnd) Then vEnd = -1
    Find = rtfMain.Find(sString, vStart, vEnd, vOptions)
End Function
'*** Line Numbering ***
Public Sub DrawLines(Optional bOverride As Boolean = False)
    'if the picbox is hidden, and we haven't been told to
    'override... abort
    'If picLines.Visible = False And bOverride = False Then Exit Sub
    Dim lLine       As Long
    Dim lCount      As Long
    Dim lCurrent    As Long
    Dim hBr         As Long
    Dim lEnd        As Long
    Dim lhDC        As Long
    Dim bComplete   As Boolean
    Dim tR          As RECT
    Dim tTR         As RECT
    Dim oCol        As OLE_COLOR
    Dim lStart      As Long
    Dim lEndLine    As Long
    Dim tPO         As POINTAPI
    Dim lLineHeight As Long
    Dim hPen        As Long
    Dim hPenOld     As Long
    Dim i As Long
    'get the hDC
    lhDC = picLines.hdc
    'get the line height


'    lLineHeight = tTR.Bottom - tTR.Top
    lLineHeight = lPixelsPerLine
    'get the line count
    lCount = LineCount
    
    
    'start and end positions
    lStart = rtfMain.SelStart
    lEnd = rtfMain.SelStart + rtfMain.SelLength - 1
    'current line...
    lCurrent = SendMessageLong(rtfMain.hwnd, EM_LINEFROMCHAR, rtfMain.SelStart, 0&)
    If (lEnd > lStart) Then
        'get the end line too
        lEndLine = SendMessageLong(rtfMain.hwnd, EM_LINEFROMCHAR, lEnd, 0)
    Else
        lEndLine = lCurrent
    End If
    lLine = GetFirstLineVisible
    If lCurrent < lLine And lEndLine > lLine Then
        'selection is above us...
        If lEndLine = lCurrent Then
            lEndLine = lLine
        End If
        lCurrent = lLine
        
    End If
    
    'get the size we have to draw in
    GetClientRect picLines.hwnd, tR
    lEnd = tR.Bottom - tR.Top
    tR.Right = 20
    hBr = CreateSolidBrush(TranslateColor(vbButtonFace))
    FillRect lhDC, tR, hBr
    DeleteObject hBr
    
    GetClientRect picLines.hwnd, tR
    If m_bShowLines Then
    
        
        tR.Left = 20
        hBr = CreateSolidBrush(TranslateColor(vbWhite))
        FillRect lhDC, tR, hBr
        DeleteObject hBr
    
        tR.Left = 20
        tR.Right = tR.Right - 2
        tR.Top = 0

        'get the maximum line width
        DrawText lhDC, CStr(lCount), -1, tR, DT_CALCRECT
        If tR.Right < 47 Then
            tR.Right = 47
        Else
            tR.Right = tR.Right + 2
        End If
        picLines.Width = (tR.Right + 2) * Screen.TwipsPerPixelX
        
        'default colour
        SetTextColor lhDC, TranslateColor(vbButtonShadow)
    End If
    tR.Bottom = tR.Top + lLineHeight
    Do
        If m_bShowLines Then
            ' Ensure correct colour:
            If (lLine = lCurrent) Then
                SetTextColor lhDC, TranslateColor(vbWindowText)
            ElseIf (lLine = lEndLine + 1) Then
                SetTextColor lhDC, TranslateColor(vbButtonShadow)
            End If
            ' Draw the line number:
            DrawText lhDC, CStr(lLine + 1), -1, tR, DT_RIGHT
        End If
        For i = 1 To m_lFlagCount Step 1
            If m_lFlag(i) = (lLine + 1) Then
                DrawIconEx lhDC, 0, tR.Top - 2, picFlag.Picture.Handle, 16, 16, 0&, 0&, 3
                Exit For
            End If
        Next
        ' Increment the line:
        lLine = lLine + 1
        ' Increment the position:
        OffsetRect tR, 0, lLineHeight
        If (tR.Bottom > lEnd) Or (lLine + 1 > lCount) Then
            bComplete = True
        End If
    Loop While Not bComplete
    
    'draw a line for 3d effect...
    MoveToEx lhDC, 20, 0, tPO
    hPen = CreatePen(PS_SOLID, 1, TranslateColor(vbButtonShadow))
    hPenOld = SelectObject(lhDC, hPen)
    LineTo lhDC, 20, lEnd
    SelectObject lhDC, hPenOld
    DeleteObject hPen
    
    If picLines.AutoRedraw Then
       picLines.Refresh
    End If
    
End Sub
' New Colorizing Code

Public Property Get IsRTF() As Boolean
    IsRTF = vSyntaxInfo.bRTF
End Property
Private Sub ResetColour(lStart As Long, lLen As Long)
    If lLen = -1 Then lLen = CharacterCount
    rtfMain.SelStart = lStart
    rtfMain.SelLength = lLen
    rtfMain.SelColor = vSyntaxInfo.vClr_Text
End Sub
'lStart starts from 1!
Public Sub ParseRange(Optional lStart As Long = 1, Optional lEnd As Long = -1, Optional bOverride As Boolean = False, Optional bOvr As Boolean = False)
    On Error GoTo ErrHandler
    
    Dim sBuffer     As String
    Dim sTmpWord    As String
    Dim nStartPos   As Long
    Dim nSelLen     As Long
    Dim nWordPos    As Long
    Dim nStart      As Long
    Dim lTempVal    As Long
    Dim bExit       As Boolean
    Dim i           As Long
    Dim sChar       As String
    Dim sCharTemp   As String
    Dim lOffset     As Long
    Dim lTemp       As Long
    Dim lLastVal    As Long 'used to check if we get in an un-ending loop!
    'if we are not code... abort
    If vSyntaxInfo.bCode = False Then Exit Sub
    'if we are HTML, and are in override mode, then goto the HTML Proc
    If vSyntaxInfo.bHTML And bOverride = False Then
        ParseHTMLRange lStart, lEnd, , bOvr
        Exit Sub
    End If
    'if there is no area between start and end, abort...
    'no point colourizing!
    If lEnd - lStart < 1 Then Exit Sub
    'reset flag
    m_bCancelColour = False
    'fill End parameter if we need to
    If lEnd = -1 Then lEnd = CharacterCount - 1
    Debug.Assert lEnd <> -1
    'reset text colour to black
    ResetColour lStart - 1, lEnd - lStart '- 1
    With rtfMain
        'correct start position if we need to...
        If lStart = 0 Then lStart = 1
        'get the buffer we are going to colour
        sBuffer = TextInRange(lStart, lEnd - lStart) & " "
        'set the offset position
        '(relative to real pos in richtextbox)
        lOffset = lStart - 1
        'reset start/end positions
        lStart = 1
        lEnd = Len(sBuffer)
        'empty buffer... abort
        If sBuffer <> " " Then
            'set where the first word is currently starting from
            nStartPos = 1
            
            'start looping through the buffer
            For i = lStart To lEnd Step 1
                If i < lLastVal Then
                    'colourizing error... we are going backwards!
                    RaiseEvent Error(vbObjectError + 999, "DevPadEditor.Editor:ParseRange", "Colourizing Error: Never-ending loop detected! Please send the current file to Developers Pad support.")
                    Exit For
                End If
                'save the position
                lLastVal = i
                'give way to windows
                DoEvents
                'check to see if we have cancelled
                If m_bCancelColour Then Exit For
                'get the current char
                sChar = Mid$(sBuffer, i, 1)
                'If sChar = "T" Then Stop
                'see if it is a char we are interested in or not
                If InStr(1, vSyntaxInfo.sInterestList, sChar) Then
                    'is it a seperator?
                    If InStr(1, vSyntaxInfo.sSeps, sChar) Then
                        'we have passed a word that might be a keyword
                        If Trim(sTmpWord) <> "" Then
                            'get it's position in the keywords list
                            nWordPos = InStr(1, vSyntaxInfo.sKeywords, "*" & sTmpWord & "*", vSyntaxInfo.vCaseSensitive)
                            If nWordPos <> 0 Then
                                'keyword found...
                                'colour it
                                'Debug.Print sTmpWord
                                .SelStart = nStartPos - 1 + lOffset
                                .SelLength = Len(sTmpWord)
                                If LCase$(.SelText) <> LCase$(sTmpWord) Then
                                    'Stop
                                Else
                                    If nWordPos > vSyntaxInfo.lSecondKeywordStart Then
                                        .SelColor = vSyntaxInfo.vClr_Keyword2
                                    Else
                                        .SelColor = vSyntaxInfo.vClr_Keyword
                                    End If
                                    'correct the case if we need to
                                    If vSyntaxInfo.bAutoCase Then .SelText = Mid$(vSyntaxInfo.sKeywords, nWordPos + 1, Len(sTmpWord))
                                End If
                            End If
                            
                        End If
                        'then check to see if this is an operator too...
                        If InStr(1, vSyntaxInfo.sOperators, sChar, 0) Then
                            'and if so, colour it
                            .SelStart = i + lOffset - 1
                            .SelLength = 1
                            .SelColor = vSyntaxInfo.vClr_Operator
                        End If
                        'clear the current word text
                        sTmpWord = ""
                        'set the word start pos
                        nStartPos = i + 1
                    Else
                        'another character of interest!
                        'see if it is a string character
                        If InStr(1, vSyntaxInfo.sStrings, sChar) Then
                            'is a string
                            'set the start position
                            lTemp = i + lOffset
                            'reset flag
                            bExit = False
                            'get the line... it can't span more than that
                            lTemp = .Find(vbCrLf, lTemp, , rtfNoHighlight) + 1
                            If lTemp <> 0 Then
                                sCharTemp = Mid$(sBuffer, i, lTemp - lOffset - i)
                                'do we need to check this?
                                'If Left$(sCharTemp, 1) <> vSyntaxInfo.sFalseQuote Then
                                    'not a string escape
                                    lTemp = 1
                                    Do
                                        lTemp = lTemp + 1
                                        'find the next string
                                        lTemp = InStr(lTemp, sCharTemp, sChar)
                                        If lTemp = 0 Or vSyntaxInfo.sFalseQuote = "" Then
                                            bExit = True
                                        ElseIf Mid$(sCharTemp, lTemp - 1, 1) <> vSyntaxInfo.sFalseQuote Then
                                            bExit = True
                                        End If
                                        
                                    Loop While bExit = False
                                'End If
                            End If
                            If lTemp = 0 Then
                                i = lEnd
                            Else
                                i = lTemp + i - 1
                            End If
                            sTmpWord = ""
                            nStartPos = i + 1
                            ' we are now at end of quote
                        ElseIf InStr(1, vSyntaxInfo.sOperators, "*" & sChar & "*") Then
                            rtfMain.SelStart = i + lOffset - 1
                            rtfMain.SelLength = 1
                            rtfMain.SelColor = vSyntaxInfo.vClr_Operator
                        Else
                            sCharTemp = LCase$(Mid$(sBuffer, i, m_lMaxCommentLen))
                            If vSyntaxInfo.sSingleComment = Left$(sCharTemp, Len(vSyntaxInfo.sSingleComment) And vSyntaxInfo.sSingleComment <> "") Or (vSyntaxInfo.sSingleCommentAlt = Left$(sCharTemp, Len(vSyntaxInfo.sSingleCommentAlt)) And vSyntaxInfo.sSingleCommentAlt <> "") Then
                                ' go back to start of comment
                                .SelStart = i - 1 + lOffset
                                ' go to the end of the line or file
                                nSelLen = .Find(Chr$(10), i - 1 + lOffset, , rtfNoHighlight) + 1
                                If nSelLen = 0 Then
                                    nSelLen = lEnd - i
                                Else
                                    nSelLen = nSelLen - i - lOffset
                                End If
                                ' Colour text
                                .SelLength = nSelLen
                                .SelColor = vSyntaxInfo.vClr_Comment
                                i = i + nSelLen
                                sTmpWord = ""
                                nStartPos = i + 1
                            Else
                                If vSyntaxInfo.bMultiComment Then
                                    'check if it is a multi-line comment
                                    sCharTemp = Mid$(sBuffer, i, Len(vSyntaxInfo.sMultiCommentStart))
                                    If InStr(1, vSyntaxInfo.sMultiCommentStart, sCharTemp) Then
                                        'multiline comment start
                                        .SelStart = i - 1 + lOffset
                                        nSelLen = .Find(vSyntaxInfo.sMultiCommentEnd, i - 1 + lOffset, , rtfNoHighlight) + 1 + Len(vSyntaxInfo.sMultiCommentEnd)
                                        If nSelLen = Len(vSyntaxInfo.sMultiCommentEnd) Then
                                            nSelLen = lEnd - i 'no end comment found
                                        Else
                                            nSelLen = nSelLen - i - lOffset 'found!
                                        End If
                                        'colour the text...
                                        .SelLength = nSelLen
                                        .SelColor = vSyntaxInfo.vClr_Comment
                                        i = i - 1 + nSelLen
                                        
                                        sTmpWord = ""
                                        nStartPos = i + 1
                                    End If
                                Else
                                    'build up current word...
                                    sTmpWord = sTmpWord & sChar
                                End If
                            End If
                        End If
                    End If
                Else
                    'build up current word
                    sTmpWord = sTmpWord & sChar
                End If
            Next
        End If
    End With
    Exit Sub
ErrHandler:
    Debug.Print "Colourizing Error: " & Err & ": " & Error & "(Colour.ParseRange)"
End Sub

Public Sub ParseHTMLRange(Optional lStart As Long = 1, Optional lEnd As Long = -1, Optional bOverride As Boolean = False, Optional bOvr As Boolean = False)

    On Error GoTo ErrHandler
    Dim sBuffer     As String
    Dim sChar       As String
    Dim sCharTemp   As String
    Dim lOffset     As Long
    Dim bHTMLExScript As Boolean
    Dim bInTag      As Boolean
    Dim bScript     As Boolean
    Dim i           As Long
    Dim bTagFound   As Boolean
    Dim nEndPos     As Long
    Dim nEndTag     As Long
    Dim sTemp       As String
    Dim bHadTag     As Boolean
        
    'abort if not code
    If vSyntaxInfo.bCode = False Then Exit Sub
    'no code to colourize?
    If lEnd - lStart - 1 < 1 And lEnd <> -1 Then Exit Sub
    'reset flag
    m_bCancelColour = False
    'do everything?
    If lEnd = -1 Then lEnd = CharacterCount - 1

    With rtfMain
        'reset the colour
        
        'get the buffer to colour
        sBuffer = TextInRange(lStart, lEnd - lStart) & " "
        'save it's offset against the real textbox
        lOffset = lStart - 1
        lStart = 1
        'set the end pos
        lEnd = Len(sBuffer)
        'colour the text blue, after first < tag
        nEndTag = InStr(1, sBuffer, "<")
        If nEndTag = 0 Then nEndTag = 1
        .SelStart = nEndTag + lOffset - 1 'lStart - 1
        .SelLength = lEnd - nEndTag
        .SelColor = vSyntaxInfo.vClr_HTMLTag
                
                
        If bOvr Then
            'we want to fake that we are in a <% tag...
            sBuffer = vSyntaxInfo.sHTMLExtensionStart & sBuffer
            lOffset = lOffset - Len(vSyntaxInfo.sHTMLExtensionStart)
            lEnd = lEnd + Len(vSyntaxInfo.sHTMLExtensionStart)
        End If
        'abort if nothing to colourize
        If sBuffer <> " " Then
            
            For i = lStart To lEnd
                'give way for a sec
                DoEvents
                'cancelled
                If m_bCancelColour Then Exit For
                'get the current char
                sChar = Mid$(sBuffer, i, 1)
                Select Case sChar
                Case "<"
                    'we have found a tag
                    bTagFound = True
                    'reset flag
                    bHTMLExScript = False
                    'we are a server-side script
                    If vSyntaxInfo.bHTMLExtension Then
                        'get 2 chars...
                        sCharTemp = Mid$(sBuffer, i, Len(vSyntaxInfo.sHTMLExtensionStart))
                        'is it an extension tag?
                        If vSyntaxInfo.sHTMLExtensionStart = sCharTemp Then
                            'yes...
                            bHTMLExScript = True
                            'get the end position of the tag
                            nEndTag = InStr(i + 1, sBuffer, vSyntaxInfo.sHTMLExtensionEnd)
                            If nEndTag = 0 Then
                                'not found... goto end
                                nEndTag = lEnd
                            End If
                            'we don't want to colour the 'false' tag that
                            'we have added if it is off the screen
                            If bOvr = False Or i <> 1 Then
                                'colour start and end tags
                                .SelStart = i + lOffset - 1
                                .SelLength = Len(vSyntaxInfo.sHTMLExtensionStart)
                                .SelColor = vSyntaxInfo.vClr_HTMLExTag
                            End If
                            'colour the end tag too...
                            If nEndTag <> lEnd Then
                                .SelStart = nEndTag + lOffset - 1
                                .SelLength = Len(vSyntaxInfo.sHTMLExtensionEnd)
                                .SelColor = vSyntaxInfo.vClr_HTMLExTag
                            End If
                            Debug.Assert nEndTag <> 0
                            'call the non-html parser
                            ParseRange (i) + Len(vSyntaxInfo.sHTMLExtensionStart) + lOffset, nEndTag + lOffset - Len(vSyntaxInfo.sHTMLExtensionStart) + 2, True
                            'move current pos to end of tag
                            i = nEndTag + Len(vSyntaxInfo.sHTMLExtensionStart) - 1 '2
                        End If
                    End If
                    If bHTMLExScript = False Then
                        sCharTemp = Mid$(sBuffer, i + 1, 6)
                        'check to see if it is a comment...
                        If Left$(sCharTemp, 3) = "!--" Then
                            'find the end of the comment...
                            nEndTag = InStr(i, sBuffer, "-->")
                            'no end of comment... set it as the end of the buffer
                            If nEndTag = 0 Then nEndTag = lEnd
                            'colour all the comment as a HTML comment
                            .SelStart = i - 1 + lOffset
                            .SelLength = nEndTag - (i - 3)
                            .SelColor = vSyntaxInfo.vClr_HTMLComment
    
                            i = nEndTag
                            ' we are out of tag
                            bInTag = False
                        Else
                            'are we in a script tag
                            If LCase$(sCharTemp) = "script" Then bScript = True
                            'get the end pos and = pos of the tag
                            'NOTE: this can give us a false value if there is a
                            '<% and %> within a quote
                            nEndTag = InStr(i + 1, sBuffer, ">")
                            'get the next = pos...
                            nEndPos = InStr(i + 1, sBuffer, "=")
                            If nEndTag = 0 Then nEndTag = lEnd
                            If (nEndPos = 0 Or nEndPos >= nEndTag) And nEndTag <> 0 Then
                                ' no =, skip to just before end of tag
                                i = nEndTag - 1
                            End If
                            'we are in a HTML tag
                            bInTag = True
                        End If
                    End If
                Case ">"
                    ' out of tag
                    bTagFound = True
                    If bScript Then 'in a script tag
                        nEndPos = InStr(i, sBuffer, "</script>", vbTextCompare)
                        'no end script tag.. set to end
                        If nEndPos = 0 Then nEndPos = lEnd + 1
                        'colour the text...
                        .SelStart = i + lOffset
                        .SelLength = nEndPos - (i + 1)
                        .SelColor = vSyntaxInfo.vClr_HTMLScript
                        i = nEndPos
                        bScript = False
                        bInTag = True
                    ElseIf bHTMLExScript Then
                        
                    Else
                        ' skip to next tag, colouring rest black
                        nEndPos = InStr(i + 1, sBuffer, "<")
                        .SelStart = i + lOffset
                        If nEndPos = 0 Then
                            nEndPos = lEnd - 1
                            
                        Else
                            nEndPos = nEndPos - 1
                        End If
                        .SelLength = nEndPos - i
                        .SelColor = vSyntaxInfo.vClr_Text
                        i = nEndPos
                        bInTag = False
                    End If
                Case "="
                    If bInTag Then
                        sTemp = Mid$(sBuffer, i + 1, 1)
                        Select Case sTemp
                        Case Chr$(34), "'"
                            sCharTemp = sTemp
                        Case Else
                            sCharTemp = " "
                            If Mid$(sBuffer, i + 2, 1) = Chr(34) Then
                                sCharTemp = Chr(34)
                                i = i + 1
                            End If
                        End Select
                        'find end of quote
                        nEndPos = InStr(i + 2, sBuffer, sCharTemp)
                        If nEndPos = 0 Then nEndPos = lEnd
                        If nEndTag < nEndPos Then
                            If sCharTemp = " " Then
                                nEndPos = nEndTag - 1
                                If nEndPos < i Then nEndPos = InStr(i + 1, sBuffer, ">") - 1
                                bInTag = False
                            End If
                        End If
                        'colour the text black...
                        .SelStart = i + lOffset
                        .SelLength = nEndPos - (i)
                        .SelColor = vSyntaxInfo.vClr_Text
                        i = nEndPos
                    End If
                End Select
            Next
            'no tag found... colour everything black
            If bTagFound = False Then
                rtfMain.SelStart = lStart + lOffset - 1
                rtfMain.SelLength = lEnd '- 1
                rtfMain.SelColor = vSyntaxInfo.vClr_Text
            End If
        End If
    End With
    Exit Sub
ErrHandler:
    Debug.Print "Colourizing Error: " & Err & ": " & Error & "(Colour.ParseHTMLRange)"
End Sub

Private Sub ParseVisibleRange(Optional bTriggered As Boolean = False, Optional bOverwrite As Boolean = False)
    Dim lEndLine As Long
    Dim lLine As Long
    Dim lCount As Long
    Dim lLineCount As Long
    Dim lStart As Long
    Dim lEnd As Long
    Dim lCharBegIndex As Long
    Dim lCharEndIndex As Long
    Dim m_hWnd As Long
    Dim lLastOpenTag As Long
    Dim cCursor As clsCursor
    Static lLastBegIndex As Long
    Static lLastEndIndex As Long
    Static lLastHeight As Long
    
    On Error GoTo ErrHandler
    'abort if not code...
    If vSyntaxInfo.bCode = False Then Exit Sub
    'textbox empty... abort
    If CharacterCount = 0 Then
        lLastLine = -1 'reset flag
        Exit Sub
    End If
    
    
    m_hWnd = rtfMain.hwnd
    If bTriggered And m_blnBusy Then
        DoEvents
        Exit Sub
    End If
    'get the first visible line
    lLine = SendMessageLong(m_hWnd, EM_GETFIRSTVISIBLELINE, 0, 0)
    'if the line hasn't changed, and we were triggered, abort
    If lLine = lLastLine And bTriggered Then
        Exit Sub
    ElseIf bResizingTriggered Then
        'triggered by a resize...
        'check if height is bigger
        If UserControl.Height <= lLastHeight Then Exit Sub
    End If
    Set cCursor = New clsCursor
    cCursor.SetCursor vbHourglass
    'save the height
    lLastHeight = UserControl.Height
    'we are busy
    Busy = True
    'Grabs the number of currently visible lines...
    'We can do this b/c the entire control uses
    'only 1 font size...if it didn't, then we'd
    'have to modify this...
    lCount = picLines.TextHeight("|Hy,^")
    lCount = rtfMain.Height / lCount
    'EM_LINEINDEX returns the character index of
    'the beginning of the line...This grabs the
    'character index for the first visible line...
    lCharBegIndex = SendMessageLong(m_hWnd, EM_LINEINDEX, lLine, 0)
    'have we moved up?
    If lLine < lLastLine And lLastLine - lLine < lCount And bTriggered And lLastLine <> -1 Then
        'we have moved up...
        'and the move is less than the screen size
        'colour from top, to the last index
        lCharEndIndex = lLastBegIndex
        lLastLine = lLine
    Else
        If lLine > lLastLine And lLine - lLastLine < lCount And bTriggered And lLastLine <> -1 Then
            'we have moved down
            'and the move is less than the screen size
            lCharBegIndex = lLastEndIndex
        End If

        lLastLine = lLine

        'This grabs the character index for the line
        'just after the visible area of the RTB...we
        'do this so we're sure to get all the chars
        'we need to do highlighting...We also need to
        'check and see if we're at the last line or not...
        lLine = lLine + lCount + 1
        lLineCount = SendMessageLong(m_hWnd, EM_GETLINECOUNT, 0, 0)
        If lLineCount <= lLine Then
            '
            lLine = lLineCount
            lCharEndIndex = SendMessageLong(m_hWnd, EM_LINEINDEX, lLine, 0)
            If lCharEndIndex = -1 Then lCharEndIndex = SendMessageLong(m_hWnd, EM_LINEINDEX, lLine - 1, 0)
            lCharEndIndex = SendMessageLong(m_hWnd, EM_LINELENGTH, lCharEndIndex, 0) + lCharEndIndex
        Else
            lCharEndIndex = SendMessageLong(m_hWnd, EM_LINEINDEX, lLine, 0)
            If lCharEndIndex = -1 Then lCharEndIndex = SendMessageLong(m_hWnd, EM_LINEINDEX, lLine - 1, 0)
        End If
    End If
    lLastBegIndex = lCharBegIndex
    lLastEndIndex = lCharEndIndex
    
    ' Return the text
    If vSyntaxInfo.bMultiComment Or vSyntaxInfo.bHTML Then
        Dim vStatus As ColourStatus
        vStatus = GetStatus(lCharBegIndex)
        If vStatus = InComment Then
            If vSyntaxInfo.bMultiComment Then
                ' include all the text upto the last comment
                lLastOpenTag = InStrRev(TextInRange(1, lCharBegIndex + 1), vSyntaxInfo.sMultiCommentStart)
                If lLastOpenTag <> 0 Then lCharBegIndex = lLastOpenTag - 1
            Else 'If vSyntaxInfo.bHTML Then
                lLastOpenTag = InStrRev(TextInRange(1, lCharBegIndex + 1), "<!--")
                If lLastOpenTag <> 0 Then lCharBegIndex = lLastOpenTag - 1
            End If
        ElseIf vStatus = InScript Then
            lLastOpenTag = InStrRev(TextInRange(1, lCharBegIndex + 1), "<script", , vbTextCompare)
            If lLastOpenTag <> 0 Then lCharBegIndex = lLastOpenTag - 1
        End If
    End If
    
    'prevent redrawing
    LockMain
    'allow current colourizing to cancel
    m_bCancelColour = True
    DoEvents
    m_bCancelColour = False
    'save cursor
    SaveCursorPos
    
    If vSyntaxInfo.bHTML Then
 '       If vStatus = OutTag Then
'            lLastOpenTag = InStrRev(TextInRange(1, lCharBegIndex + 1), "<")
'            If lLastOpenTag <> 0 Then
'                If vSyntaxInfo.bHTMLExtension Then
'                    GetStatus (lCharBegIndex)
'                    If bTempInHTMLExCode = False Then
'                        'lCharBegIndex = lLastOpenTag - 1
'                    End If
'                Else
'                    lCharBegIndex = lLastOpenTag - 1
'                End If
'            End If
'        End If
        ParseHTMLRange lCharBegIndex + 1, lCharEndIndex + 1, , (bTempInHTMLExCode And vSyntaxInfo.bHTML)
    Else
        ParseRange lCharBegIndex + 1, lCharEndIndex + 1
    End If
    'restore cursor
    RestoreCursorPos
    'unlock
    UnlockMain
    'we are not busy
    Busy = False
    Exit Sub
ErrHandler:
    RaiseEvent Error(Err, "DevPadEditor.Editor:ParseVisibleRange", Error)
End Sub

''you only need to call this once

Private Sub LoadSyntaxFile(ByVal sFile As String)
Dim i As Long
    
    vSyntaxInfo = m_cGlobal.LoadSyntaxFile(sFile, i)
    'calculate the max comment width...
    m_lMaxCommentLen = Len(vSyntaxInfo.sSingleComment)
    If Len(vSyntaxInfo.sSingleCommentAlt) > m_lMaxCommentLen Then m_lMaxCommentLen = Len(vSyntaxInfo.sSingleCommentAlt)
    
    sCurIndent = vSyntaxInfo.sIndent
    
    bInHTMLExCode = False
    If cboLanguage.Text <> m_cGlobal.SyntaxFile(i).sName Then
        bIgnoreEvent = True
        cboLanguage.Text = m_cGlobal.SyntaxFile(i).sName
        bIgnoreEvent = False
    End If
End Sub
Private Function IsBiggerThanScreen(lEnd As Long)
    Dim lBottomOfScreen As Long
    Dim lCount As Long
    Dim tR As RECT
    GetClientRect rtfMain.hwnd, tR
    lCount = UserControl.TextHeight("|Hy,^") / Screen.TwipsPerPixelY
    lCount = (tR.Bottom - tR.Top) / (lCount)
    
    lBottomOfScreen = SendMessageLong(rtfMain.hwnd, EM_LINEINDEX, GetFirstLineVisible + lCount + 1, 0)
    'If lBottomOfScreen = -1 Then lBottomOfScreen = lEnd + 2
    If lEnd > lBottomOfScreen Then
        IsBiggerThanScreen = True
    End If
End Function
Public Property Get LineIndex(lLine As Long) As Long
    LineIndex = SendMessage(rtfMain.hwnd, EM_LINEINDEX, lLine - 1, 0&)
End Property
Public Property Get CursorFile(bAbsolutePath As Boolean) As String
    Dim sText As String
    Dim lCursor As Long
    Dim lLastPos As Long
    Dim lNextPos As Long
    Dim sFile As String
    Dim i As Integer
    Dim sFindText As String
    'not saved... no good!
    If m_sFileName = "" Then Exit Property
    sText = LineText
    lCursor = rtfMain.SelStart - SendMessage(rtfMain.hwnd, EM_LINEINDEX, CurrentLine - 1, 0&)
    
    For i = 0 To 1
        sFindText = IIf(i = 0, """", "'")
        lLastPos = InStrRev(Left$(sText, lCursor), sFindText)
        lNextPos = InStr(lCursor + 1, sText, sFindText)
        If lLastPos <> 0 And lNextPos <> 0 Then
            If lLastPos <= lCursor And lNextPos > lCursor Then
                sFile = Mid$(sText, lLastPos + 1, lNextPos - lLastPos - 1)
                'if there is a ' in the string too, then continue...
                If InStr(1, sFile, "'") <> 0 And i = 0 Then sFile = ""
            End If
        End If
        If sFile <> "" Then Exit For
    Next
    'no extension... assume that it is not a file
    If GetExtension(sFile) = "" Then sFile = ""
    If sFile <> "" Then
        If Left$(sFile, 1) = "/" Then 'get root
            Dim vLocalPaths() As String
            'loop through local server paths for correct / path
            vLocalPaths = Split(GetSetting(REG_KEY, "Settings", "ServerLocalPath", "C:\Inetpub\wwwroot"), ";")
            For i = 0 To UBound(vLocalPaths)
                If LCase$(Left$(m_sFileName, Len(vLocalPaths(i)))) = LCase$(vLocalPaths(i)) Then
                    'remove /
                    sFile = Right$(sFile, Len(sFile) - 1)
                    'add the local path
                    sFile = vLocalPaths(i) & "\" & sFile
                    Exit For
                End If
            Next
        ElseIf Left$(sFile, 3) = "../" Then
            'calculate real absolute path
            sFile = GetAbsolutePath(sFile, GetFolder(m_sFileName))
        ElseIf InStr(1, sFile, "://") <> 0 Then
            'invalid
            'sFile = ""
        ElseIf Mid$(sFile, 2, 2) = ":\" Then '
            'path ok
        ElseIf LCase$(Left$(sFile, 7)) = "mailto:" Then
            'invalid
            'sFile = ""
        Else
            'add file path
            sFile = GetFolder(m_sFileName) & "\" & sFile
        End If
        
        If sFile <> "" Then
            sFile = Replace(sFile, "/", "\")
            lNextPos = InStr(1, sFile, "?")
            If lNextPos <> 0 Then
                sFile = Left$(sFile, lNextPos - 1)
            End If
            If Mid$(sFile, Len(sFile), 1) = "\" Then
                'default doc...!
                'what shall we do here?!
            End If
        End If
    End If
    CursorFile = sFile
End Property
Private Function GetAbsolutePath(ByVal sPath As String, ByVal sFilePath As String) As String
Dim sNewPath As String
    sNewPath = sFilePath
    'loop through the sPath, looking for ../
    Do While Left$(sPath, 3) = "../"
        'up another folder
        'remove folder from sNewPath
        sNewPath = Left$(sNewPath, InStrRev(sNewPath, "\") - 1)
        'remove ../
        sPath = Right$(sPath, Len(sPath) - 3)
    Loop
    sNewPath = sNewPath & "\" & sPath
    
    GetAbsolutePath = sNewPath
End Function
Public Sub InsertTag(sTag As String)
    Dim sTagName    As String
    Dim lPos        As Long
    Dim lLen        As Long
    Dim lStart      As Long
    Dim sText       As String
    'get the position of the space
    lPos = InStr(1, sTag, " ")
    'no space?
    If lPos = 0 Then lPos = Len(sTag)
    'get the name of the tag (minus any properties)
    sTagName = Mid$(sTag, 2, lPos - 2)
    With rtfMain
        LockMain
        'save the selection
        lLen = .SelLength
        lStart = .SelStart
        'add the tags around the text
        sText = .SelText
        sText = sTag & sText & "</" & sTagName & ">"
        'insert the text
        InsertCode sText, True
        'restore selection
        If lLen <> 0 Then
            .SelStart = lStart
            .SelLength = lLen + Len(sTag) + Len("</" & sTagName & ">")
        Else
            .SelStart = lStart + Len(sTag)
        End If
        UnlockMain
    End With
End Sub
Public Property Get TextRTF() As String
    TextRTF = rtfMain.TextRTF
End Property
Public Sub AddFlag(lLine As Long, Optional bAddOnly As Boolean = False)
Dim i As Long
    'check to see if it exists already
    For i = 1 To m_lFlagCount
        If m_lFlag(i) = lLine Then
            'remove flag
            If bAddOnly = False Then pRemoveFlag i
            Exit Sub
        End If
    Next
    m_lFlagCount = m_lFlagCount + 1
    ReDim Preserve m_lFlag(1 To m_lFlagCount)
    m_lFlag(m_lFlagCount) = lLine
    DrawLines
    RaiseEvent FlagsChanged
End Sub
Private Sub pRemoveFlag(Index As Long)
Dim i As Long
    For i = Index To m_lFlagCount - 1
        m_lFlag(i) = m_lFlag(i + 1)
    Next
    m_lFlagCount = m_lFlagCount - 1
    If m_lFlagCount = 0 Then
        Erase m_lFlag
    Else
        ReDim Preserve m_lFlag(1 To m_lFlagCount)
    End If
    DrawLines
    RaiseEvent FlagsChanged
End Sub
Private Sub pAdjustFlagPos(FromLine As Long, lAmount As Long)
Dim i As Long
    'move up flags
    For i = 1 To m_lFlagCount
        If m_lFlag(i) > FromLine Then
            m_lFlag(i) = m_lFlag(i) + lAmount
        End If
    Next
    If m_lFlagCount > 0 Then DrawLines
End Sub

' get the extension
Private Function GetExtension(sFileName As String) As String
    Dim lPos As Long
    lPos = InStrRev(sFileName, ".")
    If lPos = 0 Then
        GetExtension = ""
    Else
        GetExtension = LCase$(Right$(sFileName, Len(sFileName) - lPos))
    End If
End Function

Private Sub ClearStack(Stack As Collection)
    Dim i As Long
    If Not Stack Is Nothing Then
        On Error Resume Next
        i = 1
        Do While Stack.Count <> 0
            Stack.Remove (1)
        Loop
    End If
End Sub
Private Function LoWord(dwValue As Long) As Integer
    CopyMemory LoWord, dwValue, 2
End Function
Private Function MakeDWord(LoWord As Integer, HiWord As Integer) As Long
    MakeDWord = (HiWord * &H10000) Or (LoWord And &HFFFF&)
End Function
Private Function HiWord(dwValue As Long) As Integer
    CopyMemory HiWord, ByVal VarPtr(dwValue) + 2, 2
End Function
Private Sub SetDropHeight(cbo As ComboBox, lHeight As Long)
    MoveWindow cbo.hwnd, cbo.Left / Screen.TwipsPerPixelX, cbo.Top / Screen.TwipsPerPixelX, cbo.Width / Screen.TwipsPerPixelX, lHeight, 1
End Sub

Private Function GetFolder(ByVal sPath As String) As String
Dim lPos As Long
    If sPath = "" Then Exit Function
    If Right$(sPath, 1) = ":" Or Right$(sPath, 1) = ":\" Then
        'drive
    Else
        lPos = InStrRev(sPath, "\")
        If lPos <> 0 Then sPath = Left$(sPath, lPos - 1)
    End If
    GetFolder = sPath
End Function
Public Property Get FlagCount() As Long
    'return the number of flags
    FlagCount = m_lFlagCount
End Property
Public Function NextFlag(bLoop As Boolean) As Long
    NextFlag = pGetNextFlag(True, bLoop)
End Function
Public Function PreviousFlag(bLoop As Boolean) As Long
    PreviousFlag = pGetNextFlag(False, bLoop)
End Function
Public Function LastFlag() As Long
    LastFlag = pGetNextFlag(False, False, 2)
End Function
Public Function FirstFlag() As Long
    FirstFlag = pGetNextFlag(True, False, 1)
End Function
Private Function pGetNextFlag(bForward As Boolean, Optional bLoop As Boolean = True, Optional nPos As Long = 0) As Long
Dim lVal As Long
Dim lMinVal As Long
Dim lMaxVal As Long
Dim lStart As Long
Dim i As Long
    
    If nPos = 1 Then
        'first
        lStart = 0
    ElseIf nPos = 2 Then
        'last flag
        lStart = CharacterCount
    Else
        'from the current line
        lStart = CurrentLine
    End If
    
    'loop through all the flags
    For i = 1 To m_lFlagCount
        If (m_lFlag(i) - lStart > 0 And bForward) Then
            'we are after the current line
            'check if this is the shortest distance...
            If m_lFlag(i) < lVal Or lVal = 0 Then lVal = m_lFlag(i)
        ElseIf (m_lFlag(i) - lStart < 0 And Not bForward) Then
            'we are after the current line
            'check if this is the shortest distance...
            If m_lFlag(i) > lVal Or lVal = 0 Then lVal = m_lFlag(i)
        End If
        'save min/max value... used if no match is found
        If m_lFlag(i) < lMinVal Or lMaxVal = 0 Then lMinVal = m_lFlag(i)
        If m_lFlag(i) > lMaxVal Then lMaxVal = m_lFlag(i)
    Next
    
    If lVal = 0 Then
        If bLoop Then
            'go to first/last flag if no match, and we are allowing loops
            lVal = IIf(bForward, lMinVal, lMaxVal)
            pGetNextFlag = True
        End If
    Else
        'success.... flag found
        pGetNextFlag = True
    End If
    'go there...
    If lVal <> 0 Then pGetNextFlag = lVal 'rtfMain.SelStart = LineIndex(lVal)
End Function
Public Sub ClearFlags()
    Erase m_lFlag
    m_lFlagCount = 0
    DrawLines
    RaiseEvent FlagsChanged
End Sub
Private Sub pResetTimer()
Dim i As Long
    tmrColor.Enabled = False
    m_bCancelColour = True
    tmrColor.Enabled = True
End Sub
Public Sub GetFlags(ByRef lFlags() As Long)
Dim i As Long
    ReDim lFlags(1 To m_lFlagCount)
    For i = 1 To m_lFlagCount
        lFlags(i) = m_lFlag(i)
    Next
End Sub
Public Function StringEncode(ByVal sString As String, ByVal bToString As Boolean)
    'encodes a text for use in a string, or back again...
    'ie <a href="fred.htm"> becomes "<a href=\"fred.htm\">"
    If bToString Then
        If vSyntaxInfo.sFalseQuote <> "" Then
            'replace all \ with \\
            sString = Replace(sString, vSyntaxInfo.sFalseQuote, "\\")
        End If
        sString = Replace(sString, Left$(vSyntaxInfo.sStrings, 1), vSyntaxInfo.sStringEncoded)
        sString = """" & sString & """"
    Else
        sString = StripChar("""", sString)
        sString = Replace(sString, vSyntaxInfo.sStringEncoded, Left$(vSyntaxInfo.sStrings, 1))
        If vSyntaxInfo.sFalseQuote <> "" Then
            'replace all \ with \\
            sString = Replace(sString, "\\", vSyntaxInfo.sFalseQuote)
        End If
    End If
    StringEncode = sString
End Function
