Attribute VB_Name = "modCore"
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
Public Const PROJECT_EXTENSIONS = "*.dpp;*.vbp;*.vbg;*.dsp;*.vbproj;*.csproj"

'*** Enumerators ***
'Public Enum ERECViewModes
'    ercDefault = 0
'    ercWordWrap = 1
'    ercWYSIWYG = 2
'End Enum
'Public Enum ColourStatus
'    InTag = 0
'    OutTag = 1
'    InComment = 2
'    OutComment = 3
'    InScript = 4
'    InHTMLExtension = 5
'End Enum
'Public Enum ShowYesNoResult
'    Yes = 1
'    YesToAll = 2
'    No = 3
'    NoToAll = 4
'    None = 5
'    Cancelled = -1
'End Enum
'Public Enum SaveOptions
'    vbwNotSet = 0
'    vbwSave = 1
'    vbwDiscard = 2
'End Enum
Public Enum ShortcutTypes
    vbwFile = 0
    vbwFolder = 1
    vbwURL = 2
    vbwEmail = 3
    vbwLiveFolder = 4
    vbwTempLiveFolder = 5
    vbwTempLiveFile = 6
End Enum
Private Enum ESetWindowPosStyles
    SWP_SHOWWINDOW = &H40
    SWP_HIDEWINDOW = &H80
    SWP_FRAMECHANGED = &H20 ' The frame changed: send WM_NCCALCSIZE
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200 ' Don't do owner Z ordering
    SWP_NOREDRAW = &H8
    SWP_NOREPOSITION = SWP_NOOWNERZORDER
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_DRAWFRAME = SWP_FRAMECHANGED
    HWND_NOTOPMOST = -2
End Enum


'*** Public Constants ***

Public Const REG_KEY = "DevPad"

Public Const WM_USER = &H400
Public Const GW_CHILD = 5
Public Const WM_MDIACTIVATE = &H222

Public Const WM_SETREDRAW = &HB
Public Const WM_VSCROLL = &H115

Public Const SW_SHOW = 5
Public Const SW_HIDE = 0
Public Const WM_SETFOCUS = &H7
' file constants
Public Const MAX_PATH = 260
Public Const ERROR_NO_MORE_FILES = 18&
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100

Public Const CB_SETDROPPEDWIDTH = &H160&

'*** Private Constants ***

Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)

Private Const PS_SOLID = 0
Private Const SEE_MASK_INVOKEDLIST = &HC
Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Const MF_BYPOSITION = &H400&
Private Const WS_EX_CLIENTEDGE = &H200&
Private Const WS_EX_STATICEDGE = &H20000
Private Const WS_OVERLAPPED = &H0&
Private Const WS_POPUP = &H80000000
Private Const WS_CHILD = &H40000000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_VISIBLE = &H10000000
Private Const WS_DISABLED = &H8000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_CAPTION = &HC00000                 '/* WS_BORDER | WS_DLGFRAME  */
Private Const WS_BORDER = &H800000
Private Const WS_DLGFRAME = &H400000
Private Const WS_VSCROLL = &H200000
Private Const WS_HSCROLL = &H100000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const WS_GROUP = &H20000
Private Const WS_TABSTOP = &H10000
Private Const WS_SIZEBOX = WS_THICKFRAME

'*** DevPad Types ***

'Public Type TemplateInfo
'    sSyntax As String
'    sDescription As String
'    lSelLen As Long
'    lSelStart As Long
'End Type
Public Type InsertItem
    sName As String
    sValue As String
    lStart As Long
    lLen As Long
End Type
Public Type ProjectInfo
    sProjectName As String
    sProjectDescription As String
    sProjectAuthor As String
End Type
Public Type CMDLG_VALUES
    FileName As String
    FileTitle As String
    FilterIndex As Long
End Type
Public Type DefaultSettings
    sFont         As String
    nFontSize     As Integer
    nWordWrap     As ERECViewModes
End Type

'Ported to TLB file
'Public Type AddInInfo
'    sClassName      As String
'    sName           As String
'    sDescription    As String
'    sIcon           As String
'    bShowInTB       As Boolean
'    bShowInMenu     As Boolean
'    bLoadAtStartup  As Boolean
'    bLoaded         As Boolean
'End Type
'Public Type AddInInstance
'    lAddInIndex     As Long    'index to AddInInfo
'    lInstanceIndex  As Long 'index to add-in instance in clsAddIns
'    oTool           As DevPadAddInTlb.IDevPadTools
'End Type

'*** API Types ***
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type
Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime  As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh   As Long
    nFileSizeLow    As Long
    dwReserved0     As Long
    dwReserved1     As Long
    cFileName       As String * MAX_PATH
    cAlternate      As String * 14
End Type

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

Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long '  only used if FOF_SIMPLEPROGRESS
End Type

'*** API ***

'Windows...
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageByRef Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, wParam As Long, lParam As Any) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

' Cursor / Drawing
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As Any) As Long  ' lpPoint As POINTAPI) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As Any) As Long  ' lpPoint As POINTAPI) As Long
Public Declare Function PtInRect Lib "user32" (lprc As RECT, ByVal X As Long, ByVal Y As Long) As Long

' File Operations
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

'*** Private API ***
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function ShellExecuteEx Lib "shell32" (lpSEI As SHELLEXECUTEINFO) As Long
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Integer

Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long

Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

'*** Variables ***
Public vDefault         As DefaultSettings
Public CmDlg            As CMDLG_VALUES
Public cAddIns          As clsAddIns
Public cDocuments       As IDevPadDocuments
Public cDialog          As IDevPadDialog
Public cWorkspace       As New clsWorkspace
Public sFileFilter      As String
Public lProjectFilter    As Long
Public cGlobalEditor    As clsGlobalEditor
'Public cSyntaxFunctions As clsSyntaxFile
Public bStartUp         As Boolean
Public bClosing         As Boolean
Public bClosingMultiple As Boolean
Public sVBPath          As String

Private m_bInDevelopment As Boolean

Public Sub SetStatusBar(Optional sStatusText As String = "Ready", Optional sPanel As String = "Status")
On Error GoTo ErrHandler
    frmMainForm.cStatusBar.PanelText(sPanel) = sStatusText
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Core.SetStatusBar"
End Sub

Public Function OpenFile(Optional AddToProject As Boolean = False, Optional lTab As Long = -1) As Boolean
    Load frmOpen
    frmOpen.AddToProject = AddToProject
    frmOpen.Init False, lTab
End Function

'///////////////////////////////////////////////////////////////////////
'Public Sub LoadFile(ByVal sFileName As String, Optional sTitle As String, Optional bForceText As Boolean)
'    Dim frmForm As IDevPadDocument
'    Dim i As Long
'    Dim bRTF As Boolean
'    Dim bExisting As Boolean
'    On Error GoTo ErrHandler
'
'    If sFileName = "" Then Exit Sub 'stupid!
'    If Not IsMissing(sTitle) Then
'        ' We do not have file title, get one
'        sTitle = GetCaption(sFileName)
'    End If
'    ' Is that document open?
'    For i = 1 To cDocuments.Count
'        If UCase(sFileName) = UCase(cDocuments.Item(i).FileName) Then
'            'document found...
'            SendMessage frmMainForm.GetMDIClienthWnd, WM_MDIACTIVATE, cDocuments.Item(i).DocumenthWnd, 0
'            DoEvents
'            Exit Sub
'        End If
'    Next
'    If ActiveDoc Is Nothing Then
'        ' load a new doc, no prompt, do not display yet
'        Set frmForm = cDocuments.New(False) 'frmMainForm.LoadNewDoc(False)
'    Else
'        With ActiveDoc
'            If .Modified Or .Saved = True Then
'                ' load a new doc, no prompt, do not display yet
'                Set frmForm = cDocuments.New(False) ' .frmMainForm.LoadNewDoc(False)
'            Else
'                Set frmForm = ActiveDoc
'                bExisting = True
'            End If
'        End With
'    End If
'Back:
'    With frmForm
'        .LoadingFile = True
'
'        ' Set status bar
'        If Not bClosingMultiple Then SetStatusBar "Opening file...."
'       ' If IsRTF(sFileName) Then
'       '     vRTF = rtfRTF
'       ' Else
'            'vRTF = rtfText
'       ' End If
'        If GetExtension(sFileName) = "rtf" Then bRTF = True
'        If (bRTF) Then
'            .ShowLines = False
'            .ViewMode = ercWordWrap
'        End If
'        .Show
'        DoEvents
'        If .LoadFile(sFileName, bRTF, bForceText) = False Then
'            If bExisting = False Then Unload frmForm
'        Else
'            DoEvents
'            'set forms caption
'            .DocumentCaption = sTitle
'            frmMainForm.UpdateFileMenu sFileName, False
'            If Not bClosingMultiple Then frmMainForm.UpdateWindowList
'        End If
'    End With
'    If Not bClosingMultiple Then SetStatusBar
'    Exit Sub
'ErrHandler:
'    cDialog.ErrHandler Err, Error, "Core.LoadFile"
'End Sub

' get the extension
Public Function GetExtension(sFileName As String) As String
    Dim lPos As Long
    lPos = InStrRev(sFileName, ".")
    If lPos = 0 Then
        GetExtension = ""
    Else
        GetExtension = LCase$(Right$(sFileName, Len(sFileName) - lPos))
    End If
End Function

'Private Function IsRTF(sFileName As String) As Boolean
'On Error GoTo Failed
'    Dim intFileNum As Integer
'    Dim sResult As String
'    intFileNum = FreeFile
'    ' open file
'    Open sFileName For Binary As intFileNum
'    ' init sing
'    sResult = String(5, " ")
'    Get #intFileNum, 1, sResult
'    Close intFileNum
'    If sResult = "{\rtf" Then
'        IsRTF = True
'    Else
'Failed:
'        IsRTF = False
'    End If
'End Function

' Removes trailing nulls from a sing
Public Function StripTerminator(ByVal sString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(sString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(sString, intZeroPos - 1)
    Else
        StripTerminator = sString
    End If
End Function
Public Function StripChar(ByVal sChar As String, ByVal sString As String) As String
    If Left$(sString, Len(sChar)) = sChar Then sString = Right$(sString, Len(sString) - Len(sChar))
    If Right$(sString, Len(sChar)) = sChar Then sString = Left$(sString, Len(sString) - Len(sChar))
    StripChar = sString
End Function
Public Function GetFolder(ByVal sPath As String) As String
Dim lPos As Long
    If sPath = "" Then Exit Function
    If Right$(sPath, 1) = ":" Or Right$(sPath, 1) = ":\" Then
        'drive
    Else
        lPos = InStrRev(sPath, "\")
        If lPos <> 0 Then GetFolder = Left$(sPath, lPos - 1)
    End If
    'GetFolder = sPath
End Function
Public Function GetCaption(ByVal sPath As String) As String
    If IsDrive(sPath) Then
        GetCaption = sPath
    Else
        GetCaption = Right$(sPath, Len(sPath) - InStrRev(sPath, "\"))
    End If
End Function
Private Function IsDrive(ByVal sString As String) As Boolean
    If sString Like "*:" Or sString Like "*:\" Then IsDrive = True
End Function

Public Sub ExitDevPad()
    Dim i As Long
    If GetSetting(REG_KEY, "Settings", "NoExitError", 0) = 0 Then
        On Error Resume Next ' we don't want errors at this stage!
    End If
    If IsLoaded("frmMessage") Then frmMessage.Hide
    bClosing = True
    With frmMainForm
        ' remove toolbar+menu
        .rbrMain.RemoveAllRebarBands
        ' save settings
        If .WindowState = vbNormal Then
            SaveSetting REG_KEY, "WindowSettings", "MainLeft", .Left
            SaveSetting REG_KEY, "WindowSettings", "MainTop", .Top
            SaveSetting REG_KEY, "WindowSettings", "MainWidth", .Width
            SaveSetting REG_KEY, "WindowSettings", "MainHeight", .Height
        End If
        SaveSetting REG_KEY, "WindowSettings", "MainState", .WindowState
    End With
    ' unload all forms
    For i = 0 To Forms.Count - 1
        Unload Forms(i)
    Next
    Unload frmMainForm
    ' remove mutex
    EndApp
    Unload frmMainForm
End Sub
Public Property Get DocOpen() As Boolean
    DocOpen = (cDocuments.Count <> 0)
End Property



'Public Sub ErrHandler(lErrNum As Long, sErrorText As String, Optional sSource As String = "<Unknown>", Optional sDebugInfo As String = "")
'    If cDialog Is Nothing Then Set cDialog = New clsDialog
'    cDialog.cDialog.ErrHandler lErrNum, sErrorText, sSource
'End Sub
'Public Sub ShowWarning(sErrorText As String, Optional sSource As String = "", Optional sTitle As String = "")
'    If cDialog Is Nothing Then Set cDialog = New clsDialog
'    cDialog.cDialog.ShowWarning sErrorText
'End Sub


Public Function IsLoaded(sForm As String) As Boolean
    Dim i As Long
    IsLoaded = False
    For i = 0 To Forms.Count - 1
        If Forms(i).Name = sForm Then
            IsLoaded = True
            Exit Function
        End If
    Next
End Function

' Load and show the form
Public Sub LoadShow(frmForm As Form, Optional Modal As FormShowConstants)
    'Dim cCursor As New clsCursor
    'cCursor.SetCursor vbHourglass
    Load frmForm
    DoEvents
    'cCursor.ResetCursor
    frmForm.Show Modal, frmMainForm
End Sub

Public Sub AddAllFilesInDir(sDir As String, Optional cboCombo As Object = Nothing, Optional vItems As Collection)
    Dim lpFindFileData As WIN32_FIND_DATA, lFileHdl  As Long
    Dim sTemp As String, lRet As Long, bAutoAdd As Boolean
    bAutoAdd = Not (cboCombo Is Nothing)
    On Error Resume Next
    ' get a file handle
    lFileHdl = FindFirstFile(App.Path & "\" & sDir & "\*.txt", lpFindFileData)
    If lFileHdl <> -1 Then
        Do Until lRet = ERROR_NO_MORE_FILES
            DoEvents
            ' if it is a file
            sTemp = StripTerminator(lpFindFileData.cFileName) ', vbProperCase)
            If sTemp <> "." And sTemp <> ".." Then
                If bAutoAdd Then
                    cboCombo.AddItem Left$(sTemp, Len(sTemp) - 4)
                Else
                    vItems.Add Left$(sTemp, Len(sTemp) - 4)
                End If
            End If
            ' based on the file handle iterate through all files and dirs
            lRet = FindNextFile(lFileHdl, lpFindFileData)
            If lRet = 0 Then Exit Do
        Loop
    Else
        cDialog.ErrHandler 53, "The following folder does not exist, or is empty: " & App.Path & "\" & sDir, "AddAllFilesInDir"
    End If
    ' close the file handle
    lRet = FindClose(lFileHdl)
End Sub

Public Function InDevelopment() As Boolean
   '  Debug.Assert code not run in an EXE.  Therefore
   '  m_bInDevelopment variable is never set.
   Debug.Assert InDevelopmentHack() = True
   InDevelopment = m_bInDevelopment
End Function

Private Function InDevelopmentHack() As Boolean
   '  .... '
   m_bInDevelopment = True
   InDevelopmentHack = m_bInDevelopment
End Function
Public Sub LoadResStrings(ctlControls As Object)
    Dim ctl As Control
    Dim sCtlType As String
    Dim sText As String
    Dim i As Integer
    Dim nVal As Long
    On Error Resume Next
    'For nItem = 0 To ctlControls.Count
    For Each ctl In ctlControls
        
        sCtlType = TypeName(ctl)
        'Debug.Print sCtlType
        Select Case sCtlType
        Case "Label", "Menu", "CommandButton", "OptionButton", "CheckBox", "Frame"
            If IsNumeric(ctl.Caption) Then ctl.Caption = LoadResString(CInt(ctl.Caption))
        Case "TextBox"
           ' nVal = Val(ctl.Text)
            'If nVal <> 0 Then ctl.Text = LoadResString(CInt(ctl.Text))
            If IsNumeric(ctl.Text) And ctl.Tag = "" Then ctl.Text = LoadResString(CInt(ctl.Text))
            If ctl.BorderStyle <> 0 Then SetThin3DBorder ctl.hWnd
        Case "ComboBox", "ListBox"
            If sCtlType = "ListBox" Then SetThin3DBorder ctl.hWnd
            If sCtlType = "ComboBox" Then
                If ctl.Style = 1 Then SetThin3DBorder ctl.hWnd
            End If
            If ctl.Tag <> "NORES" Then
                If ctl.ListCount > 0 Then
                    For i = 0 To ctl.ListCount
                        sText = ctl.List(i)
                        If IsNumeric(sText) Then ctl.List(i) = LoadResString(CInt(sText))
                    Next i
                End If
            End If
        End Select
    Next
End Sub

Public Function LoadTextFile(sFile As String, Optional bError As Boolean = True) As String
If bError = False Then On Error Resume Next
    Dim iFileNum As Integer
    If Dir$(sFile) <> "" Then
        iFileNum = FreeFile
        Open sFile For Input As iFileNum
        LoadTextFile = Input(LOF(iFileNum), iFileNum)
        Close #iFileNum
    Else
        If bError = True Then Err.Raise 53
    End If
End Function


Public Function NumericOnly(KeyAscii As Integer) As Integer
    If Not IsNumeric(Chr$(KeyAscii)) And Not KeyAscii = vbKeyBack Then
        NumericOnly = 0
    Else
        NumericOnly = KeyAscii
    End If
End Function
Public Function IndexForKey(sKey As String) As Long
   ' Dim i As Long
   If sKey = "-1" Then
    IndexForKey = -1
   Else
    IndexForKey = frmMainForm.vbalMain.ItemIndex(UCase$(sKey))
   End If
End Function

Public Sub SetDropHeight(cbo As ComboBox, lHeight As Long, bMDIChild As Boolean)
    If bMDIChild Then
        MoveWindow cbo.hWnd, cbo.Left / Screen.TwipsPerPixelX, cbo.Top / Screen.TwipsPerPixelX, cbo.Width / Screen.TwipsPerPixelX, lHeight, 1
    Else
        Dim cRect As RECT
        GetWindowRect cbo.hWnd, cRect
        MoveWindow cbo.hWnd, cRect.Left / Screen.TwipsPerPixelX, cRect.Top / Screen.TwipsPerPixelY, (cRect.Right - cRect.Left), lHeight, 1
    End If
    
End Sub

Public Sub SetThin3DBorder(lhWnd As Long)
Dim lStyle As Long
    On Error Resume Next
    lStyle = GetWindowLong(lhWnd, GWL_EXSTYLE)
    lStyle = lStyle And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
    SetWindowLong lhWnd, GWL_EXSTYLE, lStyle
    SetNoBorder lhWnd
End Sub

Public Sub SetNoBorder(lhWnd As Long)
Dim lStyle As Long
    On Error Resume Next
    lStyle = GetWindowLong(lhWnd, GWL_STYLE)
    lStyle = lStyle And Not (WS_BORDER Or WS_DLGFRAME Or WS_CAPTION Or WS_BORDER Or WS_SIZEBOX Or WS_THICKFRAME)
    SetWindowLong lhWnd, GWL_STYLE, lStyle
    FrameChanged lhWnd
End Sub
Public Sub FrameChanged(hWnd As Long)
    SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
End Sub
' Returns the low 16-bit integer from a 32-bit long integer
Public Function LoWord(dwValue As Long) As Integer
    CopyMemory LoWord, dwValue, 2
End Function
Public Function MakeDWord(LoWord As Integer, HiWord As Integer) As Long
    MakeDWord = (HiWord * &H10000) Or (LoWord And &HFFFF&)
End Function
' Returns the hi 16-bit integer from a 32-bit long integer
Public Function HiWord(dwValue As Long) As Integer
    CopyMemory HiWord, ByVal VarPtr(dwValue) + 2, 2
End Function

Public Function TranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then TranslateColor = -1
End Function

Public Function Draw3DRect(ByVal hdc As Long, ByRef rcItem As RECT, ByVal oTopLeftColor As OLE_COLOR, ByVal oBottomRightColor As OLE_COLOR, Optional ByVal bMod As Boolean = False)
Dim hPen As Long
Dim hPenOld As Long
Dim tp As POINTAPI
   hPen = CreatePen(PS_SOLID, 1, TranslateColor(oTopLeftColor))
   hPenOld = SelectObject(hdc, hPen)
   MoveToEx hdc, rcItem.Left, rcItem.Bottom - 1, tp
   LineTo hdc, rcItem.Left, rcItem.Top
   LineTo hdc, rcItem.Right - 1, rcItem.Top
   SelectObject hdc, hPenOld
   DeleteObject hPen
   If (rcItem.Left <> rcItem.Right) Then
      hPen = CreatePen(PS_SOLID, 1, TranslateColor(oBottomRightColor))
      hPenOld = SelectObject(hdc, hPen)
      LineTo hdc, rcItem.Right - 1, rcItem.Bottom - 1
      LineTo hdc, rcItem.Left - (IIf(bMod, 1, 0)), rcItem.Bottom - 1
      SelectObject hdc, hPenOld
      DeleteObject hPen
   End If
End Function

' Disables the X
Public Sub RemoveCloseItem(lhWnd As Long)
Dim lMenu As Long
    ' Get the form's system menu handle.
    lMenu = GetSystemMenu(lhWnd, False)
    DeleteMenu lMenu, 6, MF_BYPOSITION
End Sub

Public Sub ClearStack(Stack As Collection)
    Dim i As Long
    If Not Stack Is Nothing Then
        On Error Resume Next
        i = 1
        Do While Stack.Count <> 0
            Stack.Remove (1)
        Loop
    End If
End Sub


Public Sub EnableTB(sKey As String, bEnabled As Boolean, Optional nTB As Integer = 0)
'On Error Resume Next
    frmMainForm.tbrMain(nTB).ButtonEnabled(sKey) = bEnabled
End Sub
Public Sub EnableItem(sKey As String, bEnabled As Boolean)
On Error Resume Next
    Dim lIndex As Long
    ' get item index

    'lindex = frmMainForm.ctlPopMenu.IDForIndex .IndexForKey(sKey)
    'frmMainForm.ctlPopMenu.Enabled(lindex) = bEnabled
    frmMainForm.ctlPopMenu.Enabled(sKey) = bEnabled
End Sub
Public Function GetTemplateInfo(ByVal sTemplate As String) As TemplateInfo
    Dim nFileNum As Integer
    Dim sLine As String
    Dim sKey As String
    Dim sValue As String
    Dim bFound As Boolean
    Dim bInSection As Boolean
    Dim lPos As Long
    On Error GoTo ErrHandler
    sTemplate = LCase$(sTemplate)
    nFileNum = FreeFile
    Open App.Path & "\_templates\index.ins" For Input As nFileNum
    Do While EOF(nFileNum) = False
        Line Input #nFileNum, sLine
        If Left$(sLine, 1) <> ";" And sLine <> "" Then
            Select Case Left$(sLine, 1)
            Case "["
                bInSection = (LCase$(sLine) = "[" & sTemplate & "]")
            Case Else
                If bInSection Then
                    lPos = InStr(1, sLine, "=")
                    sKey = LCase$(Trim$(Left$(sLine, lPos - 1)))
                    sValue = Trim$(Right$(sLine, Len(sLine) - lPos))
                    Select Case sKey
                    Case "synxfile"
                        GetTemplateInfo.sSyntax = sValue
                        bFound = True
                    Case "description"
                        GetTemplateInfo.sDescription = sValue
                    Case "selstart"
                        GetTemplateInfo.lSelStart = CLng(sValue)
                    Case "sellen"
                        GetTemplateInfo.lSelLen = CLng(sValue)
                    End Select
                End If
            End Select
        End If
    Loop
    If bFound = False Then
        'not found
        GetTemplateInfo.sSyntax = "text.stx"
    End If
ErrHandler:
    If Err Then cDialog.ErrHandler Err, Error, "Core.GetSyntaxFile"
    Close nFileNum
    Exit Function
End Function
Public Sub ShellExFunc(sVerb As String, sFile As String, lWin As Long)
Dim lVal As Long
'lval = ShellExecute(0, "explore", sFile, vbNullString, vbNullString, 1)
    Dim sei As SHELLEXECUTEINFO
    sei.hWnd = lWin
   ' ShellExecute lWin, sVerb, sFile, vbNullString, vbNullString, 0
    sei.lpVerb = sVerb
    If sVerb = "find" Then
        sei.lpDirectory = sFile & Chr$(0)
    Else
  '      sei.lpDirectory = sFile '& Chr$(0)
        sei.lpFile = sFile & Chr$(0)
    End If
    sei.fMask = SEE_MASK_INVOKEDLIST
    sei.cbSize = Len(sei)
    sei.nShow = 1
    ShellExecuteEx sei
End Sub
Public Sub ShellFunc(sFile As String, Optional vShow As VbAppWinStyle = 1, Optional sVerb As String = vbNullString)
    ShellExecute 0, sVerb, sFile, vbNullString, vbNullString, vShow
End Sub
Public Sub InitAddIns()
    On Error GoTo ErrHandler
    If cAddIns Is Nothing Then
        Set cAddIns = New clsAddIns
        cAddIns.ProcessAddIns
    End If
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Core.InitAddIns"
End Sub
Public Sub BuildPopupMenu(lType As Long)
    'why doesn't the VB Accelerator PopMenu control support
    'popup menus that are not available on the visible menus?!
    'we have to build and destroy them each time they are needed...
    Dim lParent As Long
    Dim cTemplates As Collection
    Dim i As Long
On Error GoTo ErrHandler
    With frmMainForm.ctlPopMenu
        Select Case lType
        Case 1
            lParent = .AddItem("docpopup", "mnuDocPopup", , , .MenuIndex("mnuHelp"))
            'frmMainForm.mnuDocPopup.Visible = False
            ' Document Popup Menu
            
            .AddItem LoadResString(147) & Chr(vbKeyTab) & "Ctrl+X", "EditPopCut", , , lParent, plGetIconIndex("Cut") ', , ,
            .AddItem LoadResString(148) & Chr(vbKeyTab) & "Ctrl+C", "EditPopCopy", , , lParent, plGetIconIndex("Copy")
            .AddItem LoadResString(149) & Chr(vbKeyTab) & "Ctrl+V", "EditPopPaste", , , lParent, plGetIconIndex("Paste")
            .AddItem LoadResString(150), "EditPopAppend", , , lParent, plGetIconIndex("Append")
            .AddItem "-", "EditPopSep1", , , lParent
            .AddItem LoadResString(142) & Chr(vbKeyTab) & "Ctrl+Z", "EditPopUndo", , , lParent, plGetIconIndex("Undo")
            .AddItem LoadResString(143) & Chr(vbKeyTab) & "Ctrl+Y", "EditPopRedo", , , lParent, plGetIconIndex("Redo")
            .AddItem "-", "EditPopSep2", , , lParent
            .AddItem LoadResString(111) & Chr(vbKeyTab) & "Ctrl+O", "EditPopOpen", , , lParent, plGetIconIndex("Open")
            .AddItem LoadResString(1279) & Chr(vbKeyTab) & "Ctrl+L", "EditPopOpenLinked", , , lParent, plGetIconIndex("JUMP")
            
        Case 2
            lParent = .AddItem("prjpopup", "mnuProjectPopup", , , .MenuIndex("mnuHelp"))
            'project menu...
            .AddItem LoadResString(1263), "PrjPopOpenFileDevPad", , , lParent, plGetIconIndex("Open")
            .AddItem LoadResString(1264), "PrjPopOpenFile", , , lParent
            .AddItem "-", "PrjPopSep1", , , lParent
            .AddItem LoadResString(136), "PrjPopRemoveItem", , , lParent, plGetIconIndex("Delete")
            .AddItem LoadResString(1265), "PrjPopRenameItem", , , lParent
            .AddItem LoadResString(1266), "PrjPopEditPath", , , lParent, plGetIconIndex("Edit")
        Case 3
            lParent = .AddItem("newpopup", "mnuNewPopup", , , .MenuIndex("mnuHelp"))
            'project menu...
            Set cTemplates = New Collection
            AddAllFilesInDir "\_templates", , cTemplates
            For i = 1 To cTemplates.Count
                .AddItem cTemplates(i), "mnuNewPop" & i, , , lParent, plGetIconIndex("PAD")
            Next
            .AddItem "-", "mnuNewPopupSep", , , lParent
            .AddItem LoadResString(1242), "mnuNewPopupMore", , , lParent
        Case 4
            'insert code window
            lParent = .AddItem("CodePopup", "mnuCodePopup", , , .MenuIndex("mnuHelp"))
            .AddItem LoadResString(161), "CodePopInsert", , , lParent
            .AddItem "-", "PrjPopSep1", , , lParent
            .AddItem LoadResString(1267), "CodePopNewEntry", , , lParent, plGetIconIndex("NEWTEXT")
            .AddItem LoadResString(1268), "CodePopEditEntry", , , lParent, plGetIconIndex("EDIT")
            .AddItem "-", "PrjPopSep2", , , lParent
            .AddItem LoadResString(135), "CodePopNewFolder", , , lParent, plGetIconIndex("FOLDERCLOSED")
            .AddItem "-", "PrjPopSep3", , , lParent
            .AddItem LoadResString(1265), "CodePopRename", , , lParent
            .AddItem LoadResString(1021) & Chr(vbKeyTab) & "Del", "CodePopDelete", , , lParent, plGetIconIndex("DELETE")
        Case 5
            'insert code window
            lParent = .AddItem("DocPopup", "mnuDocPopup", , , .MenuIndex("mnuHelp"))
            .AddItem LoadResString(112), "DocPopSave", , , lParent, plGetIconIndex("SAVE")
            .AddItem LoadResString(113), "DocPopSaveAs", , , lParent
            .AddItem "-", "PrjPopSep1", , , lParent
            .AddItem LoadResString(117), "DocPopClose", , , lParent
            .AddItem "-", "PrjPopSep2", , , lParent
            .AddItem LoadResString(177), "DocPopExplore", , , lParent, plGetIconIndex("EXPLORER")
        End Select
    End With
ErrHandler:
    Err = 0
End Sub
Public Sub ShowPopup(sIndex As String, Parent As Object)
    Dim lpPos As POINTAPI
    GetCursorPos lpPos
    ScreenToClient Parent.hWnd, lpPos
    frmMainForm.ctlPopMenu.ShowPopupMenu Parent, sIndex, lpPos.X * Screen.TwipsPerPixelX, lpPos.Y * Screen.TwipsPerPixelY
End Sub
Private Function plGetIconIndex(ByVal sKey As String) As Long
    If sKey = "-1" Then
        plGetIconIndex = -1
    Else
        On Error Resume Next
        plGetIconIndex = frmMainForm.vbalMain.ItemIndex(UCase$(sKey))
        If plGetIconIndex = 0 Then plGetIconIndex = -1
    End If
   ' If Err Then Debug.Print sKey
End Function
Public Sub DeletePopupMenu(lType As Long)
On Error Resume Next
    With frmMainForm.ctlPopMenu
        Select Case lType
        Case 1
            .RemoveItem ("mnuDocPopup")
        Case 2
            .RemoveItem ("mnuProjectPopup")
        Case 3
            .RemoveItem ("mnuNewPopup")
        Case 4
            .RemoveItem ("mnuCodePopup")
        Case 5
            .RemoveItem ("mnuDocPopup")
        End Select
    End With
End Sub

Public Sub GetVBPath()
Dim cR As clsRegistry
    If sVBPath = "" Then
        If GetSetting(REG_KEY, "Settings", "VBPath", "") = "" Then
            Set cR = New clsRegistry
            cR.ClassKey = HKEY_classes_root
            cR.SectionKey = "VisualBasic.Project\shell\open\command"
            cR.ValueType = REG_SZ
            cR.ValueKey = ""
            If cR.Value <> Empty Then
                sVBPath = Left$(cR.Value, Len(cR.Value) - 5)
            End If
            If sVBPath Like Chr$(34) & "*" & Chr$(34) Then
                sVBPath = Mid$(sVBPath, 2, Len(sVBPath) - 2)
            End If
            SaveSetting REG_KEY, "Settings", "VBPath", sVBPath
        Else
            sVBPath = GetSetting(REG_KEY, "Settings", "VBPath", "")
        End If
    End If
End Sub

Public Function DeleteFile(sFile As String) As Boolean
Dim op As SHFILEOPSTRUCT
    With op
        .hWnd = frmMainForm.hWnd
        .wFunc = FO_DELETE
        .pFrom = sFile & vbNullChar & vbNullChar
        .fFlags = FOF_ALLOWUNDO
    End With
    'why doesn't this func return a value if it is cancelled?!
    'we have no way of knowing... :-(
    SHFileOperation op
    DoEvents
End Function

Public Property Get ActiveDoc() As IDevPadDocument
    Set ActiveDoc = frmMainForm.ActiveForm
End Property
Public Function IsProject(sFile As String) As Boolean
Dim sExtension As String
    sExtension = GetExtension(sFile)
    If InStr(1, PROJECT_EXTENSIONS & ";", "." & sExtension & ";") Then IsProject = True
End Function
Public Sub LoadFileDefault(sFile As String, Optional bForceText As Boolean = False)
    ' We have a file:
    Select Case GetExtension(sFile)
    Case "vbp", "vbg", "dpp", "dsp", "csproj", "vbproj"
        frmMainForm.ShowProjectWindow
        ' open as project
        frmProject.OpenProject sFile
    Case "dpw"
        'workspace
        cWorkspace.Load sFile
    Case Else
        cDocuments.LoadFile sFile, bForceText
    End Select
End Sub
Public Sub MakeControlsFlat(Controls As Object, ByRef cFlatCombo() As clsFlatCombo, ByRef cFlatOption() As clsFlatOpt)
    Dim i As Long
    Dim lCountA As Long
    Dim lCountB As Long
    For i = 0 To Controls.Count - 1
        If TypeName(Controls(i)) = "ComboBox" Then
            lCountA = lCountA + 1
            ReDim Preserve cFlatCombo(1 To lCountA)
            Set cFlatCombo(lCountA) = New clsFlatCombo
            cFlatCombo(lCountA).Attach Controls(i).hWnd
        ElseIf TypeName(Controls(i)) = "CheckBox" Then
            lCountB = lCountB + 1
            ReDim Preserve cFlatOption(1 To lCountB)
            Set cFlatOption(lCountB) = New clsFlatOpt
            cFlatOption(lCountB).Attach Controls(i).hWnd
            Controls(i).Height = 252
        End If
    Next
End Sub
Public Sub RestoreControls(ByRef cFlatCombo() As clsFlatCombo, ByRef cFlatOption() As clsFlatOpt)
    Dim i As Long
    For i = 1 To UBound(cFlatCombo)
        Set cFlatCombo(i) = Nothing
    Next
    Erase cFlatCombo
    For i = 1 To UBound(cFlatOption)
        Set cFlatOption(i) = Nothing
    Next
    Erase cFlatOption
End Sub
