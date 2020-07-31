VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHookDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' ==========================================================================
' Class:    clsHookDlg
' Filename: clsHookDlg.cls
' Author:   James Crowley
' Date:     24 May 1998
' Note:     Incorporates code from Steve McMahon's Common Dialog DLL
'
' Provides custom dialogs...
' A wrapper around GCommonDialog to make it look more
' like the standard common dialog control.
' ==========================================================================

Implements ISubclass

Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200

Private Const WM_USER = &H400
Private Const WM_ACTIVATEAPP = &H1C
Private Const WM_SETFOCUS = &H7
Private Const WM_ACTIVATE = &H6
Private Const WM_NCACTIVATE = &H86
Private Const WM_SETTEXT = &HC
Private Const WM_NCDESTROY = &H82
Private Const WM_SETFONT = &H30
Private Const WM_NCPAINT = &H85
Private Const WM_MOUSEACTIVATE = &H21

Private Const WS_CHILD = &H40000000
Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)

Private Const WM_COMMAND = &H111
Private Const WS_CAPTION = &HC00000
Private Const WS_BORDER = &H800000
Private Const WS_DLGFRAME = &H400000
Private Const WS_THICKFRAME = &H40000
Private Const WS_SIZEBOX = WS_THICKFRAME
Private Const WS_EX_DLGMODALFRAME = &H1&

Private Const LOGPIXELSY = 90
Private Const DEFAULT_CHARSET = 1
Private Const MAX_PATH = 260
Private Const DS_MODALFRAME = &H80
Private Const BM_CLICK = &HF5
Private Const SW_HIDE = 0
Private Const SPI_GETWORKAREA = 48&

Private Const CDM_FIRST = (WM_USER + 100)
Private Const CDM_GETFILEPATH = (CDM_FIRST + &H1)
Private Const CDM_GETFOLDERPATH = (CDM_FIRST + &H2)
Private Const CDM_SETCONTROLTEXT = (CDM_FIRST + &H4)
Private Const CDM_HIDECONTROL = (CDM_FIRST + &H5)

Private Type POINTAPI
  x As Long
  y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Enum CMDLG_ITEMS
    READ_ONLY_CHECKBOX = &H410 'chx1 - The read-only check box
    FILTER_COMBO = &H470 'cmb1 - Drop-down combo box that displays the list of file type filters
    FILTER_COMBO_LABEL = &H441 'stc2 - Label for the cmb1 combo box
    FOLDER_COMBO = &H471 'cmb2 - Drop-down combo box that displays the current drive or folder, and that allows the user to select a drive or folder to open
    FOLDER_COMBO_LABEL = &H443 'stc4 - Label for the cmb2 combo box
    FILENAME_COMBO = &H47C 'cmb13 - Drop-down combobox that displays the name of the current file, allows the user to type the name of a file to open, and select a file that has been opened or saved recently.This is for earlier Explorer-compatible applications without hook or dialog template, or for Windows 2000 applications. Compare with edt1.
    FILENAME_TEXTBOX = &H480 'edt1 - Edit control that displays the name of the current file, or allows the user to type the name of the file to open. Compare with cmb13.
    FILENAME_LABEL = &H442 'stc3 - Label for the cmb13 combo box and the edt1 edit control
    FILE_LIST = &H460 'lst1 - List box that displays the contents of the current drive or folder
    FILE_LIST_LABEL = &H440 'stc1 -Label for the lst1 list box
    IDOK = 1 'The OK command button (push button)
    IDCANCEL = 2 ' The Cancel command button (push button)
    IDHELP = &H40E 'pshHelp - The Help command button (push button)
End Enum
'
'Public Enum DLG_ITEMS
'    CMD_OK = 0
'    CMD_CANCEL = 1
'    TXT_FILE = 2
'    CBO_FILTER = 3
'    LBL_FOLDER = 4
'    LBL_FILENAME = 5
'    LBL_TYPE = 6
'End Enum

' Properties to emulate the CommonDialog control:
Private m_bCancelError          As Boolean
Private m_sFilter               As String
Private m_lFilterIndex          As Long
Private m_sFileName             As String
Private m_lFlags                As Long
Private m_sDialogTitle          As String
Private m_sDefaultExt           As String
Private m_sInitDir              As String
Private m_lMaxFileSize          As Long
Private m_sFileTitle            As String
Private m_hDC                   As Long
Private m_bHookDialog           As Boolean
Private m_lFromPage As Long
Private m_oColor As OLE_COLOR
Private m_lCopies As Long
Private m_objObject As Object
Private m_iPrinterDefault       As Integer
Private m_lToPage               As Long
Private m_bFileDialog           As Boolean

Private m_cItemhWnd            As Collection
Private m_sFont                 As String
Private m_DialoghDC             As Long
Private m_bCustomDlg            As Boolean
Private m_lDialogLeft           As Long
Private m_lDialogTop            As Long
Private m_sTempCaption          As String
Private m_lhWnd                 As Long 'Owner hWnd
Private m_lParenthWnd           As Long 'Parent hWnd (Container window)
Private m_ContainerhWnd         As Long 'Container hWnd

Private MCommonDialog           As New GCommonDialog

Private hDlgParent              As Long
Private bOpenDlg                As Boolean
Private hTheDlg                 As Long 'handle to common dialog
Private bConsume                As Boolean 'consume current subclassing message?
Private bSubclassing            As Boolean 'are we subclassing?

Public Event InitDialog(ByVal hDlg As Long)
Public Event FileChange(ByVal hDlg As Long)
Public Event FolderChange(ByVal hDlg As Long)
Public Event TypeChange(ByVal hDlg As Long)
Public Event Show()
Public Event DialogClose()
Public Event DialogOK(bCancel As Boolean)


Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
'need lParam ByVal for some calls...
Private Declare Function SendMessage2 Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As Any) As Long  ' lpPoint As POINTAPI) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal e As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long


Public Property Get DialogHandle() As Long
    DialogHandle = hTheDlg
End Property
Public Property Get DialogParentHandle() As Long
    DialogParentHandle = hDlgParent
End Property
Public Property Let ParentFormhWnd(lHwnd As Long)
    m_lParenthWnd = lHwnd
End Property
Public Property Let ContainerhWnd(lHwnd As Long)
    m_ContainerhWnd = lHwnd
End Property
Public Property Let CustomDialog(bValue As Boolean)
    m_bCustomDlg = bValue
End Property
Public Property Get CustomDialog() As Boolean
    CustomDialog = m_bCustomDlg
End Property
Public Property Let DialogLeft(lValue As Long)
    m_lDialogLeft = lValue
End Property
Public Property Get DialogLeft() As Long
    DialogLeft = m_lDialogLeft
End Property
Public Property Let DialogTop(lValue As Long)
    m_lDialogTop = lValue
End Property
Public Property Get DialogTop() As Long
    DialogTop = m_lDialogTop
End Property
Public Property Get CurrentFolder() As String
    Dim nChars As Long
    Dim sNewFolder As String * MAX_PATH
    nChars = SendMessage(GetParent(hTheDlg), CDM_GETFOLDERPATH, MAX_PATH, ByVal sNewFolder)
    If (nChars > 0) Then
        CurrentFolder = Left$(sNewFolder, nChars - 1)
    End If
End Property
Public Property Get SelectedFile() As String
    Dim nChars As Long
    Dim sFile As String * MAX_PATH
    nChars = SendMessage(GetParent(hTheDlg), CDM_GETFILEPATH, MAX_PATH, ByVal sFile)
    If (nChars > 0) Then
        SelectedFile = Left$(sFile, nChars - 1)
        If GetExtension(SelectedFile) = "" Then SelectedFile = SelectedFile & "." & DefaultExt
    End If
End Property

Private Sub Class_Initialize()
    m_bHookDialog = True
    m_lDialogTop = -1
    m_lDialogLeft = -1
    m_lFilterIndex = 1
End Sub

Public Property Get ItemCaption(DlgItem As CMDLG_ITEMS) As String
    Dim sResult As String
    Dim lLen As Long

    sResult = Space(255) & Chr$(0)
    lLen = GetWindowText(ItemhWnd(DlgItem), sResult, Len(sResult))
    ItemCaption = Left$(sResult, lLen)
End Property
Public Property Let ItemCaption(DlgItem As CMDLG_ITEMS, sCaption As String)
    SendMessage2 hDlgParent, CDM_SETCONTROLTEXT, DlgItem, sCaption
End Property

Public Property Get ItemhWnd(DlgItem As CMDLG_ITEMS) As Long
    'retreive cached hWnd
    ItemhWnd = m_cItemhWnd("K" & DlgItem)
End Property
Public Function SetFont(sFont As String, hdc As Long) As Boolean
    If m_bCustomDlg Then
        m_sFont = sFont
        
        m_DialoghDC = hdc
        AddChildWindows hDlgParent, 0
    End If
End Function

Private Sub Class_Terminate()
    Set MCommonDialog = Nothing
End Sub

Public Sub HideItem(Item As CMDLG_ITEMS)
    'ShowWindow ItemhWnd(Item), SW_HIDE
    SendMessage hDlgParent, CDM_HIDECONTROL, Item, 0
End Sub

Public Sub SetItemPos(Item As CMDLG_ITEMS, lButtonhWnd As Long, lExtraTop As Long)
    If m_bCustomDlg Then
        Dim wLp As POINTAPI
        Dim wNewPos As POINTAPI
        Dim cRect As RECT
        Dim cContainerRect As RECT
        Dim wContainerPos As POINTAPI
        'get the pos of the item
        GetWindowRect ItemhWnd(Item), cRect
        ScreenToClient hTheDlg, wLp
        'get the pos of our container
        GetWindowRect m_ContainerhWnd, cContainerRect
        ScreenToClient m_lParenthWnd, wLp
        wContainerPos.x = wLp.x + cContainerRect.Left - 5
        
        If lExtraTop Then
            wContainerPos.y = wLp.y + cContainerRect.Top + cRect.Bottom + lExtraTop
        Else
            wContainerPos.y = wLp.y + cContainerRect.Top + cRect.Top
        End If
        MoveWindow lButtonhWnd, cRect.Left + wContainerPos.x, wContainerPos.y, cRect.Right - cRect.Left, cRect.Bottom - cRect.Top, True
    End If
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)

End Property
Private Property Get ISubclass_MsgResponse() As EMsgResponse
    Select Case CurrentMessage
    Case Else
        ISubclass_MsgResponse = emrPreprocess
    End Select
End Property
Public Sub ActivateWindow(hWnd As Long)
    SendMessage hWnd, WM_ACTIVATE, 1, 0
    SendMessage hWnd, WM_NCACTIVATE, 1, 0
End Sub
Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case iMsg
    Case WM_ACTIVATE
        If hWnd = m_lhWnd And wParam = 1 Then
            'activate this form, not the parent!
            SendMessage m_lParenthWnd, WM_ACTIVATE, wParam, 0
        ElseIf hWnd = hDlgParent Then
            'activate the parent form the same as the hooked dlg
            SendMessage m_lParenthWnd, WM_NCACTIVATE, wParam, 0
        End If
    End Select
End Function
Friend Sub DialogClose()
    RaiseEvent DialogClose
    If m_bCustomDlg Then
        If bSubclassing = True Then
            DetachMessage Me, hDlgParent, WM_ACTIVATE
            DetachMessage Me, m_lhWnd, WM_ACTIVATE
            bSubclassing = False
        End If
    End If
End Sub
Friend Function TypeChange(ByVal hDlg As Long) As Long
    RaiseEvent TypeChange(hDlg)
End Function
Public Property Get DialogWidth() As Long
    Dim tR As RECT
    GetWindowRect hDlgParent, tR
    DialogWidth = tR.Right - tR.Left
End Property
Public Property Get DialogHeight() As Long
    Dim tR As RECT
    GetWindowRect hDlgParent, tR
    DialogHeight = tR.Bottom - tR.Top
End Property
Friend Function InitDialog(ByVal hDlg As Long) As Long
    Dim lStyle As Long

    Dim tR As RECT
    Dim lObjhWnd As Long
    'save dialog handle
    hTheDlg = hDlg
    hDlgParent = GetParent(hDlg)

    If m_bCustomDlg = True Then
        If bSubclassing = False Then
            AttachMessage Me, hDlgParent, WM_ACTIVATE
            AttachMessage Me, m_lhWnd, WM_ACTIVATE
            bSubclassing = True
        End If
        lStyle = GetWindowLong(hDlgParent, GWL_STYLE)
        lStyle = lStyle And Not (WS_BORDER Or WS_DLGFRAME Or WS_CAPTION Or WS_BORDER Or WS_SIZEBOX Or WS_THICKFRAME)
        lStyle = lStyle Or WS_CHILD
        SetWindowLong hDlgParent, GWL_STYLE, lStyle
        lStyle = GetWindowLong(hDlgParent, GWL_EXSTYLE)
        lStyle = lStyle And Not WS_EX_DLGMODALFRAME
        SetWindowLong hDlgParent, GWL_EXSTYLE, lStyle
        GetClientRect hDlgParent, tR
        SetParent hDlgParent, m_ContainerhWnd
        
        MoveWindow hDlgParent, -5, 0, tR.Right - tR.Left, tR.Bottom - tR.Top, True
        Debug.Print tR.Right - tR.Left; tR.Bottom - tR.Top
    End If
    If m_lDialogLeft <> -1 Or m_lDialogTop <> -1 Then
        If m_lDialogLeft = -1 Then m_lDialogLeft = 0
        If m_lDialogTop = -1 Then m_lDialogTop = 0
        If m_bCustomDlg Then
            lObjhWnd = m_lParenthWnd
        Else
            lObjhWnd = hDlgParent
        End If
        GetWindowRect lObjhWnd, tR
        MoveWindow lObjhWnd, m_lDialogLeft, m_lDialogTop, tR.Right - tR.Left, tR.Bottom - tR.Top, False
    End If
    'get the hWnd's
    pRetreivehWnds hDlgParent
    RaiseEvent InitDialog(hDlg)
    RaiseEvent Show
    
    If m_bCustomDlg = True Then SendMessage m_lParenthWnd, WM_NCACTIVATE, 1, 0
End Function
Private Sub pRetreivehWnds(ParenthWnd As Long)
    Dim hWnd As Long
    
    Dim sClassName As String
    Dim sCaption As String
    Dim lLen As Long
    Dim lObjhWnd As Long
    Dim tR As RECT
    Do
        hWnd = FindWindowEx(ParenthWnd, hWnd, vbNullString, vbNullString)
        If hWnd <> 0 Then
            'get it's class name
            sClassName = Space(255) & Chr$(0)
            lLen = GetClassName(hWnd, sClassName, Len(sClassName))
            sClassName = Left$(sClassName, lLen)

            Select Case sClassName
            Case "Button"
                If ItemExists(IDCANCEL) Then
                    'item is help button
                    m_cItemhWnd.Add hWnd, "K" & IDHELP
                ElseIf ItemExists(IDOK) Then
                    'item is cancel button
                    m_cItemhWnd.Add hWnd, "K" & IDCANCEL
                ElseIf ItemExists(READ_ONLY_CHECKBOX) Then
                    'item is open/save button
                    m_cItemhWnd.Add hWnd, "K" & IDOK
                Else
                    'item is open/save button
                    m_cItemhWnd.Add hWnd, "K" & READ_ONLY_CHECKBOX
                End If
            Case "ComboBox"
                If ItemExists(FOLDER_COMBO) Then
                    'item is help button
                    m_cItemhWnd.Add hWnd, "K" & FILTER_COMBO
                Else
                    'item is cancel button
                    m_cItemhWnd.Add hWnd, "K" & FOLDER_COMBO
                End If
            Case "Static"
                If ItemExists(FILENAME_LABEL) Then
                    'item is filter label
                    m_cItemhWnd.Add hWnd, "K" & FILTER_COMBO_LABEL
                ElseIf ItemExists(FILE_LIST_LABEL) Then
                    'item is filename label
                    m_cItemhWnd.Add hWnd, "K" & FILENAME_LABEL
                ElseIf ItemExists(FOLDER_COMBO_LABEL) Then
                    'item is file list label (doesn't really exist any more...)
                    m_cItemhWnd.Add hWnd, "K" & FILE_LIST_LABEL
                Else
                    'item is label next to folder list
                    m_cItemhWnd.Add hWnd, "K" & FOLDER_COMBO_LABEL
                End If
            Case "Edit"
                'item is filename listbox
                m_cItemhWnd.Add hWnd, "K" & FILENAME_TEXTBOX
            Case "ListBox"
                m_cItemhWnd.Add hWnd, "K" & FILE_LIST
            End Select
        End If
    Loop While hWnd <> 0
End Sub
Private Function ItemExists(Item As CMDLG_ITEMS) As Boolean
    Dim lDummy As Long
    On Error Resume Next
    lDummy = m_cItemhWnd("K" & Item)
    If Err = 0 Then ItemExists = True
End Function
Friend Function FileChange(ByVal hDlg As Long) As Long
    RaiseEvent FileChange(hDlg)
End Function
Friend Function FolderChange(ByVal hDlg As Long) As Long
    RaiseEvent FolderChange(hDlg)
End Function
Friend Function ConfirmOK() As Boolean
Dim bCancel As Boolean
    bCancel = False
    RaiseEvent DialogOK(bCancel)
    If (bCancel) Then
        ConfirmOK = False
    Else
        ConfirmOK = True
    End If
End Function

Sub AddChildWindows(ByVal hwndParent As Long, ByVal Level As Long)
    Dim WT As String, CN As String, Length As Long, lHwnd As Long
    Dim lOldhWnd As Long
    Dim lFont As Long
        Dim nHeight As Long
    If hwndParent = 0 Then Exit Sub
    If Level = 0 Then
        lHwnd = hwndParent
    Else
        lHwnd = GetWindow(hwndParent, GW_CHILD)
    End If
    If lHwnd <> 0 Then
        nHeight = -MulDiv(8, GetDeviceCaps(m_DialoghDC, LOGPIXELSY), 72)
    Else
        Exit Sub
    End If
    Do While lHwnd <> 0
        lFont = CreateFont(nHeight, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, m_sFont)
        SendMessage lHwnd, WM_SETFONT, lFont, 1
        AddChildWindows lHwnd, Level + 1
        lHwnd = GetWindow(lHwnd, GW_HWNDNEXT)
    Loop
End Sub

Private Sub Reset()
    Set m_cItemhWnd = New Collection
    hTheDlg = 0
End Sub

Public Sub SimulateCancel()
    'simulate cancel click
    SetFocus ItemhWnd(IDOK)
    SendMessage ItemhWnd(IDCANCEL), BM_CLICK, 0, 0
End Sub
Public Sub SimulateOpen()
    'simulate cancel click
    SetFocus ItemhWnd(IDCANCEL)
    SendMessage ItemhWnd(IDOK), BM_CLICK, 0, 0
End Sub

Public Sub ParseMultiFileName( _
      ByRef sDir As String, _
      ByRef sFiles() As String, _
      ByRef iFileCount As Long _
   )
Dim iPos As Long
Dim iNextPos As Long
Dim sAllFiles As String
Dim i As Long

   iPos = InStr(m_sFileName, vbNullChar & vbNullChar)
   If iPos <> 0 Then
      ' multi names
      sAllFiles = Left$(m_sFileName, iPos - 1)
      iPos = 1
      iNextPos = InStr(sAllFiles, vbNullChar)
      Do While iNextPos <> 0
         If (sDir = "") Then
            sDir = Mid$(sAllFiles, iPos, iNextPos - iPos)
         Else
            iFileCount = iFileCount + 1
            ReDim Preserve sFiles(1 To iFileCount) As String
            sFiles(iFileCount) = Mid$(sAllFiles, iPos, iNextPos - iPos)
         End If
         iPos = iNextPos + 1
         iNextPos = InStr(iPos, sAllFiles, vbNullChar)
      Loop
      iFileCount = iFileCount + 1
      ReDim Preserve sFiles(1 To iFileCount) As String
      sFiles(iFileCount) = Mid$(sAllFiles, iPos)
   Else
      ' single file
      iFileCount = 1
      ReDim sFiles(1 To 1) As String
      For i = Len(m_sFileName) To 1 Step -1
         If Mid$(m_sFileName, i, 1) = "\" Then
            If (i > 1) Then
               sDir = Left$(m_sFileName, i - 1)
               sFiles(1) = Mid$(m_sFileName, i + 1)
            Else
               sDir = ""
               sFiles(1) = m_sFileName
            End If
            Exit Sub
         End If
      Next i
      sDir = ""
      sFiles(1) = m_sFileName
   End If
End Sub


Public Sub CentreDialog(Optional ByVal lhWndCentreTo As Long = 0)
'
Dim lHwnd As Long
Dim tWR As RECT, tDR As RECT
Dim tp As POINTAPI
'Dim lhWndCentreTo As Long
Dim lL As Long
Dim lT As Long
Dim lR As Long

    ' If we're showing a file dialog, then the rectangle is the
    ' parent of the dialog itself:
    If m_bCustomDlg = True Then
        lHwnd = m_lParenthWnd
    ElseIf m_bFileDialog Then
        lHwnd = hDlgParent
    Else
        lHwnd = hTheDlg
    End If
    
    GetWindowRect lHwnd, tDR

    If lhWndCentreTo <> 0 Then
        GetWindowRect lhWndCentreTo, tWR
    Else
        ' Assume the screen object:
        lR = SystemParametersInfo(SPI_GETWORKAREA, 0, tWR, 0)
        If (lR = 0) Then
            ' Call failed - just use standard screen:
            tWR.Left = 0
            tWR.Top = 0
            tWR.Right = Screen.Width \ Screen.TwipsPerPixelX
            tWR.Bottom = Screen.Height \ Screen.TwipsPerPixelY
        End If
    End If
    On Error GoTo 0
    If (tWR.Right > 0) And (tWR.Bottom > 0) Then
        lL = tWR.Left + (((tWR.Right - tWR.Left) - (tDR.Right - tDR.Left)) \ 2)
        lT = tWR.Top + (((tWR.Bottom - tWR.Top) - (tDR.Bottom - tDR.Top)) \ 2)
        MoveWindow lHwnd, lL, lT, (tDR.Right - tDR.Left), (tDR.Bottom - tDR.Top), 1
    End If
End Sub
Public Property Get CancelError() As Boolean
    CancelError = m_bCancelError
End Property
Public Property Let CancelError(ByVal bCancelError As Boolean)
    m_bCancelError = bCancelError
End Property
Public Property Get FileName() As String
    'return object's FileName property
    FileName = m_sFileName
End Property
Public Property Let FileName(ByVal sFileName As String)
    'assign object's FileName property
    m_sFileName = sFileName
End Property
Public Property Get Filter() As String
    'return object's Filter property
    Filter = m_sFilter
End Property
Public Property Let Filter(ByVal sFilter As String)
    'assign object's Filter property
    m_sFilter = sFilter
End Property
Public Property Get FilterIndex() As Long
    'return object's FilterIndex property
    FilterIndex = m_lFilterIndex
End Property
Public Property Let FilterIndex(ByVal lFilterIndex As Long)
    'assign object's FilterIndex property
    m_lFilterIndex = lFilterIndex
End Property

Public Property Get DefaultExt() As String
    'return object's DefaultExt property
    DefaultExt = m_sDefaultExt
End Property
Public Property Let DefaultExt(ByVal vNewValue As String)
    'assign object's DefaultExt property
    m_sDefaultExt = vNewValue
End Property
Public Property Get DialogTitle() As String
    'return object's FileName property
    DialogTitle = m_sDialogTitle
End Property
Public Property Let DialogTitle(ByVal vNewValue As String)
    'assign object's DialogTitle property
    m_sDialogTitle = vNewValue
End Property
Public Property Get flags() As Long
    'return object's Flags property
    flags = m_lFlags
End Property
Public Property Let flags(ByVal vNewValue As Long)
    'assign object's Flags property
    m_lFlags = vNewValue
End Property

Public Property Get hWnd() As Long
    'Return object's hWnd property
    hWnd = m_lhWnd
End Property
Public Property Let hWnd(ByVal vNewValue As Long)
    'Assign object's hWnd property
    m_lhWnd = vNewValue
End Property

Public Property Get InitDir() As String
    'Return object's InitDir property
    InitDir = m_sInitDir
End Property
Public Property Let InitDir(ByVal vNewValue As String)
    'Assign object's InitDir property
    m_sInitDir = vNewValue
End Property

Public Property Get MaxFileSize() As Long
    'Return object's MaxFileSize property
    MaxFileSize = m_lMaxFileSize
End Property
Public Property Let MaxFileSize(ByVal vNewValue As Long)
    'Assign object's MaxFileSize property
    m_lMaxFileSize = vNewValue
End Property
Public Property Get FileTitle() As String
    'return object's FileTitle property
    FileTitle = m_sFileTitle
End Property
Public Property Let FileTitle(ByVal vNewValue As String)
    'assign object's FileTitle property
    m_sFileTitle = vNewValue
End Property
Public Sub ShowOpen()
Dim bFileMustExist As Boolean
Dim bMultiSelect As Boolean
Dim bReadOnly As Boolean
Dim bHideReadOnly As Boolean
    Reset
    bOpenDlg = True
    m_bFileDialog = True
    bFileMustExist = FlagSet(m_lFlags, OFN_FILEMUSTEXIST)
    bMultiSelect = FlagSet(m_lFlags, OFN_ALLOWMULTISELECT)
    bReadOnly = FlagSet(m_lFlags, OFN_READONLY)
    bHideReadOnly = FlagSet(m_lFlags, OFN_HIDEREADONLY)
    
    If (m_lFilterIndex = 0) Then m_lFilterIndex = 1
    
    If Not (MCommonDialog.VBGetOpenFileName( _
        m_sFileName, _
        m_sFileTitle, _
        bFileMustExist, bMultiSelect, bReadOnly, bHideReadOnly, _
        m_sFilter, m_lFilterIndex, _
        m_sInitDir, _
        m_sDialogTitle, _
        m_sDefaultExt, _
        m_lhWnd, _
        m_lFlags, _
        m_bHookDialog, _
        Me)) Then
        pCommonDialogError
    End If
End Sub
Public Sub ShowSave()
Dim bOverWritePrompt As Boolean
    bOpenDlg = False
    Reset
    m_bFileDialog = True
    bOverWritePrompt = FlagSet(m_lFlags, OFN_OVERWRITEPROMPT)
    If Not (MCommonDialog.VBGetSaveFileName( _
        m_sFileName, _
        m_sFileTitle, _
        bOverWritePrompt, _
        m_sFilter, m_lFilterIndex, _
        m_sInitDir, _
        m_sDialogTitle, _
        m_sDefaultExt, _
        m_lhWnd, _
        m_lFlags, _
        m_bHookDialog, _
        Me)) Then
        pCommonDialogError
    End If
End Sub

Private Function ApiError(ByVal e As Long) As String
    Dim s As String, c As Long
    s = String(256, 0)
    c = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
                      FORMAT_MESSAGE_IGNORE_INSERTS, _
                      0, e, 0&, s, Len(s), ByVal 0)
    If c Then ApiError = Left$(s, c)
End Function

Private Function LastApiError() As String
    LastApiError = ApiError(Err.LastDllError)
End Function
Private Function FlagSet(ByVal lWord As Long, ByVal lFlagValue As Long)
    FlagSet = ((lWord And lFlagValue) = lFlagValue)
End Function

Private Sub pCommonDialogError()
    ' We have an error:
    If (MCommonDialog.APIReturn = 0) Then
        ' Cancelled:
        If (m_bCancelError) Then
            ' Note if your code stops here, that is because your error
            ' options in VB are "Break in Class Module".  Change your
            ' error option to "Break on Unhandled Errors" to see how
            ' this works at runtime.
            Err.Raise 20001, App.EXEName & ".clsHookDlg", "User selected cancel."
        End If
    Else
        Err.Raise eeBaseCommonDialog Or MCommonDialog.ExtendedError, App.EXEName & ".clsHookDlg"
    End If
End Sub

Public Property Get PrinterObject() As Object
    'Return object's Object property
    PrinterObject = m_objObject
End Property
Public Property Let PrinterObject(ByVal vNewValue As Object)
    'Assign object's Object property
    Set m_objObject = vNewValue
End Property
Public Property Get PrinterDefault() As Integer
    'Return object's PrinterDefault property
    PrinterDefault = m_iPrinterDefault
End Property
Public Property Let PrinterDefault(ByVal vNewValue As Integer)
    'Assign object's PrinterDefault property
    m_iPrinterDefault = vNewValue
End Property
Public Property Get ToPage() As Long
    'Return object's ToPage property
    ToPage = m_lToPage
End Property
Public Property Let ToPage(ByVal vNewValue As Long)
    'Assign object's ToPage property
    m_lToPage = vNewValue
End Property
Public Sub ShowColor()
Dim bAnyColor As Boolean
Dim bFullOpen As Boolean
Dim bDisableFullOpen As Boolean
Dim lColor As Long
    Reset
    m_bFileDialog = False
    lColor = TranslateColor(m_oColor)
    bAnyColor = FlagSet(m_lFlags, CC_AnyColor)
    bFullOpen = FlagSet(m_lFlags, CC_FullOpen)
    bDisableFullOpen = FlagSet(m_lFlags, CC_PreventFullOpen)
    If Not (MCommonDialog.VBChooseColor( _
            lColor, _
            bAnyColor, bFullOpen, bDisableFullOpen, _
            m_lhWnd, m_lFlags, _
            m_bHookDialog, _
            Me)) Then
        pCommonDialogError
    Else
        m_oColor = lColor
    End If
End Sub
Public Sub ShowPrinter()
Dim bDisablePageNumbers As Boolean
Dim bShowPrintToFile As Boolean
Dim bPrintToFile As Boolean
Dim bDisablePrintToFile As Boolean
Dim bCollate As Boolean
Dim bPreventWarning As Boolean
Dim bDisableSelection As Boolean
Dim ePR As EPrintRange
Dim iCopies As Integer
    Reset
    m_bFileDialog = False
    iCopies = m_lCopies
    bDisablePageNumbers = Not (FlagSet(m_lFlags, PD_PAGENUMS))
    bDisableSelection = FlagSet(m_lFlags, PD_NOSELECTION)
    bShowPrintToFile = Not (FlagSet(m_lFlags, PD_HIDEPRINTTOFILE))
    bDisablePrintToFile = FlagSet(m_lFlags, PD_DISABLEPRINTTOFILE)
    bPrintToFile = FlagSet(m_lFlags, PD_PRINTTOFILE)
    bCollate = FlagSet(m_lFlags, PD_COLLATE)
    bPreventWarning = FlagSet(m_lFlags, PD_NOWARNING)
    
    If Not (MCommonDialog.VBPrintDlg( _
        m_hDC, _
        ePR, _
        bDisablePageNumbers, _
        m_lFromPage, _
        m_lToPage, _
        bDisableSelection, _
        iCopies, _
        bShowPrintToFile, _
        bDisablePrintToFile, _
        bPrintToFile, _
        bCollate, _
        bPreventWarning, _
        m_lhWnd, _
        m_objObject, _
        m_lFlags, _
        m_bHookDialog, _
        Me)) Then
        pCommonDialogError
    Else
        ' Success
        m_lCopies = iCopies
    End If
End Sub
Public Property Get FromPage() As Long
    'Return object's FromPAge property
    FromPage = m_lFromPage
End Property
Public Property Let FromPage(ByVal vNewValue As Long)
    'Assign object's FromPage property
    m_lFromPage = vNewValue
End Property
Public Property Get Color() As OLE_COLOR
    'return object's Color property
    Color = m_oColor
End Property
Public Property Let Color(ByVal oColor As OLE_COLOR)
    'assign object's Color property
    m_oColor = oColor
End Property
Public Property Get Copies() As Long
    'return object's Copies property
    Copies = m_lCopies
End Property
Public Property Let Copies(ByVal vNewValue As Long)
    'assign object's Copies property
    m_lCopies = vNewValue
End Property
Private Function GetExtension(sFile As String)
    Dim lPos As Long
    lPos = InStrRev(sFile, ".")
    If lPos = 0 Then
        GetExtension = ""
    Else
        GetExtension = LCase$(Right$(sFile, Len(sFile) - lPos))
    End If
End Function
