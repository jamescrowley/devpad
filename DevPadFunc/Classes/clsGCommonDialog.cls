VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GCommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ==========================================================================
' Class:    GCommonDialog
' Filename: GCommonDialog.cls
' Author:   Steve McMahon, based on original by Bruce McKinney
' Date:     24 May 1998
' ==========================================================================


' ==========================================================================
' API declares:
' ==========================================================================
Public Enum EErrorCommonDialog
    eeBaseCommonDialog = 13450  ' CommonDialog
End Enum

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalCompact Lib "kernel32" (ByVal dwMinFree As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemoryStr Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Const MAX_PATH = 260
Private Const MAX_FILE = 260

Private Type OPENFILENAME
    lStructSize As Long          ' Filled with UDT size
    hwndOwner As Long            ' Tied to Owner
    hInstance As Long            ' Ignored (used only by templates)
    lpstrFilter As String        ' Tied to Filter
    lpstrCustomFilter As String  ' Ignored (exercise for reader)
    nMaxCustFilter As Long       ' Ignored (exercise for reader)
    nFilterIndex As Long         ' Tied to FilterIndex
    lpstrFile As String          ' Tied to FileName
    nMaxFile As Long             ' Handled internally
    lpstrFileTitle As String     ' Tied to FileTitle
    nMaxFileTitle As Long        ' Handled internally
    lpstrInitialDir As String    ' Tied to InitDir
    lpstrTitle As String         ' Tied to DlgTitle
    flags As Long                ' Tied to Flags
    nFileOffset As Integer       ' Ignored (exercise for reader)
    nFileExtension As Integer    ' Ignored (exercise for reader)
    lpstrDefExt As String        ' Tied to DefaultExt
    lCustData As Long            ' Ignored (needed for hooks)
    lpfnHook As Long             ' Ignored (good luck with hooks)
    lpTemplateName As Long       ' Ignored (good luck with templates)
End Type

Private Declare Function GetOpenFileName Lib "COMDLG32" _
    Alias "GetOpenFileNameA" (file As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32" _
    Alias "GetSaveFileNameA" (file As OPENFILENAME) As Long
Private Declare Function GetFileTitle Lib "COMDLG32" _
    Alias "GetFileTitleA" (ByVal szFile As String, _
    ByVal szTitle As String, ByVal cbBuf As Long) As Long

Public Enum EOpenFile
    OFN_READONLY = &H1
    OFN_OVERWRITEPROMPT = &H2
    OFN_HIDEREADONLY = &H4
    OFN_NOCHANGEDIR = &H8
    OFN_SHOWHELP = &H10
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_NOVALIDATE = &H100
    OFN_ALLOWMULTISELECT = &H200
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_PATHMUSTEXIST = &H800
    OFN_FILEMUSTEXIST = &H1000
    OFN_CREATEPROMPT = &H2000
    OFN_SHAREAWARE = &H4000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000
    OFN_EXPLORER = &H80000
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000
End Enum
Private Type TCHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
End Type

Private Declare Function ChooseColor Lib "COMDLG32.DLL" _
    Alias "ChooseColorA" (Color As TCHOOSECOLOR) As Long

Public Enum EChooseColor
    CC_RGBInit = &H1
    CC_FullOpen = &H2
    CC_PreventFullOpen = &H4
    CC_ColorShowHelp = &H8
' Win95 only
    CC_SolidColor = &H80
    CC_AnyColor = &H100
' End Win95 only
    CC_ENABLEHOOK = &H10
    CC_ENABLETEMPLATE = &H20
    CC_EnableTemplateHandle = &H40
End Enum
Private Type TPRINTDLG
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hdc As Long
    flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As Long
    lpSetupTemplateName As Long
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

'  DEVMODE collation selections
Private Const DMCOLLATE_FALSE = 0
Private Const DMCOLLATE_TRUE = 1

Private Declare Function PrintDlg Lib "COMDLG32.DLL" _
    Alias "PrintDlgA" (prtdlg As TPRINTDLG) As Integer

Public Enum EPrintDialog
    PD_ALLPAGES = &H0
    PD_SELECTION = &H1
    PD_PAGENUMS = &H2
    PD_NOSELECTION = &H4
    PD_NOPAGENUMS = &H8
    PD_COLLATE = &H10
    PD_PRINTTOFILE = &H20
    PD_PRINTSETUP = &H40
    PD_NOWARNING = &H80
    PD_RETURNDC = &H100
    PD_RETURNIC = &H200
    PD_RETURNDEFAULT = &H400
    PD_SHOWHELP = &H800
    PD_ENABLEPRINTHOOK = &H1000
    PD_ENABLESETUPHOOK = &H2000
    PD_ENABLEPRINTTEMPLATE = &H4000
    PD_ENABLESETUPTEMPLATE = &H8000
    PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
    PD_ENABLESETUPTEMPLATEHANDLE = &H20000
    PD_USEDEVMODECOPIES = &H40000
    PD_USEDEVMODECOPIESANDCOLLATE = &H40000
    PD_DISABLEPRINTTOFILE = &H80000
    PD_HIDEPRINTTOFILE = &H100000
    PD_NONETWORKBUTTON = &H200000
End Enum

Private Type DEVNAMES
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
End Type

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Type DevMode
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
Private Type POINTL
    x As Long
    y As Long
End Type

' Common dialog errors

Private Declare Function CommDlgExtendedError Lib "COMDLG32" () As Long

' Hook and notification support:
Private Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type
'// Structure used for all file based OpenFileName notifications
Private Type OFNOTIFY
    hdr As NMHDR
    lpOFN As Long           ' Long pointer to OFN structure
    pszFile As String ';        // May be NULL
End Type

'// Structure used for all object based OpenFileName notifications
Private Type OFNOTIFYEX
    hdr As NMHDR
    lpOFN As Long       ' Long pointer to OFN structure
    psf As Long
    LPVOID As Long          '// May be NULL
End Type

Private Type OFNOTIFYshort
    hdr As NMHDR
    lpOFN As Long
End Type

' Messages:
Private Const WM_INITDIALOG = &H110
Private Const WM_NOTIFY = &H4E
Private Const WM_USER = &H400
Private Const WM_GETDLGCODE = &H87
Private Const WM_NCDESTROY = &H82


' Notification codes:
Private Const H_MAX As Long = &HFFFF + 1
Private Const CDN_FIRST = (H_MAX - 601)
Private Const CDN_LAST = (H_MAX - 699)

'// Notifications when Open or Save dialog status changes
Private Const CDN_INITDONE = (CDN_FIRST - &H0)
Private Const CDN_SELCHANGE = (CDN_FIRST - &H1)
Private Const CDN_FOLDERCHANGE = (CDN_FIRST - &H2)
Private Const CDN_SHAREVIOLATION = (CDN_FIRST - &H3)
Private Const CDN_HELP = (CDN_FIRST - &H4)
Private Const CDN_FILEOK = (CDN_FIRST - &H5)
Private Const CDN_TYPECHANGE = (CDN_FIRST - &H6)
Private Const CDN_INCLUDEITEM = (CDN_FIRST - &H7)

Private Const CDM_FIRST = (WM_USER + 100)
Private Const CDM_LAST = (WM_USER + 200)

Private Const DWL_MSGRESULT = 0
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

' ==========================================================================
' Implementation:
' ==========================================================================

' Array of custom colors lasts for life of app
Private alCustom(0 To 15) As Long, fNotFirst As Boolean
Public Enum EPrintRange
    eprAll
    eprPageNumbers
    eprSelection
End Enum
Private m_lApiReturn As Long
Private m_lExtendedError As Long
Private m_dvmode As DevMode
Private m_oEventSink As clsHookDlg

Friend Function DialogHook( _
        ByVal hDlg As Long, _
        ByVal msg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long _
    )
Dim tNMH As NMHDR
'Dim tOFNs As OFNOTIFYshort
Dim tOFN As OFNOTIFY
Dim tOF As OPENFILENAME

    If Not (m_oEventSink Is Nothing) Then
        Select Case msg
        
        Case WM_INITDIALOG
            DialogHook = m_oEventSink.InitDialog(hDlg)
        Case WM_NOTIFY
            CopyMemory tOFN, ByVal lParam, Len(tOFN)
            'CopyMemory tNMH, ByVal tOFNs.hdr, Len(tNMH)
            'Debug.Print 'tOFNs.lpOFN
            'CopyMemory tNMH, ByVal lParam, Len(tNMH)
           ' tnmh.idfrom
            Select Case tOFN.hdr.code  'tNMH.code
            Case CDN_SELCHANGE
                ' Changed selected file:
                DialogHook = m_oEventSink.FileChange(hDlg)
            Case CDN_FOLDERCHANGE
                ' Changed folder:
                DialogHook = m_oEventSink.FolderChange(hDlg)
            Case CDN_FILEOK
                ' Clicked OK:
                If Not m_oEventSink.ConfirmOK() Then
                    SetWindowLong hDlg, DWL_MSGRESULT, 1
                    DialogHook = 1
                Else
                    SetWindowLong hDlg, DWL_MSGRESULT, 0
                End If
            Case CDN_HELP
                ' Help clicked
            Case CDN_TYPECHANGE
'                Dim pOFN As Long
'
'                pOFN = GlobalLock(tOFN.lpOFN)
'                CopyMemory tOF, ByVal pOFN, Len(tOF)
'                GlobalUnlock tOFN.lpOFN
                
                'CopyMemory tOF, ByVal lParam, Len(tNMH)
                ' CopyMemory tOF, ByVal tOFN.lpOFN, Len(tOF)
'                Debug.Print tOF.nFilterIndex
                DialogHook = m_oEventSink.TypeChange(hDlg)
            Case CDN_INCLUDEITEM
                ' Hmmm
            Case Else
                
            End Select
        Case WM_NCDESTROY
            m_oEventSink.DialogClose
        End Select
    End If
End Function


Public Property Get APIReturn() As Long
    'return object's APIReturn property
    APIReturn = m_lApiReturn
End Property
Public Property Get ExtendedError() As Long
    'return object's ExtendedError property
    ExtendedError = m_lExtendedError
End Property

Private Sub Class_Initialize()
#If fComponent Then
    InitColors
#End If
End Sub

Function VBGetOpenFileName(FileName As String, _
                           Optional FileTitle As String, _
                           Optional FileMustExist As Boolean = True, _
                           Optional MultiSelect As Boolean = False, _
                           Optional ReadOnly As Boolean = False, _
                           Optional HideReadOnly As Boolean = False, _
                           Optional Filter As String = "All (*.*)| *.*", _
                           Optional FilterIndex As Long = 1, _
                           Optional InitDir As String, _
                           Optional DlgTitle As String, _
                           Optional DefaultExt As String, _
                           Optional Owner As Long = -1, _
                           Optional flags As Long = 0, _
                           Optional Hook As Boolean = False, _
                           Optional EventSink As clsHookDlg _
                        ) As Boolean

    Dim opfile As OPENFILENAME, s As String, afFlags As Long
    
    m_lApiReturn = 0
    m_lExtendedError = 0

With opfile
    .lStructSize = Len(opfile)
    
    ' Add in specific flags and strip out non-VB flags
    
    .flags = (-FileMustExist * OFN_FILEMUSTEXIST) Or _
            (-MultiSelect * OFN_ALLOWMULTISELECT) Or _
             (-ReadOnly * OFN_READONLY) Or _
             (-HideReadOnly * OFN_HIDEREADONLY) Or _
             (flags And CLng(Not (OFN_ENABLEHOOK Or _
                                  OFN_ENABLETEMPLATE)))
    ' Owner can take handle of owning window
    If Owner <> -1 Then .hwndOwner = Owner
    ' InitDir can take initial directory string
    .lpstrInitialDir = InitDir
    ' DefaultExt can take default extension
    .lpstrDefExt = DefaultExt
    ' DlgTitle can take dialog box title
    .lpstrTitle = DlgTitle
    
    If (Hook) Then
        HookedDialog = Me
        .lpfnHook = lHookAddress(AddressOf DialogHookFunction)
        .flags = .flags Or OFN_ENABLEHOOK Or OFN_EXPLORER
        Set m_oEventSink = EventSink
    End If
    
    ' To make Windows-style filter, replace | and : with nulls
    Dim ch As String, i As Integer
    For i = 1 To Len(Filter)
        ch = Mid$(Filter, i, 1)
        If ch = "|" Or ch = ":" Then
            s = s & vbNullChar
        Else
            s = s & ch
        End If
    Next
    ' Put double null at end
    s = s & vbNullChar & vbNullChar
    .lpstrFilter = s
    .nFilterIndex = FilterIndex

   ' Pad file and file title buffers to maximum path
    If (.flags And OFN_ALLOWMULTISELECT) = OFN_ALLOWMULTISELECT Then
      s = FileName & String$(8192 - Len(FileName), 0)
      .lpstrFile = s
      .nMaxFile = 8192
      s = FileTitle & String$(8192 - Len(FileTitle), 0)
      .lpstrFileTitle = s
      .nMaxFileTitle = 8192
    Else
       s = FileName & String$(MAX_PATH - Len(FileName), 0)
      .lpstrFile = s
      .nMaxFile = MAX_PATH
      s = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
      .lpstrFileTitle = s
      .nMaxFileTitle = MAX_FILE
    End If
    ' All other fields set to zero
    
   m_lApiReturn = GetOpenFileName(opfile)
   Set m_oEventSink = Nothing
   ClearHookedDialog
   Select Case m_lApiReturn
   Case 1
      ' Success
      VBGetOpenFileName = True

      If (.flags And OFN_ALLOWMULTISELECT) = OFN_ALLOWMULTISELECT Then
         FileName = .lpstrFile 'StrZToStr(.lpstrFile)
      Else
         FileName = StrZToStr(.lpstrFile)
         FileTitle = StrZToStr(.lpstrFileTitle)
      End If
     ' Filename = StripTerminator(Filename)
     ' FileTitle = StripTerminator(FileTitle)
      flags = .flags
      ' Return the filter index
      FilterIndex = .nFilterIndex
      ' Look up the filter the user selected and return that
      Filter = FilterLookup(.lpstrFilter, FilterIndex)
      If (.flags And OFN_READONLY) Then ReadOnly = True
      
   Case 0
      ' Cancelled
      VBGetOpenFileName = False
      FileName = ""
      FileTitle = ""
      flags = 0
      FilterIndex = -1
      Filter = ""
   Case Else
      ' Extended error
      m_lExtendedError = CommDlgExtendedError()
      VBGetOpenFileName = False
      FileName = ""
      FileTitle = ""
      flags = 0
      FilterIndex = -1
      Filter = ""
   End Select
   Set m_oEventSink = Nothing
End With
End Function
Private Function lHookAddress(lPtr As Long) As Long
    lHookAddress = lPtr
End Function
Private Function StrZToStr(s As String) As String
    StrZToStr = Left$(s, lstrlen(s))
End Function

Function VBGetSaveFileName(FileName As String, _
                           Optional FileTitle As String, _
                           Optional OverWritePrompt As Boolean = True, _
                           Optional Filter As String = "All (*.*)| *.*", _
                           Optional FilterIndex As Long = 1, _
                           Optional InitDir As String, _
                           Optional DlgTitle As String, _
                           Optional DefaultExt As String, _
                           Optional Owner As Long = -1, _
                           Optional flags As Long, _
                           Optional Hook As Boolean = False, _
                           Optional EventSink As clsHookDlg _
                        ) As Boolean
            
    Dim opfile As OPENFILENAME, s As String

    m_lApiReturn = 0
    m_lExtendedError = 0

With opfile
    .lStructSize = Len(opfile)
    
    ' Add in specific flags and strip out non-VB flags
    .flags = (-OverWritePrompt * OFN_OVERWRITEPROMPT) Or _
             OFN_HIDEREADONLY Or _
             (flags And CLng(Not (OFN_ENABLEHOOK Or _
                                  OFN_ENABLETEMPLATE)))
    ' Owner can take handle of owning window
    If Owner <> -1 Then .hwndOwner = Owner
    ' InitDir can take initial directory string
    .lpstrInitialDir = InitDir
    ' DefaultExt can take default extension
    .lpstrDefExt = DefaultExt
    ' DlgTitle can take dialog box title
    .lpstrTitle = DlgTitle
    
    If (Hook) Then
        HookedDialog = Me
        .lpfnHook = lHookAddress(AddressOf DialogHookFunction)
        .flags = .flags Or OFN_ENABLEHOOK Or OFN_EXPLORER
        Set m_oEventSink = EventSink
    End If
    
    ' Make new filter with bars (|) replacing nulls and double null at end
    Dim ch As String, i As Integer
    For i = 1 To Len(Filter)
        ch = Mid$(Filter, i, 1)
        If ch = "|" Or ch = ":" Then
            s = s & vbNullChar
        Else
            s = s & ch
        End If
    Next
    ' Put double null at end
    s = s & vbNullChar & vbNullChar
    .lpstrFilter = s
    .nFilterIndex = FilterIndex

    ' Pad file and file title buffers to maximum path
    s = FileName & String$(MAX_PATH - Len(FileName), 0)
    .lpstrFile = s
    .nMaxFile = MAX_PATH
    s = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
    .lpstrFileTitle = s
    .nMaxFileTitle = MAX_FILE
    ' All other fields zero
    
    m_lApiReturn = GetSaveFileName(opfile)
    Set m_oEventSink = Nothing
    ClearHookedDialog
    Select Case m_lApiReturn
    Case 1
        VBGetSaveFileName = True
        FileName = StrZToStr(.lpstrFile)
        FileTitle = StrZToStr(.lpstrFileTitle)
        flags = .flags
        ' Return the filter index
        FilterIndex = .nFilterIndex
        ' Look up the filter the user selected and return that
        Filter = FilterLookup(.lpstrFilter, FilterIndex)
    Case 0
        ' Cancelled:
        VBGetSaveFileName = False
        FileName = ""
        FileTitle = ""
        flags = 0
        FilterIndex = 0
        Filter = ""
    Case Else
        ' Extended error:
        VBGetSaveFileName = False
        m_lExtendedError = CommDlgExtendedError()
        FileName = ""
        FileTitle = ""
        flags = 0
        FilterIndex = 0
        Filter = ""
    End Select
End With
End Function

Private Function FilterLookup(ByVal sFilters As String, ByVal iCur As Long) As String
    Dim iStart As Long, iEnd As Long, s As String
    iStart = 1
    If sFilters = "" Then Exit Function
    Do
        ' Cut out both parts marked by null character
        iEnd = InStr(iStart, sFilters, vbNullChar)
        If iEnd = 0 Then Exit Function
        iEnd = InStr(iEnd + 1, sFilters, vbNullChar)
        If iEnd Then
            s = Mid$(sFilters, iStart, iEnd - iStart)
        Else
            s = Mid$(sFilters, iStart)
        End If
        iStart = iEnd + 1
        If iCur = 1 Then
            FilterLookup = s
            Exit Function
        End If
        iCur = iCur - 1
    Loop While iCur
End Function

Function VBGetFileTitle(sFile As String) As String
    Dim sFileTitle As String, cFileTitle As Integer

    cFileTitle = MAX_PATH
    sFileTitle = String$(MAX_PATH, 0)
    cFileTitle = GetFileTitle(sFile, sFileTitle, MAX_PATH)
    If cFileTitle Then
        VBGetFileTitle = ""
    Else
        VBGetFileTitle = Left$(sFileTitle, InStr(sFileTitle, vbNullChar) - 1)
    End If

End Function
#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.EXEName & ".CommonDialog"
        Err.Raise COMError(e), sSource, sText
    Else
        ' Raise standard Visual Basic error
        sSource = App.EXEName & ".VBError"
        Err.Raise e, sSource
    End If
End Sub
#End If

Private Sub StrToBytes(ab() As Byte, s As String)
    If IsArrayEmpty(ab) Then
        ' Assign to empty array
        ab = StrConv(s, vbFromUnicode)
    Else
        Dim cab As Long
        ' Copy to existing array, padding or truncating if necessary
        cab = UBound(ab) - LBound(ab) + 1
        If Len(s) < cab Then s = s & String$(cab - Len(s), 0)

        CopyMemoryStr ab(LBound(ab)), s, cab
    End If
End Sub

Private Function BytesToStr(ab() As Byte) As String
    BytesToStr = StrConv(ab, vbUnicode)
End Function

Private Function COMError(e As Long) As Long
    COMError = e Or vbObjectError
End Function

Private Function IsArrayEmpty(va As Variant) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = va(LBound(va))
    IsArrayEmpty = (Err <> 0)
End Function














' ChooseColor wrapper
Function VBChooseColor(Color As Long, _
                       Optional AnyColor As Boolean = True, _
                       Optional FullOpen As Boolean = False, _
                       Optional DisableFullOpen As Boolean = False, _
                       Optional Owner As Long = -1, _
                       Optional flags As Long, _
                       Optional Hook As Boolean = False, _
                      Optional EventSink As clsHookDlg _
                    ) As Boolean

    Dim chclr As TCHOOSECOLOR
    chclr.lStructSize = Len(chclr)
    
    ' Color must get reference variable to receive result
    ' Flags can get reference variable or constant with bit flags
    ' Owner can take handle of owning window
    If Owner <> -1 Then chclr.hwndOwner = Owner

    ' Assign color (default uninitialized value of zero is good default)
    chclr.rgbResult = Color

    ' Mask out unwanted bits
    Dim afMask As Long
    afMask = CLng(Not (CC_ENABLEHOOK Or _
                       CC_ENABLETEMPLATE))
    ' Pass in flags
    chclr.flags = afMask And (CC_RGBInit Or _
                  IIf(AnyColor, CC_AnyColor, CC_SolidColor) Or _
                  (-FullOpen * CC_FullOpen) Or _
                  (-DisableFullOpen * CC_PreventFullOpen))

'    If (Hook) Then
'        HookedDialog = Me
'        chclr.lpfnHook = lHookAddress(AddressOf CCHookProc)
'        chclr.flags = chclr.flags Or CC_ENABLEHOOK
'        Set m_oEventSink = EventSink
'    End If
    
    ' If first time, initialize to white
    If fNotFirst = False Then InitColors

    chclr.lpCustColors = VarPtr(alCustom(0))
    ' All other fields zero
    
    m_lApiReturn = ChooseColor(chclr)
    Set m_oEventSink = Nothing
    ClearHookedDialog
    
    Select Case m_lApiReturn
    Case 1
        ' Success
        VBChooseColor = True
        Color = chclr.rgbResult
    Case 0
        ' Cancelled
        VBChooseColor = False
        Color = -1
    Case Else
        ' Extended error
        m_lExtendedError = CommDlgExtendedError()
        VBChooseColor = False
        Color = -1
    End Select

End Function

Friend Sub InitColors()
    Dim i As Integer
    ' Initialize with first 16 system interface colors
    For i = 0 To 15
        alCustom(i) = GetSysColor(i)
    Next
    fNotFirst = True
End Sub

' Property to read or modify custom colors (use to save colors in registry)
Public Property Get CustomColor(i As Integer) As Long
    ' If first time, initialize to white
    If fNotFirst = False Then InitColors
    If i >= 0 And i <= 15 Then
        CustomColor = alCustom(i)
    Else
        CustomColor = -1
    End If
End Property

Public Property Let CustomColor(i As Integer, iValue As Long)
    ' If first time, initialize to system colors
    If fNotFirst = False Then InitColors
    If i >= 0 And i <= 15 Then
        alCustom(i) = iValue
    End If
End Property
' PrintDlg wrapper
Function VBPrintDlg(hdc As Long, _
                    Optional PrintRange As EPrintRange = eprAll, _
                    Optional DisablePageNumbers As Boolean, _
                    Optional FromPage As Long = 1, _
                    Optional ToPage As Long = &HFFFF, _
                    Optional DisableSelection As Boolean, _
                    Optional Copies As Integer, _
                    Optional ShowPrintToFile As Boolean, _
                    Optional DisablePrintToFile As Boolean = True, _
                    Optional PrintToFile As Boolean, _
                    Optional Collate As Boolean, _
                    Optional PreventWarning As Boolean, _
                    Optional Owner As Long, _
                    Optional Printer As Object, _
                    Optional flags As Long, _
                    Optional Hook As Boolean = False, _
                    Optional EventSink As clsHookDlg _
                ) As Boolean
    Dim afFlags As Long
    
    m_lApiReturn = 0
    m_lExtendedError = 0
    
    ' Set PRINTDLG flags
    afFlags = flags
    afFlags = afFlags Or (Abs(DisablePageNumbers) * PD_NOPAGENUMS) Or _
              (Abs(DisablePrintToFile) * PD_DISABLEPRINTTOFILE) Or _
              (Abs(DisableSelection) * PD_NOSELECTION) Or _
              (Abs(PrintToFile) * PD_PRINTTOFILE) Or _
              (Abs(Not ShowPrintToFile) * PD_HIDEPRINTTOFILE) Or _
              (Abs(PreventWarning) * PD_NOWARNING) Or _
              (Abs(Collate) * PD_COLLATE) Or _
              PD_USEDEVMODECOPIESANDCOLLATE Or _
              PD_RETURNDC
    If PrintRange = eprPageNumbers Then
        afFlags = afFlags Or PD_PAGENUMS
    ElseIf PrintRange = eprSelection Then
        afFlags = afFlags Or PD_SELECTION
    End If
    ' Mask out unwanted bits
    afFlags = afFlags And Not PD_ENABLEPRINTHOOK
    afFlags = afFlags And Not PD_ENABLEPRINTTEMPLATE
    afFlags = afFlags And Not PD_ENABLESETUPHOOK
    afFlags = afFlags And Not PD_ENABLESETUPTEMPLATE
        
    ' Fill in PRINTDLG structure
    Dim pd As TPRINTDLG
    pd.lStructSize = Len(pd)
    pd.hwndOwner = Owner
    pd.flags = afFlags
    pd.nFromPage = FromPage
    pd.nToPage = ToPage
    pd.nMinPage = 1
    pd.nMaxPage = &HFFFF
'    If (Hook) Then
'        HookedDialog = Me
'        Set m_oEventSink = EventSink
'        If (pd.flags And PD_PRINTSETUP) = PD_PRINTSETUP Then
'            pd.flags = pd.flags Or PD_ENABLESETUPHOOK
'            pd.lpfnSetupHook = lHookAddress(AddressOf PrintSetupHookProc)
'        Else
'            pd.flags = pd.flags Or PD_ENABLEPRINTHOOK
'            pd.lpfnPrintHook = lHookAddress(AddressOf PrintHookProc)
'        End If
'    End If
    
    ' Show Print dialog
    m_lApiReturn = PrintDlg(pd)
    ClearHookedDialog
    Set m_oEventSink = Nothing
    Select Case m_lApiReturn
    Case 1
        VBPrintDlg = True
        ' Return dialog values in parameters
        hdc = pd.hdc
        If (pd.flags And PD_PAGENUMS) Then
            PrintRange = eprPageNumbers
        ElseIf (pd.flags And PD_SELECTION) Then
            PrintRange = eprSelection
        Else
            PrintRange = eprAll
        End If
        FromPage = pd.nFromPage
        ToPage = pd.nToPage
        PrintToFile = (pd.flags And PD_PRINTTOFILE)
        ' Get DEVMODE structure from PRINTDLG
        
        Dim pDevMode As Long
        
        pDevMode = GlobalLock(pd.hDevMode)
        CopyMemory m_dvmode, ByVal pDevMode, Len(m_dvmode)
        GlobalUnlock pd.hDevMode
        
        If (pd.flags And PD_COLLATE) = PD_COLLATE Then
            ' User selected collate option but printer driver
            ' does not support collation.
            ' Collation option must be set from the
            ' PRINTDLG structure:
            Collate = True
            Copies = pd.nCopies
        Else
            ' Print driver supports collation or collation
            ' not switched on.
            ' DEVMODE structure contains Collation and copy
            ' information
            ' Get Copies and Collate settings from DEVMODE structure
            Collate = (m_dvmode.dmCollate = DMCOLLATE_TRUE)
            Copies = m_dvmode.dmCopies
        End If
        
        ' Set default printer properties
        On Error Resume Next
        If Not (Printer Is Nothing) Then
            Printer.Copies = Copies
            Printer.Orientation = m_dvmode.dmOrientation
            Printer.PaperSize = m_dvmode.dmPaperSize
            Printer.PrintQuality = m_dvmode.dmPrintQuality
        End If
        On Error GoTo 0
    Case 0
        ' Cancelled
        VBPrintDlg = False
    Case Else
        ' Extended error:
        m_lExtendedError = CommDlgExtendedError()
        VBPrintDlg = False
    End Select
    
End Function
Friend Property Get DevMode() As DevMode
    DevMode = m_dvmode
End Property

