VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2385
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5595
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   Moveable        =   0   'False
   ScaleHeight     =   159
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   373
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00B1523B&
      FillStyle       =   0  'Solid
      Height          =   2625
      Left            =   -45
      Top             =   -45
      Width           =   900
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Initializing..."
      Height          =   240
      Left            =   975
      TabIndex        =   5
      Top             =   1485
      Width           =   3345
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "This program is protected by US and International copyright laws as described in Help About."
      Height          =   390
      Left            =   960
      TabIndex        =   4
      Top             =   1890
      Width           =   4560
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "© 1999-2000 VB Web Development"
      Height          =   276
      Left            =   972
      TabIndex        =   3
      Top             =   816
      Width           =   4464
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "All Rights Reserved"
      Height          =   255
      Left            =   975
      TabIndex        =   2
      Top             =   1050
      Width           =   2505
   End
   Begin VB.Label lblBuild 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Build 1.5.200"
      Height          =   240
      Left            =   960
      TabIndex        =   1
      Top             =   330
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Developers Pad"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   0
      Top             =   90
      Width           =   2040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   64
      X2              =   362
      Y1              =   95
      Y2              =   95
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   0
      X1              =   64
      X2              =   362
      Y1              =   118
      Y2              =   118
   End
End
Attribute VB_Name = "frmSplash"
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

Private Const SMTO_NORMAL = &H0
Private Const ERROR_ALREADY_EXISTS = 183&
Private Const WM_COPYDATA = &H4A
Private Const HWND_TOPMOST = -1
Private Const HWND_TOP = 0
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByVal lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long

'*** Startup Code ***

Private Sub Form_Load()
    'set mouse pointer
    MousePointer = vbHourglass
    ' Set the version label
    lblBuild.Caption = "Build " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

'called from modStartup
Public Sub Main()
Dim i As Long
Dim gi As Long
On Error GoTo ErrHandler
    Set cDialog = New clsDialog
    ' This procedure has been modified from the initial source code by VB Web
    ' Check if this is the first instance:
    If (WeAreAlone(mcTHISAPPID & "_APPLICATION_MUTEX")) Or GetSetting(REG_KEY, "Settings", "OneProgramInstance", "1") = "0" Then
        ' If it is, then start the app:
        bStartUp = True
        'show and refresh the splash form
        Show
        Refresh
        'call the startup code
        StartUp
        'ensure we are still on top
        BringWindowToTop hWnd
        'parse the command line arguments
        SetStatus "Loading Document..."
        DoEvents
        ParseCommand Command
        'hide the splash form
        Hide
        'we are not starting up!
        bStartUp = False
    Else
        ' There is an existing instance.
        ' First try to find it:
        For gi = 1 To 5
            EnumerateWindows
            If (m_hWndPrevious <> 0) Then Exit For
            'wait...
            For i = 1 To 10000
                DoEvents
            Next
        Next
        ' If we get it:
        If (m_hWndPrevious <> 0) Then
            ' Do we have a command to send, or is the main window hidden?
            If (Command <> "") Or (IsWindowVisible(m_hWndPrevious) = 0) Then
                ' Send.  The app must subclass the WM_COPYDATA message
                ' to get this information:
                Dim tCDS As COPYDATASTRUCT, b() As Byte, lR As Long
                If (Command <> "") Then
                    b = StrConv(Command, vbFromUnicode)
                    tCDS.dwData = 0
                    tCDS.cbData = UBound(b) + 1
                    tCDS.lpData = VarPtr(b(0))
                Else
                    ReDim b(0 To 0) As Byte
                    tCDS.dwData = 0
                    tCDS.cbData = 1
                    tCDS.lpData = VarPtr(b(0))
                End If
                ' Give in if the existing app is not responding:
                lR = SendMessageTimeout(m_hWndPrevious, WM_COPYDATA, 0, tCDS, SMTO_NORMAL, 5000, lR)
            Else
                ' Try to activate the existing window:
                RestoreAndActivate m_hWndPrevious
            End If
        End If
        'end this instance
        EndApp
    End If
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Main"
End Sub

Private Sub StartUp()
On Error GoTo ErrHandler
    'set status bar
    SetStatus "Loading Settings..."
    'fill variables with saved settings
    With vDefault
        'word wrap mode
        .nWordWrap = GetSetting(REG_KEY, "Settings", "WordWrap", "0")
        'default font
        .sFont = GetSetting(REG_KEY, "Settings", "FontName", "Lucida Console")
        'default font size
        .nFontSize = Val(GetSetting(REG_KEY, "Settings", "FontSize", 8))
    End With
    SetStatus "Loading Editor..."
    ' Create a new instance of frmMainForm
    'Set frmMainForm = New frmMainForm
    ' this is where any ocx errors will occur!
    Load frmMainForm
    
    SetStatus "Configuring Editor..."
    'tag window for multiple instances
    TagWindow frmMainForm.hWnd
    ' Setup the main form
    SetupMainform
Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Main.StartUp", "IsMainLoaded: " & IsLoaded("frmMainForm")
    If Err <> 339 Then
        Resume Next
    Else
        'gives us another option when debugging on client machine
        If GetSetting(REG_KEY, "Settings", "AbortOnOCXFile", 1) = "1" Then
            'registration error
            Unload Me
            End
        End If
    End If
End Sub

Private Sub SetupMainform()
On Error GoTo ErrHandler
    Dim intPos As Integer
    Dim bErrOnce As Boolean
    With frmMainForm
        Set cDocuments = New clsDocuments
        
        'restore position of form
        .Top = GetSetting(REG_KEY, "WindowSettings", "MainTop", 1000)
        .Left = GetSetting(REG_KEY, "WindowSettings", "MainLeft", 1000)
        .Width = GetSetting(REG_KEY, "WindowSettings", "MainWidth", 8000)
        .Height = GetSetting(REG_KEY, "WindowSettings", "MainHeight", 6500)
        .WindowState = GetSetting(REG_KEY, "WindowSettings", "MainState", 0)
        'create the toolbars
        SetStatus "Loading Toolbars..."
        LoadToolbars
        'create the menus
        SetStatus "Creating Menus..."
        LoadMenus
        'build the status bar
        SetStatus "Creating Status Bar..."
        BuildStatusBar
        'retreive the history lists
        SetStatus "Getting History List..."
        .InitFindHistory
        .InitFileHistory
        SetStatus "Retreiving Syntax Information..."
        'init the global editor
        Set cGlobalEditor = New clsGlobalEditor
        With cGlobalEditor
            .SyntaxPath = App.Path & "\_syntax\"
            .TemplatePath = App.Path & "\_templates\"
            .ListLanguages
            sFileFilter = "All Files|*.*" & .FileFilter & "|Project Files|" & PROJECT_EXTENSIONS & "|Workspaces|*.dpw|All Files (Force Text)|*.*"
            lProjectFilter = .FilterCount + 2
        End With
        'Set cSyntaxFunctions = New clsSyntaxFile
        'display the form
        SetStatus "Initializing..."
        .Show
        BringWindowToTop hWnd
        DoEvents
        If GetSetting(REG_KEY, "Settings", "SU_InitAddIns", "1") = "1" Then
            SetStatus "Initializing Add-Ins..."
            Set cAddIns = New clsAddIns
            cAddIns.ProcessAddIns
        End If
        'give way for a sec
        DoEvents
        If Val(GetSetting(REG_KEY, "WindowSettings", "frmProject_Pos", 5)) <> 5 And GetSetting(REG_KEY, "Settings", "SU_ShowProject", "1") = "1" Then
            'load the project
            Refresh
            SetStatus "Loading Project..."
            Load frmProject
            frmProject.ShowWindow True
        End If
        .pAttachMessages
    End With
Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Main.SetupMain"
    'try to continue loading...
    Resume Next
End Sub

'*** Toolbars ***
Public Sub LoadToolbars()
On Error GoTo ErrHandler
    With frmMainForm
        'load the standard toolbar
        With .tbrMain(0)
            'assign image list
            .ImageSource = CTBExternalImageList
            .SetImageList frmMainForm.vbalMain, CTBImageListNormal
            .CreateToolbar 16, , True
            'add the toolbar items
            .AddButton "New (Ctrl+N)", plGetIconIndex("New"), , , Empty, CTBDropDown, "New"
            .AddButton "Open (Ctrl+O)", plGetIconIndex("Open"), , , , CTBNormal, "Open"
            .AddButton "Save (Ctrl+S)", plGetIconIndex("Save"), , , Empty, , "Save"
            .AddButton "Save All (Ctrl+Shift+S)", plGetIconIndex("SaveAll"), , , Empty, , "SaveAll"
            .AddButton Empty, -1, , , , CTBSeparator
            .AddButton "Undo (Ctrl+Z)", plGetIconIndex("Undo"), , , Empty, , "Undo"
            .AddButton "Redo (Ctrl+Y)", plGetIconIndex("Redo"), , , Empty, , "Redo"
            .AddButton Empty, -1, , , , CTBSeparator
            .AddButton "Cut (Ctrl+X)", plGetIconIndex("Cut"), , , Empty, , "Cut"
            .AddButton "Copy (Ctrl+C)", plGetIconIndex("Copy"), , , Empty, , "Copy"
            .AddButton "Paste (Ctrl+V)", plGetIconIndex("Paste"), , , Empty, , "Paste"
            .AddButton "Append to clipboard", plGetIconIndex("Append"), , , Empty, , "Append"
            .AddButton Empty, -1, , , , CTBSeparator
            .AddButton "Turn on/off line numbering", plGetIconIndex("Numbering"), , , Empty, CTBCheck, "Numbering"
            .AddButton "Show the Project Window", plGetIconIndex("FILE_VBP"), , , Empty, , "ShowProject"
            .AddButton "Preview this page in a browser (F4)", plGetIconIndex("Preview"), , , Empty, , "Preview"
            .AddButton Empty, -1, , , , CTBSeparator
            .AddButton "Print (Ctrl+P)", plGetIconIndex("Print"), , , Empty, , "Print"
            .AddButton Empty, -1, , , , CTBSeparator
            .AddButton "Help (F1)", plGetIconIndex("Help"), , , Empty, , "Help"
        End With
        'set up standard toolbar host
        With .tbhMain(0)
            .BorderStyle = etbhBorderStyleNone
            .Height = frmMainForm.tbrMain(0).ToolbarHeight * Screen.TwipsPerPixelY
            'capture Main toolbar
            .Capture frmMainForm.tbrMain(0)
        End With
        'load the next item in array
        Load .tbrMain(1)
        Load .tbhMain(1)
        'create find toolbar
        With .tbrMain(1)
            'assign image list
            .ImageSource = CTBExternalImageList
            .SetImageList frmMainForm.vbalMain, CTBImageListNormal
            .CreateToolbar 16, , True
            'add the buttons
            .AddButton "Find Next (Ctrl+N)", plGetIconIndex("FindNext"), , , Empty, CTBNormal, "Find Next"
            .AddButton "Find (Ctrl+F)", plGetIconIndex("Find"), , , Empty, CTBNormal, "Find"
        End With
        'setup Find toolbar host
        With .tbhMain(1)
            .BorderStyle = etbhBorderStyleNone
            .Width = 100
            .Height = frmMainForm.tbrMain(1).ToolbarHeight * Screen.TwipsPerPixelY
            'capture the combo
            .Capture frmMainForm.cboFind
            'capture the find toolbar
            .Capture frmMainForm.tbrMain(1)
        End With
        'load next item in array
        Load .tbrMain(2)
        Load .tbhMain(2)
        'setup edit toolbar
        With .tbrMain(2)
            'set image list
            .ImageSource = CTBExternalImageList
            .SetImageList frmMainForm.vbalMain, CTBImageListNormal
            .CreateToolbar 16, , True
            'add buttons
            .AddButton "Indent (Tab)", plGetIconIndex("Indent"), , , Empty, , "Indent"
            .AddButton "Outdent (Shift+Tab)", plGetIconIndex("Outdent"), , , Empty, , "Outdent"
            .AddButton Empty, -1, , , , CTBSeparator
            .AddButton "Comment Block", plGetIconIndex("Comment"), , , Empty, , "Comment"
            .AddButton "Uncomment Block", plGetIconIndex("Uncomment"), , , Empty, , "Uncomment"
            .AddButton Empty, -1, , , , CTBSeparator
            .AddButton LoadResString(1283) & " (Ctrl+B)", plGetIconIndex("LINE_B"), , , Empty, , "PreviousLine"
            .AddButton LoadResString(1282) & " (Ctrl+Shift+B)", plGetIconIndex("LINE_F"), , , Empty, , "NextLine"
            .AddButton Empty, -1, , , , CTBSeparator
            .AddButton "Quick Tag (Ctrl+I)", plGetIconIndex("EDITTAG"), , , Empty, , "InsertTag"
            .AddButton Empty, -1, , , , CTBSeparator
            .AddButton LoadResString(1302), plGetIconIndex("BOOKMARK"), , , Empty, , "ToggleFlag"
            .AddButton LoadResString(1303), plGetIconIndex("NEXT_BOOKMARK"), , , Empty, , "NextFlag"
            .AddButton LoadResString(1304), plGetIconIndex("LAST_BOOKMARK"), , , Empty, , "LastFlag"
            .AddButton LoadResString(1305), plGetIconIndex("CLEAR_BOOKMARK"), , , Empty, , "ClearFlags"
        End With
        'setup edit toolbar host
        With .tbhMain(2)
            .BorderStyle = etbhBorderStyleNone
            .Width = 100
            .Height = frmMainForm.tbrMain(2).ToolbarHeight * Screen.TwipsPerPixelY
            .Capture frmMainForm.tbrMain(2)
        End With
        'load next item in array
        Load .tbrMain(3)
        Load .tbhMain(3)
        'setup edit toolbar
        With .tbrMain(3)
            'set image list
            .ImageSource = CTBExternalImageList
            .SetImageList frmMainForm.vbalMain, CTBImageListNormal
            .CreateToolbar 16, , True
            'don't add any buttons
            .AddButton " ", -1, , , "", , "DUMMY"
            
        End With
        'setup edit toolbar host
        With .tbhMain(3)
            .BorderStyle = etbhBorderStyleNone
            .Width = 100
            .Height = frmMainForm.tbrMain(3).ToolbarHeight * Screen.TwipsPerPixelY
            .Capture frmMainForm.tbrMain(3)
            frmMainForm.tbrMain(3).ButtonVisible("DUMMY") = False
        End With
        
        'setup rebar
        With .rbrMain
             .Position = erbPositionTop
             .CreateRebar frmMainForm.picHolder.hWnd
             'capture the toolbars
             .AddBandByHwnd frmMainForm.tbhMain(0).hWnd, , , , "Main"
             .AddBandByHwnd frmMainForm.tbhMain(3).hWnd, , False, , "AddIn"
             .AddBandByHwnd frmMainForm.tbhMain(1).hWnd, , , , "Find"
             .AddBandByHwnd frmMainForm.tbhMain(2).hWnd, , False, , "Edit"
             '.AutoSize

             'set minimum width for find rebar
             If frmMainForm.Width >= 7000 Then
                .BandChildMinWidth(0) = 400 '400
                'ok
             End If
             If frmMainForm.Width >= 2445 Then
                .BandChildMinWidth(2) = 163
                'ok
            End If
             '.BandChildMinWidth(3) = 187
             'minimize it
            ' .BandMinimise 2
             '.BandMaximise 3
             'show/hide the rebars
             .BandVisible(0) = Abs(GetSetting(REG_KEY, "Settings", "ShowMainTB", "1"))
             .BandVisible(2) = Abs(GetSetting(REG_KEY, "Settings", "ShowFindTB", "1"))
             .BandVisible(3) = Abs(GetSetting(REG_KEY, "Settings", "ShowCodingTB", "1"))
             .BandVisible(1) = Abs(GetSetting(REG_KEY, "Settings", "ShowAddInTB", "1"))
         End With
    End With
Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "DevPad.Splash:LoadToolbars"
    'try to continue loading...
    Resume Next
End Sub

'*** Menus ***
Private Sub LoadMenus()
On Error GoTo ErrHandler
    Dim iP(0 To 14)     As Long
    Dim iFileNum        As Integer
    Dim sLine           As String
    Dim sName           As String
    Dim sKey            As String
    Dim sValue          As String
    Dim sIcon           As String
    Dim sClassName      As String
    Dim lPos            As Long
    Dim bShow           As Boolean
    Dim lParentIndex    As Long
    Dim lIndex          As Long

    With frmMainForm.ctlPopMenu
        'associate the image list:
        .ImageList = frmMainForm.vbalMain.hIml
        .TickIconIndex = plGetIconIndex("Tick")
        .HighlightCheckedItems = True
        'parse through the VB designed menu and sub class the items:
        .SubClassMenu frmMainForm
        '*** File Menu ***
        iP(0) = .MenuIndex("mnuFile")
        .AddItem LoadResString(102) & Chr$(vbKeyTab) & "Ctrl+N", "FileNew", , , iP(0), plGetIconIndex("NEWTEXT")
        .AddItem LoadResString(111) & Chr$(vbKeyTab) & "Ctrl+O", "FileOpen", , , iP(0), plGetIconIndex("Open")
        .AddItem LoadResString(1279) & Chr$(vbKeyTab) & "Ctrl+L", "FileOpenLinkedFile", , , iP(0), plGetIconIndex("JUMP")
        .AddItem "-", "FileSep10", , , iP(0)
        .AddItem LoadResString(112) & Chr$(vbKeyTab) & "Ctrl+S", "FileSave", , , iP(0), plGetIconIndex("Save")
        .AddItem LoadResString(113), "FileSaveAs", , , iP(0)
        .AddItem LoadResString(115) & Chr$(vbKeyTab) & "Ctrl+Shift+S", "FileSaveAll", , , iP(0), plGetIconIndex("SaveAll")
        .AddItem LoadResString(116), "FileRevert", , , iP(0)
        .AddItem "-", "FileSep2", , , iP(0)
        .AddItem LoadResString(1296), "FileSaveWorkspace", , , iP(0)
        .AddItem LoadResString(1297), "FileSaveWorkspaceAs", , , iP(0)
        .AddItem LoadResString(1299), "FileCloseWorkspace", , , iP(0)
        .AddItem "-", "FileSep4", , , iP(0)
        .AddItem LoadResString(117) & Chr$(vbKeyTab) & "Ctrl+F4", "FileClose", , , iP(0)
        .AddItem LoadResString(118) & Chr$(vbKeyTab) & "Ctrl+Shift+Q", "FileCloseAll", , , iP(0)
        .AddItem "-", "FileSep6", , , iP(0)
        .AddItem LoadResString(1056), "FileProperties", , , iP(0)
        .AddItem "-", "FileSep7", , , iP(0)
        .AddItem LoadResString(119), "FilePrintSetup", , , iP(0)
        .AddItem LoadResString(120) & Chr$(vbKeyTab) & "Ctrl+P", "FilePrint", , , iP(0), plGetIconIndex("Print")
        .AddItem "-", "FileSep8", , , iP(0)
        
        ' both the recent file and recent project menus
        ' are created in frmMainForm.LoadRecentFiles
        .AddItem LoadResString(200), "FileRecentFile", , , iP(0)
        .AddItem LoadResString(201), "FileRecentProject", , , iP(0)
        .AddItem LoadResString(1298), "FileRecentWorkspace", , , iP(0)
        .AddItem "-", "FileSep9", , , iP(0)
        .AddItem LoadResString(121), "FileExit", , , iP(0)
        
        '*** Project Menu ***
        iP(4) = .MenuIndex("mnuProject")
            .AddItem LoadResString(123), "ProjectNew", , , iP(4)
            .AddItem "-", "ProjectSep1", , , iP(4)
            .AddItem LoadResString(124), "ProjectSave", , , iP(4), plGetIconIndex("Save")
            .AddItem LoadResString(125), "ProjectSaveAs", , , iP(4)
            .AddItem "-", "ProjectSep2", , , iP(4)
            .AddItem LoadResString(126), "ProjectClose", , , iP(4)
            .AddItem "-", "ProjectSep3", , , iP(4)
            iP(12) = .AddItem(LoadResString(127), "ProjectAdd", , , iP(4))
                .AddItem LoadResString(128), "ProjectAddFile", , , iP(12), 0
                .AddItem LoadResString(129), "ProjectAddCurrentFile", , , iP(12)
                .AddItem LoadResString(130), "ProjectAddAllFiles", , , iP(12)
                .AddItem LoadResString(131), "ProjectAddFolder", , , iP(12), plGetIconIndex("FOLDERSHORTCUT")
                .AddItem LoadResString(132), "ProjectAddURL", , , iP(12), plGetIconIndex("Hyperlink")
                .AddItem LoadResString(133), "ProjectAddEmail", , , iP(12), plGetIconIndex("Email")
                .AddItem LoadResString(134), "ProjectAddLiveFolder", , , iP(12), plGetIconIndex("LIVEFOLDER")
                .AddItem "-", "ProjectSep4", , , iP(4)
                .AddItem LoadResString(135), "ProjectNewFolder", , , iP(4)
                .AddItem LoadResString(136), "ProjectRemoveItem", , , iP(4), plGetIconIndex("Delete")
                .AddItem "-", "ProjectSep5", , , iP(4)
                .AddItem LoadResString(140), "ProjectProperties", , , iP(4), plGetIconIndex("Properties")

        '*** Edit Menu ***
        iP(5) = .MenuIndex("mnuEdit")
            .AddItem LoadResString(142) & Chr$(vbKeyTab) & "Ctrl+Z", "EditUndo", , , iP(5), plGetIconIndex("Undo")
            .AddItem LoadResString(143) & Chr$(vbKeyTab) & "Ctrl+Shift+Z", "EditRedo", , , iP(5), plGetIconIndex("Redo")
            .AddItem "-", "EditSep1", , , iP(5)
            .AddItem LoadResString(144) & Chr$(vbKeyTab) & "Ctrl+F", "EditFind", , , iP(5), plGetIconIndex("Find")
            .AddItem LoadResString(145) & Chr$(vbKeyTab) & "F3", "EditFindNext", , , iP(5), plGetIconIndex("FindNext")
            .AddItem LoadResString(146) & Chr$(vbKeyTab) & "Ctrl+H", "EditFindReplace", , , iP(5)
            .AddItem LoadResString(1294), "EditFindInFiles", , , iP(5), plGetIconIndex("SEARCH")
            .AddItem "-", "EditSep2", , , iP(5)
            .AddItem LoadResString(147) & Chr$(vbKeyTab) & "Ctrl+X", "EditCut", , , iP(5), plGetIconIndex("Cut")
            .AddItem LoadResString(148) & Chr$(vbKeyTab) & "Ctrl+C", "EditCopy", , , iP(5), plGetIconIndex("Copy")
            .AddItem LoadResString(149) & Chr$(vbKeyTab) & "Ctrl+V", "EditPaste", , , iP(5), plGetIconIndex("Paste")
            .AddItem LoadResString(150), "EditAppend", , , iP(5), plGetIconIndex("Append")
            .AddItem "-", "EditSep3", , , iP(5)
            .AddItem LoadResString(151) & Chr$(vbKeyTab) & "Ctrl+A", "EditSelectAll", , , iP(5)
            
            .AddItem "-", "EditSep4", , , iP(5)
            
            iP(6) = .AddItem(LoadResString(1361), "EditLine", , , iP(5))
                .AddItem LoadResString(1362) & Chr$(vbKeyTab) & "Ctrl+Y", "EditDeleteLine", , , iP(6)
                .AddItem LoadResString(152) & Chr$(vbKeyTab) & "Ctrl+J", "EditGoto", , , iP(6)
                .AddItem "-", "EditSep8", , , iP(6)
                .AddItem LoadResString(1283) & Chr$(vbKeyTab) & "Ctrl+B", "EditLastLine", , , iP(6), plGetIconIndex("LINE_B")
                .AddItem LoadResString(1282) & Chr$(vbKeyTab) & "Ctrl+Shift+B", "EditNextLine", , , iP(6), plGetIconIndex("LINE_F")
                
            iP(6) = .AddItem(LoadResString(1353), "EditSelection", , , iP(5))
                .AddItem LoadResString(1354) & Chr$(vbKeyTab) & "Tab", "EditIndent", , , iP(6), plGetIconIndex("INDENT")
                .AddItem LoadResString(1355) & Chr$(vbKeyTab) & "Shift+Tab", "EditOutdent", , , iP(6), plGetIconIndex("OUTDENT")
                .AddItem "-", "EditSep6", , , iP(6)
                .AddItem LoadResString(1356), "EditComment", , , iP(6), plGetIconIndex("COMMENT")
                .AddItem LoadResString(1357), "EditUncomment", , , iP(6), plGetIconIndex("UNCOMMENT")
                
            .AddItem "-", "EditSep5", , , iP(5)
                
            iP(6) = .AddItem(LoadResString(1364), "EditConvert", , , iP(5))
                .AddItem LoadResString(1358), "EditMakeUppercase", , , iP(6), plGetIconIndex("UPPERCASE")
                .AddItem LoadResString(1359), "EditMakeLowercase", , , iP(6), plGetIconIndex("LOWERCASE")
                .AddItem "-", "EditSep7", , , iP(6)
                .AddItem LoadResString(1360), "EditStringEncode", , , iP(6)
                .AddItem LoadResString(1363), "EditStringDecode", , , iP(6)
                
            iP(6) = .AddItem(LoadResString(1306), "EditBookmarks", , , iP(5))
                .AddItem LoadResString(1302) & Chr$(vbKeyTab) & "Ctrl+F2", "EditToggleBookmark", , , iP(6), plGetIconIndex("BOOKMARK")
                .AddItem LoadResString(1303) & Chr$(vbKeyTab) & "F2", "EditNextBookmark", , , iP(6), plGetIconIndex("NEXT_BOOKMARK")
                .AddItem LoadResString(1304) & Chr$(vbKeyTab) & "Shift+F2", "EditPreviousBookmark", , , iP(6), plGetIconIndex("LAST_BOOKMARK")
                .AddItem LoadResString(1305), "EditClearBookmarks", , , iP(6), plGetIconIndex("CLEAR_BOOKMARK")
        '*** View Menu ***
        iP(6) = .MenuIndex("mnuView")
            iP(13) = .AddItem(LoadResString(1226), "ViewTools", , , iP(6))
                .AddItem LoadResString(154), "ViewMainTB", , , iP(13), , frmMainForm.rbrMain.BandVisible(0)
                .AddItem LoadResString(1011), "ViewFindTB", , , iP(13), , frmMainForm.rbrMain.BandVisible(1)
                .AddItem LoadResString(155), "ViewCodingTB", , , iP(13), , frmMainForm.rbrMain.BandVisible(2)
                .AddItem LoadResString(1286), "ViewAddInTB", , , iP(13), , frmMainForm.rbrMain.BandVisible(3)
            .AddItem "-", "ViewSep1", , , iP(13)
            .AddItem LoadResString(158), "ViewTBTop", , , iP(13), , True
            .AddItem LoadResString(159), "ViewTBBottom", , , iP(13), , False
            .AddItem LoadResString(156), "ViewStatus", , , iP(6), , True
            .AddItem "-", "ViewSep2", , , iP(6)
            .AddItem LoadResString(157), "ViewProject", , , iP(6), plGetIconIndex("FILE_VBP")
            .AddItem LoadResString(144), "ViewFind", , , iP(6), plGetIconIndex("Find")
            .AddItem LoadResString(202), "ViewFindResults", , , iP(6)
            .AddItem "-", "ViewSep3", , , iP(6)
            iP(14) = .AddItem(LoadResString(1136), "ViewWrapping", , , iP(6))
                .AddItem LoadResString(402), "ViewWordWrap", , , iP(14), , IIf(vDefault.nWordWrap = 1, True, False)
                .AddItem LoadResString(403), "ViewNoWordWrap", , , iP(14), , IIf(vDefault.nWordWrap = 0, True, False)
                .AddItem LoadResString(404), "ViewWYSIWYG", , , iP(14), , IIf(vDefault.nWordWrap = 2, True, False)

        '*** Insert Menu ***
        iP(7) = .MenuIndex("mnuInsert")
            .AddItem LoadResString(162) & Chr$(vbKeyTab) & "Ctrl+T", "InsertTime", , , iP(7), plGetIconIndex("Time")
            .AddItem LoadResString(163) & Chr$(vbKeyTab) & "Ctrl+D", "InsertDate", , , iP(7), plGetIconIndex("Date")
            .AddItem "-", "InsertSep1", , , iP(7)
            .AddItem LoadResString(164), "InsertSymbol", , , iP(7)
            .AddItem LoadResString(1281) & Chr$(vbKeyTab) & "Ctrl+I", "InsertTag", , , iP(7), plGetIconIndex("EDITTAG")
            .AddItem "-", "InsertSep2", , , iP(7)
            .AddItem LoadResString(166), "InsertTextFromFile", , , iP(7)

        '*** Tools Menu ***
        iP(9) = .MenuIndex("mnuTools")
            .AddItem LoadResString(205), "ToolsManage", , , iP(9), plGetIconIndex("ADDINS")
            .AddItem LoadResString(1276) & Chr$(vbKeyTab) & "F4", "ToolsDocPreview", , , iP(9), plGetIconIndex("PREVIEW")
            .AddItem "-", "ToolsSep1", , , iP(9)
            .AddItem "-", "ToolsSep2", , , iP(9)
            'and continue with tools menu...
            .AddItem LoadResString(176) & Chr$(vbKeyTab) & "Ctrl+M", "ToolsSearch", , , iP(9), plGetIconIndex("Search")
            .AddItem LoadResString(177) & Chr$(vbKeyTab) & "Ctrl+X", "ToolsExplorer", , , iP(9), plGetIconIndex("Explorer")
            .AddItem LoadResString(178) & Chr$(vbKeyTab) & "Ctrl+G", "ToolsRegEdit", , , iP(9), plGetIconIndex("RegEdit")
            .AddItem LoadResString(1288) & Chr$(vbKeyTab) & "Ctrl+R", "ToolsRun", , , iP(9), plGetIconIndex("RUN")
            .AddItem "-", "ToolsSep3", , , iP(9)
        iP(10) = .AddItem(LoadResString(179), "ToolsVBCommands", , , iP(9), plGetIconIndex("FILE_VBP"))
            .AddItem LoadResString(180), "ToolsCompilePrj", , , iP(10)
            .AddItem LoadResString(181), "ToolsRunPrj", , , iP(10)
            .AddItem LoadResString(182), "ToolsOpenPrj", , , iP(10)
            .AddItem "-", "ToolsSep4", , , iP(9)
            .AddItem LoadResString(185), "ToolsOptions", , , iP(9)

        '*** Window Menu ***
        iP(11) = .MenuIndex("mnuWindow")
            .AddItem LoadResString(188), "WindowCascade", , , iP(11), plGetIconIndex("Cascade")
            .AddItem LoadResString(189), "WindowTileH", , , iP(11), plGetIconIndex("TileHorz")
            .AddItem LoadResString(190), "WindowTileV", , , iP(11), plGetIconIndex("TileVert")
            .AddItem "-", "WindowSep1", , , iP(11)
            .AddItem LoadResString(1228), "WindowMinimize", , , iP(11), plGetIconIndex("MINWIN")
            .AddItem LoadResString(191), "WindowMaximize", , , iP(11), plGetIconIndex("MAXWIN")
            'openw windows listed in frmMainForm.ListWindows
            .AddItem "-", "WindowSep2", , , iP(11)
        
        '*** Help Menu ***
        iP(12) = .MenuIndex("mnuHelp")
            .AddItem LoadResString(193), "HelpContents", , , iP(12), plGetIconIndex("Help")
            .AddItem LoadResString(194), "HelpIndex", , , iP(12)
            .AddItem LoadResString(1365), "HelpReportBug", , , iP(12), plGetIconIndex("EMAIL")
            .AddItem "-", "HelpSep1", , , iP(12)
            .AddItem LoadResString(1301), "HelpReadme", , , iP(12), plGetIconIndex("FILE_HTM")

            iP(13) = .AddItem(LoadResString(1231), "HelpWeb", , , iP(12), plGetIconIndex("INTERNET"))
                .AddItem LoadResString(1233), "HelpForum", , , iP(13), plGetIconIndex("FILE_HTM")
                .AddItem LoadResString(1234), "HelpKB", , , iP(13), plGetIconIndex("FILE_HTM")
                .AddItem "-", "HelpWebSep", , , iP(13)
                .AddItem LoadResString(1232), "HelpDevPad", , , iP(13), plGetIconIndex("FILE_HTM")
                .AddItem LoadResString(199), "HelpVBWeb", , , iP(13), plGetIconIndex("FILE_HTM")
            .AddItem "-", "HelpSep2", , , iP(12)
            .AddItem LoadResString(198), "HelpAbout", , , iP(12)
            
        'Menus built!
    End With
Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Main.LoadMenus"
End Sub


'*** Status Bar ***
Private Sub BuildStatusBar()
On Error GoTo ErrHandler
    Set frmMainForm.cStatusBar = New clsStatusBar
    With frmMainForm.cStatusBar
        .Create frmMainForm.picStatus
      '  frmMainForm.picStatus.BorderStyle = 0
      '  frmMainForm.picStatus.ZOrder
        '.ImageList = ilsIcons
        .AddPanel estbrNoBorders, "Ready", , , True, , , "Status"
        .AddPanel estbrStandard, "Pos 0 Ln 1/1 Col 0", , , , True, , "CharNum"
        .SizeGrip = True
        .SetLeftTopOffsets 0, 2
        frmMainForm.picStatus.Height = .Height
   End With
   Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "StartUp.BuildStatusBar"
End Sub

'*** Instancing ***
Public Function EnumerateWindows() As Boolean
   ' Enumerate top-level windows:
   EnumWindows AddressOf EnumWindowsProc, 0
End Function
Private Function WeAreAlone(ByVal sMutex As String) As Boolean
    ' Don't call Mutex when in VBIDE because it will apply
    ' for the entire VB IDE session, not just the app's
    ' session.
    If InDevelopment Then
        WeAreAlone = Not (App.PrevInstance)
    Else
        ' Ensures we don't run a second instance even
        ' if the first instance is in the start-up phase
        m_hMutex = CreateMutex(ByVal 0&, 1, sMutex)
        If (Err.LastDllError = ERROR_ALREADY_EXISTS) Then
            CloseHandle m_hMutex
        Else
            WeAreAlone = True
        End If
    End If
End Function
 
'*** Other ***
'returns an imagelist index
Private Function plGetIconIndex(ByVal sKey As String) As Long
    If sKey = "-1" Then
        plGetIconIndex = -1
    Else
        On Error Resume Next
        plGetIconIndex = frmMainForm.vbalMain.ItemIndex(UCase$(sKey))
        If plGetIconIndex = 0 Then plGetIconIndex = -1
    End If
End Function
Public Sub SetStatus(sMessage As String)
    'updates the status message
    lblStatus.Caption = sMessage
    lblStatus.Refresh
    DoEvents
End Sub
Private Sub MakeTopMost(hWnd As Long)
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub
