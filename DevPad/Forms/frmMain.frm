VERSION 5.00
Object = "{C4925FC3-1606-11D4-82BB-004005AAE138}#5.2#0"; "VBWIML.OCX"
Object = "{1F5999A2-1D0B-11D4-82CF-004005AAE138}#6.1#0"; "VBWTBA~1.OCX"
Object = "{C93B2CCD-391A-424C-9BF8-49622ED15ACC}#1.0#0"; "vbwPopMenu.ocx"
Begin VB.MDIForm frmMainForm 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Developers Pad"
   ClientHeight    =   5010
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   7905
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   OLEDropMode     =   1  'Manual
   ScrollBars      =   0   'False
   Tag             =   "1"
   Begin VB.PictureBox picStatus 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   527
      TabIndex        =   4
      Top             =   4635
      Width           =   7905
      Begin DevPad.vbwProgressBar ctlProg 
         Height          =   225
         Left            =   1140
         TabIndex        =   5
         Top             =   105
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox picProjectHolder 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3510
      Left            =   0
      ScaleHeight     =   3510
      ScaleWidth      =   525
      TabIndex        =   3
      Top             =   1125
      Visible         =   0   'False
      Width           =   525
   End
   Begin vbwPopMenu.PopMenu ctlPopMenu 
      Left            =   3750
      Top             =   2370
      _ExtentX        =   1058
      _ExtentY        =   1058
      HighlightCheckedItems=   0   'False
      TickIconIndex   =   0
   End
   Begin vbwIml.vbalImageList vbalMain 
      Left            =   2670
      Top             =   2385
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   8
      Size            =   62040
      Images          =   "frmMain.frx":0442
      KeyCount        =   66
      Keys            =   $"frmMain.frx":F6BA
   End
   Begin VB.PictureBox picHolder 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   0
      ScaleHeight     =   1125
      ScaleWidth      =   7905
      TabIndex        =   0
      Top             =   0
      Width           =   7905
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   630
         Left            =   3120
         TabIndex        =   6
         Top             =   270
         Width           =   765
      End
      Begin VB.ComboBox cboFind 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   825
         TabIndex        =   1
         Top             =   645
         Width           =   1695
      End
      Begin vbwTBar.cToolbarHost tbhMain 
         Height          =   495
         Index           =   0
         Left            =   1170
         TabIndex        =   2
         Top             =   0
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
      End
      Begin vbwTBar.cToolbar tbrMain 
         Index           =   0
         Left            =   4020
         Top             =   60
         _ExtentX        =   2090
         _ExtentY        =   873
      End
      Begin vbwTBar.cReBar rbrMain 
         Left            =   0
         Top             =   0
         _ExtentX        =   1931
         _ExtentY        =   767
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
   End
   Begin VB.Menu mnuProject 
      Caption         =   "&Project"
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Insert"
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
   Begin VB.Menu mnuDocPopup 
      Caption         =   "&Doc"
      Visible         =   0   'False
      Begin VB.Menu mnudummy 
         Caption         =   "test"
      End
   End
End
Attribute VB_Name = "frmMainForm"
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
Implements ISubclass

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function SHRunDialog Lib "shell32" Alias "#61" _
                            (ByVal hOwner As Long, _
                             ByVal Unknown1 As Long, _
                             ByVal Unknown2 As Long, _
                             ByVal szTitle As String, _
                             ByVal szPrompt As String, _
                             ByVal uFlags As Long) As Long
Private Const WM_MDIICONARRANGE = &H228
Private Const WM_MDIMAXIMIZE = &H225
Private Const WM_COPYDATA = &H4A
Private Const CB_SHOWDROPDOWN = &H14F

Public cFindHistory         As clsHistory
Public cReplaceHistory      As clsHistory
'Public cFileHistory         As clsHistory
Public cStatusBar           As clsStatusBar

Public bCancelClose         As Boolean ' Cancel Exit
Private cFileHistory(1 To 3) As clsHistory
'Private cLastOpenHistory    As clsHistory
'Private cProjectHistory     As clsHistory
Private cPrint              As clsPrint
Private cPreview            As clsPreview
Private cFindCombo          As clsFlatCombo
Private cText               As clsText

Private bDropping           As Boolean
Private bFindFocus          As Boolean
Private lDocumentCount      As Long 'count of new documents loaded...
Private m_bDocsOpen         As Boolean

Public Sub SetDocStatus()
    'check if a document is open, and
    'update the toolbars/menus
    If DocOpen() = False Then
        ' no more documents open (apart from this one)
        SetNoDocsOpen
    Else
        ' another document open
        SetDocsOpen
    End If
End Sub
Private Sub cboFind_Change()
    'Enable FindNext button on tb if we have text to look for
    EnableTB "Find Next", (cboFind.Text <> ""), 1
    'update the find text in frmFind
    If IsLoaded("frmFind") Then frmFind.txtFind = cboFind.Text
End Sub
Private Sub cboFind_Click()
    'an item from the list has been selected
    'simulate change
    cboFind_Change
End Sub

Private Sub cboFind_DropDown()
    If bFindFocus = False Then bDropping = True
End Sub

Private Sub cboFind_GotFocus()
Static bIgnore As Boolean
    bFindFocus = True
    If bIgnore Then Exit Sub
    'select all text
    cboFind.SelStart = 0
    cboFind.SelLength = Len(cboFind.Text)
    'for some reason, mdi child does not lose focus when cboFind is activated...
    'so we need send the focus to the tb host, and then return it here...
    '(without getting ourselves in a loop)
    bIgnore = True
    SendMessage tbhMain(1).hWnd, WM_SETFOCUS, 0&, 0&
    SendMessage cboFind.hWnd, WM_SETFOCUS, 0&, 0&
    'if we don't give way, the bIgnore flag is set to false
    'before the SetFocus events have a chance to execute,
    'and we get into a loop!
    DoEvents
    If bDropping Then
        SendMessage cboFind.hWnd, CB_SHOWDROPDOWN, True, 0
    End If
    bDropping = False
    bIgnore = False
End Sub

Private Sub cboFind_LostFocus()
    bFindFocus = False
End Sub

Private Sub cboFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        'if the FindNext tb button is enabled, run findnext
        If tbrMain(1).ButtonEnabled("Find Next") = True Then FindNext
        'ensure combo still has focus
        cboFind.SetFocus
    End If
End Sub
Private Sub FindNext()
    'load find form if it isn't already
    If IsLoaded("frmFind") = False Then
        Load frmFind
        'update the find text
        frmFind.txtFind.Text = cboFind.Text
    End If
    If cboFind.Text = "" Then
        frmFind.ShowFind (False)
    Else
        'simulate Find click on find form...
        frmFind.cmdFind_Click
    End If
End Sub




Private Sub Command1_Click()
'    frmOutput.Show
'    frmOutput.RunCommand ""
'    frmTaskList.Show
    frmBrowser.Show
End Sub

'*** Menu Handling ***

Private Sub ctlPopMenu_Click(ItemNumber As Long)
On Error GoTo ErrHandler

    Dim i           As Long
    Dim sValue      As String
    Dim sRootKey    As String
    Dim sKey        As String
    
    'no pressing keys as devpad loads!
    If bStartUp Then Exit Sub
    'get the menu item's key
    sKey = ctlPopMenu.MenuKey(ItemNumber)
    Select Case sKey
    '*** file menu ***
    Case "FileNew"
        'display the open dialog, with New tab selected
        OpenFile , 0
    Case "FileOpen"
        'display the open dialog
        OpenFile
    Case "FileOpenLinkedFile"
        OpenLinkedFile
    Case "FileSave"
        'save the active form's document
        ActiveDoc.Save
    Case "FileSaveAs"
        'save as...
        ActiveDoc.SaveAs
    Case "FileSaveAll"
        'save all the open documents
        SaveAll
    Case "FileClose"
        'close the current form
        ActiveDoc.Close
    Case "FileCloseAll"
        'close all the windows!
        CloseAll
    Case "FileRevert"
        'revert to the saved version
        Revert
    Case "FilePrintSetup"
        'init the class if it hasn't been created yet
        If cPrint Is Nothing Then Set cPrint = New clsPrint
        'Show the Print setup dialog
        cPrint.ShowPrintDialog &H40
    Case "FilePrint"
        'init the class if it hasn't been created yet
        If cPrint Is Nothing Then Set cPrint = New clsPrint
        'print the page...
        cPrint.PrintPage
    Case "FileProperties"
        If ActiveDoc.FileName <> "" Then
            'display properties for the current file
            ShellExFunc "properties", ActiveDoc.FileName, hWnd
        End If
    Case "FileSaveWorkspace"
        cWorkspace.Save
    Case "FileCloseWorkspace"
        cWorkspace.CloseWorkspace
    Case "FileSaveWorkspaceAs"
        cWorkspace.SaveAs
    Case "FileExit"
        'exit Developers Pad
        Unload Me
    
    '*** project menu ***
    Case "ProjectNew"
        'create a new project
        frmProject.NewProject
    Case "ProjectSave"
        'save the project
        frmProject.SaveProject
    Case "ProjectSaveAs"
        'save project as...
        frmProject.SaveProject True
    Case "ProjectClose"
        'close the project
        frmProject.ClearProject True
    Case "ProjectAddFile"
        'add a file..
        frmProject.AddFiles
    Case "ProjectAddLiveFolder"
        frmProject.AddLiveFolder
    Case "ProjectAddCurrentFile"
        'add the current file to the project
        frmProject.NewItem ActiveDoc.FileName, ActiveDoc.DocumentCaption
    Case "ProjectAddURL"
        'add a url
        frmProject.AddInternet True
    Case "ProjectAddEmail"
        'add an emil address
        frmProject.AddInternet False
    Case "ProjectAddFolder"
        'add a folder shortcut
        frmProject.AddFolder
    Case "ProjectNewFolder"
        'add a new folder
        frmProject.NewFolder
    Case "ProjectRemoveItem"
        'delete the selected item
        frmProject.DeleteItem
    Case "ProjectAddAllFiles"
        'add all the open files to the project
        frmProject.AddAllFiles
    Case "ProjectProperties"
        'display the properties for this project
        frmProject.ShowProperties
        
    '*** edit menu ***
    Case "EditUndo"
        'undo
        ActiveDoc.Undo
    Case "EditRedo"
        'redo
        ActiveDoc.Redo
    Case "EditFind", "ViewFind"
        'display the find dialog
        frmFind.ShowFind (False)
    Case "EditFindNext"
        'do a findnext
        FindNext
    Case "EditFindReplace"
        'show find dialog with replace part
        frmFind.ShowFind (True)
    Case "EditCut"
        ActiveDoc.Cut
    Case "EditCopy"
        ActiveDoc.Copy
    Case "EditPaste"
        ActiveDoc.Paste
    Case "EditAppend"
        ActiveDoc.Append
    Case "EditSelectAll"
        ActiveDoc.SelectAll
    Case "EditGoto"
        'display the Goto dialog
        LoadShow frmGoto, vbModeless
    Case "EditLastLine"
        'go to last pos
        ActiveDoc.PreviousLine
    Case "EditNextLine"
        'go to next pos
        ActiveDoc.NextLine
    Case "EditDeleteLine"
        ActiveDoc.DeleteLine
    Case "EditMakeUppercase"
        ActiveDoc.ChangeCase True
    Case "EditMakeLowercase"
        ActiveDoc.ChangeCase False
    Case "EditStringEncode"
        ActiveDoc.StringEncode True
    Case "EditStringDecode"
        ActiveDoc.StringEncode False
    Case "EditFindInFiles"
        LoadShow frmFindFiles, vbModeless
    Case "EditToggleBookmark"
        ActiveDoc.ToggleFlag
    Case "EditNextBookmark"
        pNextFlag
    Case "EditPreviousBookmark"
        pLastFlag
    Case "EditClearBookmarks"
        pClearFlags
    Case "EditUncomment"
        ActiveDoc.UncommentBlock
    Case "EditComment"
        ActiveDoc.CommentBlock
    Case "EditIndent"
        ActiveDoc.Indent
    Case "EditOutdent"
        ActiveDoc.Outdent
    '*** view menu ***
    Case "ViewMainTB"
        'view the main rebar
        ShowHideTB ctlPopMenu.Checked(ItemNumber), 0, ItemNumber
    Case "ViewFindTB"
        'view the find rebar
        ShowHideTB ctlPopMenu.Checked(ItemNumber), 2, ItemNumber
    Case "ViewCodingTB"
        'view the find rebar
        ShowHideTB ctlPopMenu.Checked(ItemNumber), 3, ItemNumber
    Case "ViewAddInTB"
        ShowHideTB ctlPopMenu.Checked(ItemNumber), 1, ItemNumber
    Case "ViewStatus"
        'show/hide the status bar
        ctlPopMenu.Checked(ItemNumber) = Not ctlPopMenu.Checked(ItemNumber)
        picStatus.Visible = ctlPopMenu.Checked(ItemNumber)
    Case "ViewProject"
        'display the project window
        ShowProjectWindow
    Case "ViewTBTop"
        'make the toolbar appear at the top...
        picHolder.Align = vbAlignTop
        'check the menu item
        ctlPopMenu.Checked(ItemNumber) = True
        'uncheck the other menu item
        ctlPopMenu.Checked(ctlPopMenu.MenuIndex("ViewTBBottom")) = False
    Case "ViewTBBottom"
        'make the toolbar appear at the bottom...
        'don't ask me who would want it like that!
        picHolder.Align = vbAlignBottom
        ctlPopMenu.Checked(ItemNumber) = True
        ctlPopMenu.Checked(ctlPopMenu.MenuIndex("ViewTBTop")) = False
    Case "ViewWordWrap"
        'word wrap
        SetWordWrap 1
    Case "ViewNoWordWrap"
        'no wrap
        SetWordWrap 0
    Case "ViewWYSIWYG"
        'WYSIWYG - What you see is what you get (when printing)
        SetWordWrap 2
    Case "ViewFindResults"
        'display find results window
        LoadShow frmFindResults
    '*** insert menu ***
    Case "InsertTime"
        'insert the current time
        ActiveDoc.SelText = Time
    Case "InsertDate"
        'insert the current date
        ActiveDoc.SelText = DateValue(Now)
    Case "InsertSymbol"
        'display the project window...
        ShowProjectWindow
        'and ensure \symbols folder is visible
        '(unless someone has deleted it!)
        frmProject.InsertSymbols
    Case "InsertTag"
        'insert a HTML Tag
        ActiveDoc.InsertTag
    Case "InsertTextFromFile"
        'init class if needed
        If cText Is Nothing Then Set cText = New clsText
        'load the text from a file
        cText.InsertTextFromFile
    Case "ToolsDocPreview"
        If cPreview Is Nothing Then Set cPreview = New clsPreview
        cPreview.PreviewDocument ActiveDoc
    Case "ToolsManage"
        'init Add-in DLL
        InitAddIns
        'display add-in dialog
        LoadShow frmAddIns, vbModeless
    Case "ToolsExplorer"
        'if we have an active doc, then get its path,
        'and use this as the initial explorer path
        If DocOpen Then sValue = GetFolder(ActiveDoc.FileName)
        'open explorer using ShellExecute
        ShellFunc sValue, , "explore"
    Case "ToolsRun"
        SHRunDialog hWnd, 0, 0, vbNullString, vbNullString, 0
    Case "ToolsRegEdit"
        'run regedit.exe
        ShellFunc "RegEdit", vbNormalFocus
    Case "ToolsSearch"
        'if we have an active doc, then get its path,
        'and use this as the initial search path
        If DocOpen Then sValue = GetFolder(ActiveDoc.FileName)
        'open search using ShellExecuteEx
        ShellExFunc "find", sValue, hWnd
    Case "ToolsCompilePrj", "ToolsRunPrj", "ToolsOpenPrj"
        'get the current VB path, if we haven't already
        GetVBPath
        'select the correct option
        Select Case sKey
        Case "ToolsCompilePrj"
            sValue = " /make"
        Case "ToolsRunPrj"
            sValue = " /run"
        'Case "ToolsOpenPrj"
            'no option needed
        End Select
        'run VB, with the correct option
        Shell Chr$(34) & sVBPath & Chr$(34) & " " & Chr$(34) & ActiveDoc.FileName & Chr$(34) & sValue, vbNormalFocus
    Case "ToolsOptions"
        'display the options dialog
        LoadShow frmOptions, vbModeless
    
    '*** window menu ***
    Case "WindowCascade"
        'cascade windows...
        Arrange vbCascade
    Case "WindowTileH"
        'tile windows horizontally
        Arrange vbTileHorizontal
    Case "WindowTileV"
        'tile windows vertically
        Arrange vbTileVertical
    Case "WindowMaximize"
        'maximize windows
        If Not ActiveDoc Is Nothing Then SendMessageLong GetMDIClienthWnd(), WM_MDIMAXIMIZE, ActiveDoc.hWnd, 0
    Case "WindowMinimize"
        'minimize all the windows
        For i = 1 To cDocuments.Count
            cDocuments.Item(i).WindowState = vbMinimized
        Next i
        SendMessageLong GetMDIClienthWnd(), WM_MDIICONARRANGE, 0, 0
        
    '*** help menu ***
    Case "HelpAbout"
        'display the about dialog
        LoadShow frmAbout, vbModal
    Case "HelpForum"
        'open the support page
        ShellFunc "http://www.developerspad.com/forum/"
    Case "HelpKB"
        'open faq page
        ShellFunc "http://www.developerspad.com/kb/"
    Case "HelpDevPad"
        'open devpad home page
        ShellFunc "http://www.developerspad.com/"
    Case "HelpReportBug"
        ShellFunc "http://www.developerspad.com/forum/post.php?forum=2"
    Case "HelpVBWeb"
        'open VB Web home page
        ShellFunc "http://www.vbweb.co.uk/"
    Case "HelpContents"
        'display contents tab of html help file
        cDialog.ShowHelp True
    Case "HelpIndex"
        'display contents tab of html help file
        cDialog.ShowHelp False
    Case "HelpReadme"
        ShellFunc App.Path & "\readme.htm"
    '*** other cases ***
    Case Else
        'we need to do something else...
        'get the root key
        sRootKey = ctlPopMenu.MenuKey(1)
        'don't like it, but that's life...
        'no menu parent item...
        If sKey Like "EditPop*" Then
            'document context menu
            'get its command (excluding the EditPop bit)
            Select Case Right$(sKey, Len(sKey) - 7)
            Case "Cut", "Copy", "Paste", "Append", "Undo", "Redo"
                'simply simulate Edit|Cut etc...
                ctlPopMenu_Click (ctlPopMenu.MenuIndex("Edit" & Right$(sKey, Len(sKey) - 7)))
            Case "Open"
                'simulate File|Open...
                ctlPopMenu_Click (ctlPopMenu.MenuIndex("FileOpen"))
            Case "OpenLinked"
                OpenLinkedFile
            End Select
        ElseIf sKey Like "PrjPop*" Then
            'project popup menu...
            frmProject.PopupMenuHandle ItemNumber
        ElseIf sKey Like "CodePop*" Then
            frmProject.PopupMenuHandle ItemNumber
        ElseIf sKey Like "DocPop*" Then
            frmProject.PopupMenuHandle ItemNumber
        ElseIf sKey Like "mnuNew*" Then
            'get the caption
            sValue = ctlPopMenu.Caption(ItemNumber)
            If sValue = LoadResString(1242) Then '"More..."
                'more... selected. Display open dialog
                OpenFile , 0
            Else
                'load a new document using the specified template
                cDocuments.New True, sValue & ".txt" 'LoadNewDoc , sValue & ".txt"
            End If
        ElseIf sKey Like "FileRecentFile*" Then
            ' recent file
            If cDocuments.LoadFile(cFileHistory(1).Item(ctlPopMenu.ItemData(ItemNumber))) Is Nothing Then
                'failed
                cFileHistory(1).Remove (ctlPopMenu.ItemData(ItemNumber))
                LoadFileHistory 1, "FileRecentFile"
            End If
        ElseIf sKey Like "FileRecentProject*" Then
            ShowProjectWindow
            ' recent project
            ' Call the file open procedure, passing a reference to the selected file name
            frmProject.OpenProject cFileHistory(2).Item(ctlPopMenu.ItemData(ItemNumber))
        ElseIf sKey Like "FileRecentWorkspace*" Then
            cWorkspace.Load cFileHistory(3).Item(ctlPopMenu.ItemData(ItemNumber))
        ElseIf sKey Like "ToolsAddIn*" Then
            'init the Add-in dll
            InitAddIns
            'load the selected add-in
            i = CLng(Right$(sKey, Len(sKey) - 10))
            If (cAddIns.Add(i)) Then cAddIns.Tool(i).ShowDialog
            
        ElseIf sKey Like "WindowItem*" Then
            'activate the document using its ID
            cDocuments.ItemByID(ctlPopMenu.ItemData(ItemNumber)).SetFocus
        Else
            If cAddIns.RaiseMenuClicks(ItemNumber) = False Then
                'menu item not handled
                Debug.Print "Menu Item Unknown: " & ctlPopMenu.MenuKey(ItemNumber)
            End If

        End If
    End Select
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Main.Menus"
End Sub

Private Sub ctlPopMenu_InitPopupMenu(ParentItemNumber As Long)
Dim bEnabled As Boolean
    On Error Resume Next
    'prepares menus before displaying
    Select Case ctlPopMenu.MenuKey(ParentItemNumber)
    Case "mnuView"
        'check the correct items for text wrapping
        If DocOpen Then
            ctlPopMenu.Checked(ctlPopMenu.MenuIndex("ViewWordWrap")) = IIf(ActiveDoc.ViewMode = 1, True, False)
            ctlPopMenu.Checked(ctlPopMenu.MenuIndex("ViewNoWordWrap")) = IIf(ActiveDoc.ViewMode = 0, True, False)
            ctlPopMenu.Checked(ctlPopMenu.MenuIndex("ViewWYSIWYG")) = IIf(ActiveDoc.ViewMode = 2, True, False)
        End If
        ctlPopMenu.Enabled("ViewWYSIWYG") = Not (Printer Is Nothing)
    Case "mnuFile"
        If DocOpen = True Then bEnabled = (ActiveDoc.Saved)
        'set revert menu item
        ctlPopMenu.Enabled("FileRevert") = bEnabled
        ctlPopMenu.Enabled("FileProperties") = bEnabled
        bEnabled = False
        If Not cWorkspace Is Nothing Then bEnabled = (cWorkspace.FileName <> "")
        ctlPopMenu.Enabled("FileCloseWorkspace") = bEnabled
        If DocOpen Then
            bEnabled = (ActiveDoc.CursorFile(False) <> "")
        Else
            bEnabled = False
        End If
        ctlPopMenu.Enabled("FileOpenLinkedFile") = bEnabled
    Case "mnuProject"
        bEnabled = DocOpen
        If bEnabled Then bEnabled = (ActiveDoc.Saved)
        ctlPopMenu.Enabled("ProjectAddCurrentFile") = bEnabled
    Case "mnuWindow"
        'populate the window list
        ListWindows
    End Select
End Sub

Public Sub SetProjectMenu(bEnabled As Boolean)
    If bClosing Then Exit Sub
    'Prepares project menu depending on if the project form
    'is loaded or not, and if it is a VB project
    With ctlPopMenu
        .Enabled("ProjectSave") = bEnabled
        .Enabled("ProjectSaveAs") = bEnabled
        .Enabled("ProjectClose") = bEnabled
        .Enabled("ProjectAddFile") = bEnabled
        .Enabled("ProjectAddLiveFolder") = bEnabled
        .Enabled("ProjectAddAllFiles") = bEnabled
        .Enabled("ProjectRemoveItem") = bEnabled
        
        If DocOpen() Then
            'Enable AddCurrentFile if there is an open document
            .Enabled("ProjectAddCurrentFile") = bEnabled
        Else
            .Enabled("ProjectAddCurrentFile") = False
        End If
        .Enabled("ProjectProperties") = bEnabled
'        If IsLoaded("frmProject") Then
'            bEnabled = IIf(bEnabled, frmProject.DevPadProject, False)
'        End If
        .Enabled("ProjectAddURL") = bEnabled
        .Enabled("ProjectAddFolder") = bEnabled
        .Enabled("ProjectAddEmail") = bEnabled
        .Enabled("ProjectNewFolder") = bEnabled
        
    End With
End Sub


'*** Toolbar Handling ***

Private Sub tbrMain_ButtonClick(Index As Integer, ByVal lButton As Long)
    Dim sKey As String
    Dim lItem As Long
    'handle toolbar clicks
    On Error GoTo ErrHandler
    sKey = tbrMain(Index).ButtonKey(lButton)

        Select Case sKey
    '*** actions ***
    Case "New"
        'load a new document
        cDocuments.New
    Case "Open"
        'open a file
        OpenFile
    Case "Save"
        'save the active file
        ActiveDoc.Save
    Case "SaveAll"
        'save all the open files
        SaveAll
    
    '*** edit ***
    Case "Cut"
        ActiveDoc.Cut
    Case "Copy"
        ActiveDoc.Copy
    Case "Paste"
        ActiveDoc.Paste
    Case "Append"
        ActiveDoc.Append
    Case "Undo"
        'undo...
        ActiveDoc.Undo
    Case "Redo"
        'redo...
        ActiveDoc.Redo
    
    '*** other ***
    Case "Print"
        If cPrint Is Nothing Then Set cPrint = New clsPrint
        cPrint.PrintPage
    Case "ShowProject"
        ShowProjectWindow
        frmProject.SelectTab "Project"
    Case "Preview"
        If cPreview Is Nothing Then Set cPreview = New clsPreview
        cPreview.PreviewDocument ActiveDoc
    Case "Numbering"
        If ActiveDoc.IsRTF And ActiveDoc.ShowLines = False Then
            cDialog.ShowWarning "You cannot have line numbering for RTF files", "tbrMain_ButtonClick"
        Else
            ActiveDoc.ShowLines = Not ActiveDoc.ShowLines
            tbrMain(0).ButtonPressed("Numbering") = ActiveDoc.ShowLines
            'ActiveDoc.Form_Resize
            SaveSetting REG_KEY, "Settings", "ShowLines", tbrMain(0).ButtonPressed("Numbering")
        End If
        
    Case "Help"
        'ShowHelp True
        
    '*** Find Toolbar ***
    Case "Find Next"
        FindNext
    Case "Find"
        frmFind.ShowFind False
    '*** Edit Toolbar ***
    Case "Indent"
        'indent selected text
        ActiveDoc.Indent
    Case "Outdent"
        'outdent selected text
        ActiveDoc.Outdent
    Case "Comment"
        'comment out selected text
        ActiveDoc.CommentBlock
    Case "Uncomment"
        'un-comment selected text
        ActiveDoc.UncommentBlock
    Case "PreviousLine"
        'go to last pos
        ActiveDoc.PreviousLine
    Case "NextLine"
        'go to next pos
        '(if we have previously moved backwards)
        ActiveDoc.NextLine
    Case "InsertTag"
        ActiveDoc.InsertTag
    Case "ToggleFlag"
        ActiveDoc.ToggleFlag
    Case "NextFlag"
        pNextFlag
    Case "LastFlag"
        pLastFlag
    Case "ClearFlags"
        pClearFlags
    Case Else
        If Index = 3 Then
            'addin tb
            'init the Add-in dll
            InitAddIns
            'load the selected add-in
            lItem = CLng(Right$(sKey, Len(sKey) - 3))
            If (cAddIns.Add(lItem)) Then cAddIns.Tool(lItem).ShowDialog
        End If
    End Select
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Main.ToolbarClick"
    Exit Sub
End Sub

'handles drop down for New...
Private Sub tbrMain_DropDownPress(Index As Integer, ByVal lButton As Long)
    Dim X As Long, Y As Long
    'get the dropdown pos
    tbrMain(0).GetDropDownPosition lButton, X, Y
    'adjust value if we are at the bottom
    If picHolder.Align = vbAlignBottom Then Y = Y - 1000
    'build the correct menu
    BuildPopupMenu 3
    'display the popup menu
    ctlPopMenu.ShowPopupMenu Me, "mnuNewPopup", X, Y
    'remove the menu
    DeletePopupMenu 3
End Sub

Private Sub rbrMain_HeightChanged(lNewHeight As Long)
    'adjust height of container
    picHolder.Height = lNewHeight * Screen.TwipsPerPixelY
End Sub
Private Sub picHolder_Resize()
    'tell rebar control it needs to resize
    rbrMain.RebarSize
    'adjust height
    picHolder.Height = rbrMain.RebarHeight * Screen.TwipsPerPixelY
End Sub

'updates the revert menu/numbering tb item
Public Sub SetToolbars()
    'abort if there are no documents open, we are starting up, or exiting!
    If DocOpen = False Or bClosing = True Then Exit Sub
    If Not ActiveDoc Is Nothing Then
        
        'set lines tb item state
        tbrMain(0).ButtonPressed("Numbering") = ActiveDoc.ShowLines
    End If
End Sub

Private Sub SetDocsOpen()
    'adjusts toolbars for when there is a doc open
    Dim i As Integer
    'abort if we are exiting DevPad
    If bClosing = True Or m_bDocsOpen = True Then Exit Sub
    'enable the menu items we disabled...
    EnableItems True
    'update the recent files
    LoadRecentFiles
    'adjust the project menu as needed
    SetProjectMenu IsLoaded("frmProject")
    m_bDocsOpen = True
End Sub
Private Sub SetNoDocsOpen()
    'adjusts toolbars for when there is no docs open
    Dim i As Integer
    'abort if we are exiting DevPad
    If bClosing = True Or m_bDocsOpen = False Then Exit Sub
    'enable/disable menu items and toolbars
    EnableItems False
    'enable a few buttons we want to leave alone...
    tbrMain(0).ButtonEnabled("New") = True
    tbrMain(0).ButtonEnabled("Open") = True
    tbrMain(0).ButtonEnabled("Help") = True
    tbrMain(0).ButtonEnabled("ShowProject") = True
    ' Update the recent files
    LoadRecentFiles
    ' Set the status bar
    SetProjectMenu IsLoaded("frmProject")
    m_bDocsOpen = False
End Sub
Private Sub EnableItems(bEnable As Boolean)
Dim i As Long
    'enable/disable all the toolbar items
    For i = 0 To tbrMain(0).ButtonCount
        tbrMain(0).ButtonEnabled(i) = bEnable
    Next
'    For i = 0 To tbrMain(1).ButtonCount
'        tbrMain(1).ButtonEnabled(i) = bEnable
'    Next
    For i = 0 To tbrMain(2).ButtonCount
        tbrMain(2).ButtonEnabled(i) = bEnable
    Next
    With ctlPopMenu
        ' disable/enable specific menu items...
        .Enabled("InsertTextFromFile") = bEnable
        .Enabled("FileSave") = bEnable
        .Enabled("FileSaveAs") = bEnable
        .Enabled("FileSaveAll") = bEnable
        .Enabled("FileClose") = bEnable
        .Enabled("FileCloseAll") = bEnable
        '.Enabled("FileRevert") = bEnable
        .Enabled("FilePrint") = bEnable
        '.Enabled("FileProperties") = bEnable
        .Enabled("ToolsOpenPrj") = bEnable
        .Enabled("ToolsRunPrj") = bEnable
        .Enabled("ToolsCompilePrj") = bEnable
        .Enabled("ViewWordWrap") = bEnable
        .Enabled("ViewNoWordWrap") = bEnable
        .Enabled("ViewWYSIWYG") = bEnable
        'disable all edit/insert menu items
        For i = 1 To .Count
            If Left$(.MenuKey(i), 4) = "Edit" And Left$(.MenuKey(i), 8) <> "EditFind" Then
                .Enabled(i) = bEnable
            ElseIf Left$(.MenuKey(i), 6) = "Insert" Then
                .Enabled(i) = bEnable
            End If
        Next
    End With
End Sub
Private Sub ShowHideTB(bCurrent As Boolean, lBand As Long, ByRef lItem As Long)
    'shows/hides a toolbar
    'check/uncheck menu item for this toolbar
    ctlPopMenu.Checked(lItem) = Not bCurrent
    'hide/show the toolbar
    rbrMain.BandVisible(lBand) = Not bCurrent
End Sub

'*** Subclassing ***

Private Property Let ISubclass_MsgResponse(ByVal RHS As vbwSubClass.EMsgResponse)
End Property
Private Property Get ISubclass_MsgResponse() As vbwSubClass.EMsgResponse
    ISubclass_MsgResponse = emrPreprocess
End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim tCDS As COPYDATASTRUCT
Dim b() As Byte
Dim sCommand As String

    Select Case iMsg
    Case WM_COPYDATA
        'this message is sent by another instance of devpad
        'starting up, to pass on its command parameter
        ' Copy for processing:
        CopyMemory tCDS, ByVal lParam, Len(tCDS)
        Debug.Assert (tCDS.cbData > 0)
        If (tCDS.cbData > 0) Then
            ReDim b(0 To tCDS.cbData - 1) As Byte
            CopyMemory b(0), ByVal tCDS.lpData, tCDS.cbData
            sCommand = StrConv(b, vbUnicode)
            ' We've got the info, now do it:
            ParseCommand sCommand
        Else
            ' no data
        End If
    End Select
End Function

Public Sub pAttachMessages()
    'attempt to catch WM_COPYDATA message
    On Error Resume Next
    If InDevelopment = False Then
        AttachMessage Me, hWnd, WM_COPYDATA
    End If
End Sub
Private Sub pDetachMessages()
    On Error Resume Next
    'remove subclassing
    DetachMessage Me, hWnd, WM_COPYDATA
End Sub

'*** Window Functions ***

' Updates the list of documents
Public Sub UpdateWindowList()
On Error GoTo ErrHandler
    If Not cWorkspace Is Nothing Then cWorkspace.Changed = True
    'update list in project window, if loaded
    If IsLoaded("frmProject") Then frmProject.UpdateDocs
    'update list in find window, if loaded
    If IsLoaded("frmFind") Then frmFind.ListDocs
    'update the toolbars & menus... (ie revert)
    SetToolbars
    SetDocStatus
    UpdateFlags
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Main.ListWindows"
End Sub
' Lists the windows for the Window List menu
Private Sub ListWindows()
On Error GoTo ErrHandler
    Static WinCount     As Integer
    
    Dim bChecked        As Boolean
    Dim i               As Long
    
    'Remove all the previous window items
    For i = 1 To WinCount Step 1
        ctlPopMenu.RemoveItem "WindowItem" & i
    Next
    'remove 'no windows' item if needed
    If WinCount = -1 Then ctlPopMenu.RemoveItem ("WindowNoWindows")
    If cDocuments.Count = 0 Then
        'no windows open, add (No Windows Open) item instead
        ctlPopMenu.AddItem LoadResString(187), "WindowNoWindows", , , ctlPopMenu.MenuIndex("mnuWindow"), , , False
        WinCount = -1
    Else
        'populate the window list
        For i = 1 To cDocuments.Count Step 1
            'check the item if the current form matches this form
            bChecked = (cDocuments.Item(i).hWnd = ActiveDoc.hWnd)
            'add the item to the menu
            ctlPopMenu.AddItem cDocuments.Item(i).DocumentCaption, "WindowItem" & i, , cDocuments.Item(i).DocID, ctlPopMenu.MenuIndex("mnuWindow"), IndexForKey("PAD"), bChecked
        Next
        WinCount = i - 1
    End If
    'change wincount from 0 so we can tell the difference
    'between initial value, and no windows
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Main.ListWindows"
End Sub

'*** MDI Form Code ***

Private Sub MDIForm_Load()
    'make the find combo box flat
    Set cFindCombo = New clsFlatCombo
    cFindCombo.Attach cboFind.hWnd
    m_bDocsOpen = True
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim lCount  As Long
Dim i       As Long
    If bStartUp Then
        'if we are starting up, abort close!
        Cancel = -1
    Else
        'reset flag
        bCancelClose = False
        'save the list of open documents
        For i = 1 To cDocuments.Count
            SaveSetting REG_KEY, "History", "LastOpen" & lCount, cDocuments.Item(i).FileName
        Next
        SaveSetting REG_KEY, "History", "LastOpenCount", cDocuments.Count
        'save workspace
        If Not cWorkspace Is Nothing Then
            'save the workspace name
            SaveSetting REG_KEY, "Settings", "LastWorkspace", cWorkspace.FileName
            If cWorkspace.FileName <> "" Then
                'close the workspace
                If cWorkspace.CloseWorkspace(True) = False Then
                    'abort...
                    Cancel = -1
                    Exit Sub
                End If
            End If
        End If
        
        'display save dialog listing unsaved documents
        frmSave.PrepareForExit
        'check flag state now
        If bCancelClose = True Then
            ' abort exit
            Cancel = -1
            Exit Sub
        End If
        'we are closing!
        bClosing = True
        If IsLoaded("frmProject") Then
            'unload the project window...
            Unload frmProject
            'check flat state
            If bCancelClose = True Then
                'abort
                Cancel = -1
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
On Error GoTo ErrHandler
    SetStatusBar "Closing Add-Ins..."
    If Not cAddIns Is Nothing Then
        'unload all the loaded add-ins
        cAddIns.UnloadAll
        cAddIns.Terminate
    End If
    Set cAddIns = Nothing
    SetStatusBar "Closing Components..."
    
    'save visibility of Bands...
    With rbrMain
        SaveSetting REG_KEY, "Settings", "ShowMainTB", Abs(.BandVisible(0))
        SaveSetting REG_KEY, "Settings", "ShowFindTB", Abs(.BandVisible(1))
        SaveSetting REG_KEY, "Settings", "ShowCodingTB", Abs(.BandVisible(2))
        SaveSetting REG_KEY, "Settings", "ShowAddInTB", Abs(.BandVisible(3))
    End With
    'remove subclassing
    pDetachMessages
    'clear classes
    Set cFindCombo = Nothing
    Set cFindHistory = Nothing
    Set cReplaceHistory = Nothing
    Set cFileHistory(1) = Nothing
    Set cFileHistory(2) = Nothing
    Set cFileHistory(3) = Nothing
    Set cPrint = Nothing
    Set cPreview = Nothing
    Set cText = Nothing
    SetStatusBar "Quitting..."
    Set cStatusBar = Nothing
    Hide
    'finish of exiting in a module, so we can unload this form
    ExitDevPad
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "MainForm.Unload"
End Sub

Public Function GetMDIClienthWnd() As Long
    'return the mdi client hwnd
    GetMDIClienthWnd = FindWindowEx(Me.hWnd, 0, "MDICLIENT", vbNullString) 'GetWindow(Me.hWnd, GW_CHILD)
End Function

'*** drag drop ***

' Thanks to Austin Kauffman for this code
Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' check the format of the data that is being dropped
    DropFiles Data
End Sub
Private Sub MDIForm_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    ' check the data to see if it is what we will allow. if not so "no drop"
    If Data.GetFormat(vbCFFiles) Then
        Effect = vbDropEffectCopy And Effect
    Else
        Effect = vbDropEffectNone
    End If
End Sub
Public Sub DropFiles(Data As Variant) 'dataobject
    Dim i As Long
    On Error Resume Next
    ' check the format of the data that is being dropped
    If Data.GetFormat(vbCFFiles) = True Then
        For i = 1 To Data.Files.Count
            ' try to load file
            LoadFileDefault Data.Files(i)
        Next
    End If
End Sub

'*** other operations ***
Public Sub UpdateFileMenu(ByVal sFileName As String, iType As Integer)
    'adds a file to the history list
    cFileHistory(iType).Add sFileName
    ' Update the list of the most recently opened files in the File menu control array.
    LoadRecentFiles
End Sub
Private Sub SaveAll()
    Dim i As Long
    For i = 1 To cDocuments.Count
        ' save the file
        ' and abort save all if dialog cancelled
        If cDocuments.Item(i).Save = False Then Exit For
    Next
End Sub
Private Sub CloseAll()
    frmSave.PrepareForExit
End Sub

'Sets the word wrap mode...
Private Function SetWordWrap(iWordWrap As Integer)
    'save mode to local variable
    vDefault.nWordWrap = iWordWrap
    'store in registry
    SaveSetting REG_KEY, "Settings", "WordWrap", vDefault.nWordWrap
    'apply changes to current document
    ActiveDoc.ViewMode = vDefault.nWordWrap
End Function

Public Sub ShowProjectWindow()
    'displays the project window
    Load frmProject
    frmProject.ShowWindow False
End Sub
'*** Status Bar Code ***

Private Sub picStatus_Paint()
    'tell the status bar to redraw itself
    If Not cStatusBar Is Nothing Then cStatusBar.Draw
End Sub
Private Sub picStatus_Resize()
    Dim tR As RECT
    'redraw the status bar
    picStatus_Paint
    If Not cStatusBar Is Nothing Then
        'get the shape of the main status panel
        cStatusBar.GetPanelRect "Status", tR 'lLeft, ltop, lRight, lBottom
        If tR.Right - tR.Left > 150 Then tR.Right = tR.Left + 150
        'move the progress bar to fit inside it
        ctlProg.Move tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top
    End If
End Sub
'*** File History ***

Public Sub InitFindHistory()
    'initialize find & replace history
    Set cFindHistory = New clsHistory
    Set cReplaceHistory = New clsHistory
    'set registry keys
    cFindHistory.RegKey = "FindHistory"
    cFindHistory.RegSection = "History"
    cReplaceHistory.RegKey = "ReplaceHistory"
    cReplaceHistory.RegSection = "History"
    'load the history!
    LoadFindHistory
End Sub
Public Sub InitFileHistory()
Dim i As Long
    'init the file/project history
    For i = 1 To 3
        Set cFileHistory(i) = New clsHistory
        cFileHistory(i).RegKey = "File"
    Next i
    'set the registry keys
    'file...
    cFileHistory(1).RegSection = "FileHistory"
    cFileHistory(1).MaxCount = 30
    cFileHistory(2).RegSection = "ProjectHistory"
    cFileHistory(3).RegSection = "WorkspaceHistory"
    'load the history files
    LoadRecentFiles
End Sub
Public Sub LoadFindHistory()
    Dim cItems As Collection
    Dim i As Long
    'load the find history...
    Set cItems = cFindHistory.Items
    cboFind.Clear
    For i = 1 To cItems.Count
        cboFind.AddItem cItems(i)
    Next
    'update items...
    cboFind_Change
End Sub
Public Sub LoadRecentFiles()
    'loads both the file and project history
    LoadFileHistory 1, "FileRecentFile"
    LoadFileHistory 2, "FileRecentProject"
    LoadFileHistory 3, "FileRecentWorkspace"
End Sub
Public Sub LoadFileHistory(iType As Integer, sRecentMenu As String)
    Dim cItems      As Collection
    Dim i           As Long
    Dim sCaption    As String

    With ctlPopMenu
        'remove none item, if there
        If .MenuExists(sRecentMenu & "None") Then .RemoveItem (sRecentMenu & "None")
        ' remove all the items
        .ClearSubMenusOfItem (sRecentMenu)
        'get the history collection
        Set cItems = cFileHistory(iType).Items
        If cItems.Count = 0 Then
            'there are no items... add a (none) item
            .AddItem LoadResString(1243), sRecentMenu & "None", , , .MenuIndex(sRecentMenu), , , False
        Else
            'add all the items to the menu
            For i = 1 To IIf(cItems.Count > 10, 10, cItems.Count) 'upto a max of 10
                sCaption = cItems(i)
                If Len(sCaption) > 70 Then
                    If InStr(70, sCaption, "\") <> 0 Then
                        sCaption = Left$(sCaption, 3) & "..." & Right$(sCaption, Len(sCaption) - InStr(70, sCaption, "\") + 1)
                    End If
                    'If sCaption = "..." & cItems(i) Then sCaption = cItems(i)
                   ' sCaption = Left$(sCaption, 6) & "..." & Right$(sCaption, 61)
                End If
                .AddItem "&" & i & " " & sCaption, sRecentMenu & i, , i, .MenuIndex(sRecentMenu)
            Next
        End If
    End With
End Sub


Private Sub Revert()
    Dim sFilePath   As String
    Dim frmForm     As IDevPadDocument
    Dim Result      As ShowYesNoResult
    ' if no file is specified, presume it is the active form
    Set frmForm = ActiveDoc
    With frmForm
        '"Are you sure you want to revert to the saved version of "
        If cDialog.ShowYesNo(LoadResString(1274) & LoadResString(1273) & .DocumentCaption & "?", False) = No Then Exit Sub
        ' get file location
        sFilePath = .FileName
        ' pretend the file has not changed
        .Modified = False
    End With
    ' unload specified form
    Unload frmForm
    ' load the file again
    cDocuments.LoadFile sFilePath
End Sub
Public Sub OpenLinkedFile()
On Error GoTo ErrHandler
    Dim sValue As String
    sValue = ActiveDoc.CursorFile(True)
    If sValue <> "" Then
        If Mid$(sValue, Len(sValue), 1) = "\" Then
            'folder...
            SaveSetting REG_KEY, "Settings", "DefaultFilePath", sValue
            OpenFile
        Else
            If InStr(1, sValue, ":\\") <> 0 Or LCase$(Left$(sValue, 7)) = "mailto:" Then
                ShellFunc sValue
            Else
                If InStr(1, GetSetting(REG_KEY, "Settings", "UseDefaultEditor", ".gif .jpg"), GetExtension(sValue)) = 0 Then
                    cDocuments.LoadFile sValue
                Else
                    ShellFunc sValue
                End If
            End If
        End If
    End If
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Main.OpenLinkedFile"
End Sub


'*** Flags implementation ***
Public Sub UpdateFlags()
Dim bEnabled As Boolean
Dim i As Long
    'see if there are any flags in any docs...
    If GetSetting(REG_KEY, "Settings", "GlobalFlags", 1) = "1" Then
        For i = 1 To cDocuments.Count Step 1
            bEnabled = IIf(bEnabled, True, (cDocuments.Item(i).FlagCount <> 0))
        Next i
    Else
        bEnabled = DocOpen
        If bEnabled = True Then bEnabled = (ActiveDoc.FlagCount <> 0)
    End If
    EnableTB "NextFlag", bEnabled, 2
    EnableTB "LastFlag", bEnabled, 2
    EnableTB "ClearFlags", bEnabled, 2
End Sub
Private Sub pNextFlag()
Dim nIndex As Long
Dim nOrigIndex As Long
Dim bResult As Boolean

    nIndex = cDocuments.DocumentIndex(ActiveDoc.DocID)
    nOrigIndex = nIndex
    'try the current doc...
    'allow looping if the flags are not global
    bResult = cDocuments.Item(nIndex).NextFlag(GetSetting(REG_KEY, "Settings", "GlobalFlags", 1) <> "1")
    If bResult = False And GetSetting(REG_KEY, "Settings", "GlobalFlags", 1) = "1" Then
        'no good in current doc... and we are allowing global flags...
        Do
            'get the next document
            nIndex = cDocuments.NextIndex(nIndex)
            'abort if we are back to the original one
            If nIndex = nOrigIndex Then Exit Do
            'only try if there are flags...
            If cDocuments.Item(nIndex).FlagCount > 0 Then
                'go to the first flag
                cDocuments.Item(nIndex).FirstFlag
                'activate the doc
                cDocuments.Item(nIndex).SetFocus
                Exit Do
            End If
        Loop
        'if still no match, then run NextFlag allowing looping
        cDocuments.Item(nIndex).NextFlag True
    End If
End Sub
Private Sub pLastFlag()
Dim nIndex As Long
Dim nOrigIndex As Long
Dim bResult As Boolean

    nIndex = cDocuments.DocumentIndex(ActiveDoc.DocID)
    nOrigIndex = nIndex
    'try the current doc...
    'allow looping if the flags are not global
    bResult = cDocuments.Item(nIndex).PreviousFlag(GetSetting(REG_KEY, "Settings", "GlobalFlags", 1) <> "1")
    If bResult = False And GetSetting(REG_KEY, "Settings", "GlobalFlags", 1) = "1" Then
        'no good in current doc, and we allow global flags
        Do
            'get the next document
            nIndex = cDocuments.PreviousIndex(nIndex)
            'abort if we are back to the original one
            If nIndex = nOrigIndex Then Exit Do
            'only try if there are flags...
            If cDocuments.Item(nIndex).FlagCount > 0 Then
                'go to the first flag
                cDocuments.Item(nIndex).LastFlag
                'activate the doc
                cDocuments.Item(nIndex).SetFocus
                Exit Do
            End If
        Loop
        'if still no match, then run NextFlag allowing looping
        If nIndex = nOrigIndex Then cDocuments.Item(nIndex).PreviousFlag True
    End If
End Sub
Private Sub pClearFlags()
Dim i As Long
    If GetSetting(REG_KEY, "Settings", "GlobalFlags", 1) = "1" Then
        'clear flags from all docs
        For i = 1 To cDocuments.Count Step 1
            cDocuments.Item(i).ClearFlags
        Next i
    Else
        'only current doc
        If DocOpen Then ActiveDoc.ClearFlags
    End If
End Sub
Public Property Get FileHistoryItems() As Collection
    Set FileHistoryItems = cFileHistory(1).Items
End Property
