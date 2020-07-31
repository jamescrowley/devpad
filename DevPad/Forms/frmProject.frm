VERSION 5.00
Object = "{C4925FC3-1606-11D4-82BB-004005AAE138}#5.2#0"; "VBWIML.OCX"
Object = "{1F5999A2-1D0B-11D4-82CF-004005AAE138}#6.1#0"; "VBWTBA~1.OCX"
Object = "{017E002E-D7CC-11D2-8E21-44B10AC10000}#22.0#0"; "VBWGRID.OCX"
Object = "{5C0E11AE-2C8C-4C35-BC7A-D9B469D5DE4D}#6.1#0"; "VBWTRE~1.OCX"
Begin VB.Form frmProject 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   5235
   ClientLeft      =   615
   ClientTop       =   2475
   ClientWidth     =   6450
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   Begin DevPad.ctlFrame ctlFrame1 
      Height          =   345
      Left            =   105
      TabIndex        =   7
      Top             =   330
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   609
      Begin vbwTBar.cToolbar tbrMain 
         Left            =   975
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   318
      End
      Begin vbwTBar.cToolbarHost tbhMain 
         Height          =   210
         Left            =   270
         TabIndex        =   6
         Top             =   75
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   370
         BorderStyle     =   0
      End
   End
   Begin VB.PictureBox picProjects 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   465
      ScaleHeight     =   3975
      ScaleWidth      =   4935
      TabIndex        =   0
      Top             =   825
      Width           =   4935
      Begin vbwTreeView.TreeView tvProject 
         Height          =   1485
         Left            =   660
         TabIndex        =   4
         Top             =   1125
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   2619
         Lines           =   0   'False
         LabelEditing    =   0   'False
         PlusMinus       =   0   'False
         RootLines       =   0   'False
         ToolTips        =   0   'False
         BorderStyle     =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxScrollTime   =   0
      End
      Begin VB.TextBox txtSymbolPath 
         Height          =   285
         Left            =   1665
         TabIndex        =   5
         Tag             =   "ProjectName"
         Top             =   120
         Width           =   2190
      End
      Begin vbAcceleratorGrid.vbalGrid lvwSymbols 
         Height          =   765
         Left            =   4005
         TabIndex        =   3
         Top             =   1545
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   1349
         RowMode         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Header          =   0   'False
         HeaderButtons   =   0   'False
         BorderStyle     =   2
         DisableIcons    =   -1  'True
         DefaultRowHeight=   17
      End
      Begin vbAcceleratorGrid.vbalGrid lvwAllDocs 
         Height          =   975
         Left            =   4065
         TabIndex        =   2
         Top             =   450
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1720
         MultiSelect     =   -1  'True
         RowMode         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   2
         DisableIcons    =   -1  'True
         DefaultRowHeight=   18
      End
      Begin VB.Label lblMsg 
         Caption         =   "Please Wait..."
         Height          =   1095
         Left            =   75
         TabIndex        =   1
         Top             =   1590
         Width           =   2655
      End
   End
   Begin DevPad.vbwCaption vbwDock 
      Align           =   1  'Align Top
      Height          =   270
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   476
   End
   Begin vbwIml.vbalImageList vbalMain 
      Left            =   5280
      Top             =   3975
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   8
      Size            =   14100
      Images          =   "frmProject.frx":0000
      KeyCount        =   15
      Keys            =   "FILE_VBPÿSCRIPTÿDOCSÿOPENÿMISCÿFOLDERUPÿFOLDERCLOSEDÿLIVEFOLDERÿFILE_TXTÿEMAILÿHYPERLINKÿFOLDEROPENÿFILE_VBÿFOLDERSHORTCUTÿALERT"
   End
   Begin VB.Menu mnuAllDocsTOP 
      Caption         =   "&All Docs"
      Visible         =   0   'False
      Begin VB.Menu mnuAllDocs 
         Caption         =   "&Save"
         Index           =   0
      End
      Begin VB.Menu mnuAllDocs 
         Caption         =   "&Save As..."
         Index           =   1
      End
      Begin VB.Menu mnuAllDocs 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuAllDocs 
         Caption         =   "&Close"
         Index           =   3
      End
      Begin VB.Menu mnuAllDocs 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuAllDocs 
         Caption         =   "Copy Path"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmProject"
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
'WARNING: This code is very messy!
Option Explicit

'*** Structures ***

Private Type SHFILEINFO
    hIcon As Long                      '  out: icon
    iIcon As Long          '  out: icon index
    dwAttributes As Long               '  out: SFGAO_ flags
    szDisplayName As String * MAX_PATH '  out: display name (or path)
    szTypeName As String * 80         '  out: type name
End Type

Private Type ItemInfo
    nType As Integer
    sPath As String
    sName As String
    sFolder As String
End Type
Private Type FolderInfo
    sPath As String
    sName As String
    bExpanded As Boolean
    'sTag As String
End Type

'*** API Constants ***
Private Const VK_RETURN = &HD
Private Const VK_DELETE = &H2E
Private Const VK_F5 = &H74

Private Const SHGFI_ICON = &H100
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1

'*** API ***
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ToAscii Lib "user32" (ByVal uVirtKey As Long, ByVal uScanCode As Long, lpbKeyState As Byte, lpwTransKey As Long, ByVal fuState As Long) As Long

'*** Member Variables ***

Private hDragItem           As Long
Private bOrderChanged       As Boolean
Private lCurrentTab         As Long
Private sProjectPath        As String
Private sPaths()            As String
Private sOldPath            As String
Private bChanged            As Boolean
Private bProjectOpen        As Boolean
Private bBusy               As Boolean
Private bClosing            As Boolean
Private bItemClicked        As Boolean
Private lPathsCount         As Long
Private bNoUpdate           As Boolean

Private pProjectInfo        As ProjectInfo
Private m_cProjectEx        As IDevPadProjectEx
Private cSymbolCombo        As clsFlatCombo
Private m_sProjectRoot      As String
Private m_tProjectType      As ProjectTypes

Private m_cInserts()        As InsertItem
Private m_cFolders()        As FolderInfo
Private m_cItems()          As ItemInfo
Private m_lInsertsCount     As Long
Private m_lItemCount        As Long
Private m_lFolderCount      As Long
Private Property Get ProjectRoot() As String
Dim sRoot As String
    sRoot = m_sProjectRoot
    If sRoot = ".\" Or sRoot = "" Then
        sRoot = GetFolder(sProjectPath) & "\"
    End If
    If Right$(sRoot, 1) <> "\" Then sRoot = sRoot & "\"
    ProjectRoot = sRoot
End Property
Private Property Let ProjectRoot(sNew As String)
    If sNew = GetFolder(sProjectPath) & "\" Then sNew = ""
    m_sProjectRoot = sNew
End Property
'return a folder item's tag
'Public Property Let FolderTag(lIndex As Long, sTag As String)
'   ' m_cFolders(tvProject.ItemData(lIndex)).sTag = sTag
'End Property
'Public Property Get FolderTag(lIndex As Long) As String
'   ' FolderTag = m_cFolders(tvProject.ItemData(lIndex)).sTag
'End Property
'return an item's file...
Public Property Get ItemFile(lIndex As Long) As String
    ItemFile = m_cItems(tvProject.ItemData(lIndex)).sPath
End Property
'show the project window
Public Sub ShowWindow(bInitialPos As Boolean)
    vbwDock.ShowWindow bInitialPos
End Sub



Private Sub ctlFrame1_Resize()
    'resize toolbar to fit host
    tbhMain.Width = ctlFrame1.Width - 30
End Sub

Private Sub Form_Activate()
'On Error Resume Next
'    If tvProject.Visible Then tvProject.SetFocus
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'display help window
    If KeyCode = vbKeyF1 Then cDialog.ShowHelpTopic 5, hWnd
End Sub

Private Sub Form_Load()
    'init array
    ReDim sPaths(0 To 0)
    'don't waste time...
    tvProject.DisableCustomDraw = True
    'Set vbwDock = New clsDock
    With vbwDock
        Set .frmDockWindow = Me
        Set .frmParent = frmMainForm
        Set .picHolder = frmMainForm.picProjectHolder
        'Set .picSizer = frmMainForm.picSizer
        .RegAppName = REG_KEY
        .RegKey = "frmProject"
        .RegSection = "WindowSettings"
       ' Set .picSizerTop = picSizerTop
    End With
    'set treeview properties
    With tvProject
        .LabelEditing = True
        .Lines = True
        .PlusMinus = True
        .Rootlines = False
        .ShowSelected = True
        .NoDragDrop = False
    End With
    sOldPath = "-1"
    SetThin3DBorder txtSymbolPath.hWnd
    'lvwAllDocs.HeaderHotTrack = False
    'tbsTabs.ImageList = frmMainForm.vbalMain
    lvwSymbols.ImageList = vbalMain
    'prepare controls
    
    With tbrMain
        'build the toolbar
        'set image list
        .ImageSource = CTBExternalImageList
        .SetImageList vbalMain, CTBImageListNormal
        'create tb
        .CreateToolbar 16, , True, True
        'add buttons
        .AddButton "Project", IndexForKey("FILE_VBP"), , , Empty, CTBCheckGroup, "Project"
        .AddButton "Code Library", IndexForKey("SCRIPT"), , , Empty, CTBCheckGroup, "Insert"
        .AddButton "", , , , Empty, CTBSeparator, "Sep"
        .AddButton "Open Document List", IndexForKey("DOCS"), , , Empty, CTBCheckGroup, "Docs"
    End With
    
    With tbhMain
        'init toolbar host
        'no border
        .BorderStyle = etbhBorderStyleNone
        'set width
        .Width = ctlFrame1.Width - 30
        .Left = 15
        .Top = 30
        'set height
        .Height = tbrMain.ToolbarHeight * Screen.TwipsPerPixelX
        ctlFrame1.Height = .Height + 60
        'capture main toolbar
        .Capture tbrMain
    End With
    
    
    lvwAllDocs.AddColumn "Document", "Document", , , 107
    lvwAllDocs.AddColumn "Path", "Path", , , 80
    lvwSymbols.Top = txtSymbolPath.Height + 50
    lvwSymbols.Left = 12
    lvwSymbols.AddColumn , , , , (lvwSymbols.Width / Screen.TwipsPerPixelX) + 12

    ChangePath GetSetting(REG_KEY, "Settings", "ProjectInsertPath", "")
    
    tvProject.hImageList = vbalMain.hIml
    lvwAllDocs.ImageList = vbalMain

    If GetSetting(REG_KEY, "Settings", "ProjectReloadLast", "1") = "1" And GetSetting(REG_KEY, "Settings", "LastProject") <> "" Then
        OpenProject GetSetting(REG_KEY, "Settings", "LastProject")
    Else
        NewProject
    End If
    ' restore pos
    tbrMain.RaiseButtonClick Val(GetSetting(REG_KEY, "Settings", "ProjectTab", 1)) - 1
    Form_Resize
End Sub
Private Sub SortInsertList()
    With lvwSymbols.SortObject
        .Clear
        .SortColumn(1) = 1
        .SortOrder(1) = CCLOrderAscending
        .SortType(1) = CCLSortIcon

        .SortColumn(2) = 1
        .SortOrder(2) = CCLOrderAscending
        .SortType(2) = CCLSortString
        lvwSymbols.Sort
    End With
End Sub
Private Sub SortDocsList(lCol As Long)
    Dim bAsc As Boolean
    Static lLastCol As Long
    With lvwAllDocs.SortObject
        .Clear
        .SortColumn(1) = lCol
        .SortOrder(1) = IIf(lLastCol <> lCol, 1, 2) 'CCLOrderAscending
        .SortType(1) = CCLSortString
        lvwAllDocs.Sort
        If lLastCol = lCol Then
            lLastCol = 0
        Else
            lLastCol = lCol
        End If
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If OKToClose = False Then
        'cancel
        Cancel = -1
        frmMainForm.bCancelClose = True
        Exit Sub
    End If
    'NewProject
    tvProject.Clear
    bChanged = False
End Sub
Public Sub Form_Resize()
    On Error Resume Next
   ' With vbwDock
    ctlFrame1.Move 15, vbwDock.Height + 15, ScaleWidth - 30 ', ScaleHeight - 280 '252
    
        'tbsTabs.Move 2, vbwDock.Height, ScaleWidth - 24, ScaleHeight - 280 '252
        picProjects.Move 15, ctlFrame1.Top + ctlFrame1.Height + 30, ScaleWidth - 30, ScaleHeight - ctlFrame1.Top - ctlFrame1.Height - 60 ' 640
    'End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    With frmMainForm
        'Set vbwDock = Nothing

        SaveSetting REG_KEY, "Settings", "ProjectTab", lCurrentTab 'tbsTabs.SelectedItem.Index
        SaveSetting REG_KEY, "Settings", "ProjectInsertPath", sOldPath
        
        bClosing = True
        .SetProjectMenu (False)
        ChangePath ""

        bClosing = False
    End With
    Set cSymbolCombo = Nothing
    Erase sPaths
    Erase m_cInserts
End Sub

Private Sub lvwAllDocs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, bDoDefault As Boolean)
    If Button = vbRightButton And lvwAllDocs.Rows <> 0 Then
        bDoDefault = False
        Dim lRow As Long
        Dim lCount As Long
        Dim i As Long
        lvwAllDocs.CellFromPoint X / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY, lRow, 0
        For i = 1 To lvwAllDocs.Rows
            If lvwAllDocs.CellSelected(i, 1) Then
                lCount = lCount + 1
            End If
        Next
        If lCount < 2 And lRow <> 0 Then lvwAllDocs.SelectedRow = lRow
        BuildPopupMenu 5
        ShowPopup "mnuDocPopup", Me
        DeletePopupMenu 5
'        Dim lpPos As POINTAPI
'        BuildPopupMenu 5
'        GetCursorPos lpPos
'        ScreenToClient hWnd, lpPos
'        frmMainForm.ctlPopMenu.ShowPopupMenu Me, "mnuDocPopup", lpPos.x * Screen.TwipsPerPixelX, lpPos.y * Screen.TwipsPerPixelY
'        DeletePopupMenu 5
'
        'PopupMenu mnuAllDocsTOP
    End If
End Sub

Private Sub lvwSymbols_DblClick(ByVal lRow As Long, ByVal lCol As Long)

Dim lEntryLen As Long
Dim lLen As Long
Dim lCursorPos As Long
Dim sText As String
Dim sKey As String

On Error Resume Next
    If lRow <> 0 Then
        sKey = lvwSymbols.RowKey(lRow)
        If Left$(sKey, 1) = "&" Then
            If sKey = "&^UP" Then
                UpFolder
            Else
                'is a folder
                ChangePath (sOldPath & "\" & Right(sKey, Len(sKey) - 1))
            End If
        Else
           ' frmMainForm.tbhMenu.AllowLock = False
            sText = GetInsertText(lvwSymbols.SelectedRow, lEntryLen, lCursorPos, lLen)
            With ActiveDoc
                '.NoEnterFocus = True
                '.InsertCode sText, True, True
                .SelText = sText
                .SelStart = .SelStart - lEntryLen + lCursorPos
                .SelLength = lLen
                '.NoEnterFocus = False
                '.SetFocus
            End With
        End If
    End If
End Sub
Private Function GetInsertText(lIndex As Long, ByRef lEntryLen As Long, ByRef lCursorPos As Long, ByRef lLen As Long) As String
    Dim sText As String
    Dim i As Long
    Dim lNewLen As Long
    Dim lOldLen As Long
    Dim lNamePos As Long
    Dim lData As Long
    Dim iFileNum As Integer
    Dim sConstant As String
    Dim sConstantValue As String
    Dim sConstantName As String
    Dim lPos As Long
    Dim sIndent As String
    'get the item we get the data from
    lData = lvwSymbols.RowKey(lIndex)
    'retreive the text
    sText = m_cInserts(lData).sValue  'lvwSymbols.RowKey(lIndex)
    'indent string
    sIndent = ActiveDoc.GetIndent(ActiveDoc.LineText)

    'save original len
    lOldLen = Len(sText)
    'get the pos the cursor is supposed to go to
    lCursorPos = m_cInserts(lData).lStart
    If lCursorPos = 0 Then lCursorPos = lOldLen
    'get the length
    lLen = m_cInserts(lData).lLen
    
    iFileNum = FreeFile
    On Error Resume Next
    Open App.Path & "\constants.txt" For Input As iFileNum
    If Err = 0 Then
        Do While Not EOF(iFileNum)
            Line Input #iFileNum, sConstant
            'ignore comments
            If Left$(sConstant, 1) <> ";" And sConstant <> "" Then
                'get the position of the = sign
                lNamePos = InStr(1, sConstant, "=")
                'get the constants name & value
                sConstantName = Left$(sConstant, lNamePos - 1)
                sConstantValue = Right$(sConstant, Len(sConstant) - lNamePos)
               ' sConstantValue = Right$(sConstantValue, Len(sConstant) - (lNamePos + 1))
                'get the position
                lNamePos = InStr(1, sConstantValue, "@")
                If lNamePos = 0 Then lNamePos = Len(sConstantValue) + 1
                Select Case Left$(sConstantValue, lNamePos - 1)
                Case "DATE"
                    sConstantValue = Format(Date, Right$(sConstantValue, Len(sConstantValue) - lNamePos))
                Case "TIME"
                    sConstantValue = Format(Time, Right$(sConstantValue, Len(sConstantValue) - lNamePos))
                Case "FILENAME"
                    sConstantValue = ActiveDoc.FileName
                Case "FILETITLE"
                    sConstantValue = GetCaption(ActiveDoc.FileName)
                Case "FILEFOLDER"
                    sConstantValue = GetFolder(ActiveDoc.FileName)
                Case Else
                End Select
                AdjustValue sText, sConstantName, sConstantValue, lCursorPos
            End If
        Loop
        Close iFileNum
    End If
    If GetSetting(REG_KEY, "Settings", "ApplyIndentToCodeLib", "1") = "1" Then
        'apply current indenting
        AdjustValue sText, vbCrLf, vbCrLf & sIndent, lCursorPos
    End If

    lEntryLen = Len(sText)
    GetInsertText = sText
End Function
Private Sub AdjustValue(ByRef sText As String, ByVal sFind As String, ByVal sReplace As String, ByRef lCursorPos As Long)
    'This sub adjusts the cursor pos value when replacing text
    Dim sTextLeft As String
    Dim sTextRight As String
    Dim lPos As Long
    Dim lOldLen As String
    lPos = InStr(1, sText, sFind)
    
    Do While lPos <> 0
        lOldLen = Len(sText)
        sTextLeft = Left$(sText, lPos - 1)
        sTextRight = Right$(sText, lOldLen - lPos - Len(sFind) + 1)
        sText = sTextLeft & sReplace & sTextRight
        If lPos < lCursorPos Then lCursorPos = lCursorPos + Len(sText) - lOldLen
        lPos = InStr(lPos + Len(sReplace), sText, sFind)
    Loop
End Sub
Private Sub lvwSymbols_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
    On Error Resume Next
    If KeyCode = vbKeyDelete Then PopupMenuHandle 0, "CodePopDelete" 'mnuInsert_Click (2) 'delete
End Sub
Private Sub lvwSymbols_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, bDoDefault As Boolean)
    If Button = vbRightButton Then
        Dim lRow As Long
        Dim lCol As Long
        Dim sKey As String

        lvwSymbols.CellFromPoint X / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY, lRow, lCol
        
'        mnuInsert(5).Enabled = False
'        mnuInsert(6).Enabled = False
        If lRow = 0 Then
            'DisableInsertMenu True
        Else
            sKey = lvwSymbols.RowKey(lRow)
            lvwSymbols.SelectedRow = lRow
            '
            If Left$(sKey, 1) = "&" Then
                'folder
                If sKey <> "&^UP" Then
                    'mnuInsert(5).Enabled = True
                    'mnuInsert(6).Enabled = True
                End If
                'DisableInsertMenu True
            Else
                'DisableInsertMenu False
            End If
        End If
        BuildPopupMenu 4
        ShowPopup "mnuCodePopup", Me
        DeletePopupMenu 4

'        Dim lpPos As POINTAPI
'        BuildPopupMenu 4
'        GetCursorPos lpPos
'        ScreenToClient hWnd, lpPos
'        frmMainForm.ctlPopMenu.ShowPopupMenu Me, "mnuCodePopup", lpPos.x * Screen.TwipsPerPixelX, lpPos.y * Screen.TwipsPerPixelY
'        DeletePopupMenu 4
'
        'PopupMenu mnuInsertTOP
    Else
        If IsLoaded("frmEntry") Then
            With frmEntry
                If .Visible Then
                    .bNoUnload = True
                    If .ValidateEntry = True Then ShowCurrentEntry
                    .bNoUnload = False
                End If
            End With
        End If
    End If
End Sub

Private Property Let sCaption(sNewCaption As String)
    vbwDock.Caption = sNewCaption
End Property
Private Property Get sCaption() As String
    Caption = vbwDock.Caption
End Property
Private Sub SetCaption()
    Select Case tbrMain.ButtonKey(lCurrentTab - 1) 'tbsTabs.TabKey(tbsTabs.SelectedTab)
    Case "Project"
        If bProjectOpen = False Then
            sCaption = "No Project Open"
            frmMainForm.SetProjectMenu (False)
        Else
            sCaption = LoadResString(1032) & " - " & ProjectName
            frmMainForm.SetProjectMenu (True)
        End If
    Case "Docs"
        sCaption = "Open Documents"
    Case "Insert"
        sCaption = "Code Library"
    End Select
End Sub


Public Function OKToClose() As Boolean
    OKToClose = True
    If bChanged = True Then
        ' project has changed
        Select Case cDialog.ShowYesNo("Save changes to current project?", True)
        Case Yes
            OKToClose = SaveProjectFile(sProjectPath, False)
        Case Cancelled
            OKToClose = False
            Exit Function
        End Select
    End If
End Function
Public Sub ShowProperties()
    Load frmProjectProperties
    With frmProjectProperties
        .txtProjectName.Text = pProjectInfo.sProjectName
        .txtAuthor.Text = pProjectInfo.sProjectAuthor
        .txtDescription.Text = pProjectInfo.sProjectDescription
        .txtFileName.Text = sProjectPath
        .txtRoot.Text = ProjectRoot
        
        LoadShow frmProjectProperties, vbModal
        If Not .bCancel Then
            pProjectInfo.sProjectAuthor = .txtAuthor.Text
            pProjectInfo.sProjectDescription = .txtDescription.Text
            pProjectInfo.sProjectName = .txtProjectName.Text
            ProjectRoot = .txtRoot.Text
            bChanged = True
        End If
        tvProject.Tag = " Project - " & pProjectInfo.sProjectName
        SetCaption
    End With
    Unload frmProjectProperties
End Sub


Private Sub lvwAllDocs_ColumnClick(ByVal lCol As Long)
    SortDocsList lCol
End Sub

Private Sub lvwAllDocs_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
    If lRow <> 0 Then SendMessage frmMainForm.GetMDIClienthWnd, WM_MDIACTIVATE, cDocuments.ItemByID(lvwAllDocs.CellTag(lRow, 1)).DocumenthWnd, 0
    DoEvents
   ' If lvwAllDocs.Visible Then lvwAllDocs.SetFocus
End Sub

Public Sub NextEntry()
    With lvwSymbols
        If .SelectedRow <> .Rows Then
            .SelectedRow = (.SelectedRow + 1)
            lvwSymbols_SelectionChange .SelectedRow, .SelectedCol
            PopupMenuHandle 0, "CodePopEditEntry" 'mnuInsert_Click 1
        End If
    End With
End Sub
Public Sub LastEntry()
    With lvwSymbols
        If .SelectedRow <> 1 Then
            .SelectedRow = (.SelectedRow - 1)
            lvwSymbols_SelectionChange .SelectedRow, .SelectedCol
            PopupMenuHandle 0, "CodePopEditEntry" 'mnuInsert_Click 1
        End If
    End With
End Sub

'Private Sub DisableInsertMenu(bDisable As Boolean)
'    mnuInsert(1).Enabled = Not bDisable
'    mnuInsert(2).Enabled = Not bDisable
'End Sub

Private Sub lvwSymbols_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
    If IsLoaded("frmEntry") Then
        frmEntry.cmdLast.Enabled = Not (lRow = 1)
        frmEntry.cmdNext.Enabled = Not (lRow = lvwSymbols.Rows)
    End If
End Sub


Public Sub AddEntry(sDescription As String, sValue As String, lPos As Long, lLen As Long)
    Dim lRow As Long
    With lvwSymbols
        ReDim Preserve m_cInserts(1 To m_lInsertsCount + 1)
        m_lInsertsCount = m_lInsertsCount + 1
        lRow = .Rows + 1
        .AddRow , m_lInsertsCount  'strValue
        .CellText(lRow, 1) = sDescription
        .CellIcon(lRow, 1) = IndexForKey("FILE_TXT")
        With m_cInserts(m_lInsertsCount)
            .sName = sDescription
            .sValue = sValue
            .lStart = lPos
            .lLen = lLen
        End With
        SortInsertList
    End With
End Sub

Public Sub ChangeEntry(sDescription As String, sValue As String, lPos As Long, lLen As Long, lRow As Long)
On Error Resume Next
    With lvwSymbols
        .CellText(lRow, 1) = sDescription
        With m_cInserts(.RowKey(lRow))
            .sName = sDescription
            .sValue = sValue
            .lStart = lPos
            .lLen = lLen
        End With
    End With
    SortInsertList
End Sub


Private Sub ShowCurrentEntry()
    Dim lSel As Long
    Dim sItem As String
    With lvwSymbols
        lSel = .SelectedRow
        sItem = .RowKey(lSel)
        If IsNumeric(sItem) Then
            frmEntry.DisplayEntry False, .CellText(lSel, 1), m_cInserts(sItem).sValue, m_cInserts(sItem).lStart, m_cInserts(sItem).lLen, lSel         '.RowKey(.SelectedRow), .RowTag(.SelectedRow), .CellTag(.SelectedRow, 1), .SelectedRow
        Else
            frmEntry.DisplayEntry True, "", "", 0, 0, 0
        End If
    End With
End Sub
Public Sub PopupMenuHandle(lItem As Long, Optional sKey As String = "")
    Dim i           As Long
    Dim lRow        As Long
    Dim sPrompt     As String
    Dim sPath       As String
    Dim sResult     As String
    Dim m_cItem     As ItemInfo
    Dim frmDoc      As IDevPadDocument
    
    If sKey = "" Then sKey = frmMainForm.ctlPopMenu.MenuKey(lItem)
    Select Case sKey
    Case "PrjPopEditPath"
        
        With tvProject
            If .ItemData(.Selected) = 0 Then Exit Sub
            LSet m_cItem = m_cItems(.ItemData(.Selected))
            'sTag = sPaths(.ItemData(.Selected))

           ' sType = 'Left$(sTag, 1)
            sPath = pGetAbsolutePath(m_cItem.sPath) '(sTag, Len(sTag) - 1)
            Select Case m_cItem.nType
            Case 1, 4 ' folder
                sPath = cDialog.BrowseForFolder(sPath)
                If sPath = "" Then Exit Sub 'abort
            Case 0 ' file
                If cDialog.ShowOpenSaveDialog(False, "Edit Path", "", sPath) = False Then Exit Sub
                sPath = CmDlg.FileName
            Case 2, 3 ' URL, email
                If m_cItem.nType = 2 Then
                    sPrompt = "URL"
                Else
                    sPrompt = "Email Address"
                End If
                sPath = cDialog.InputBox("Please enter a new " & sPrompt, "Modify " & sPrompt, sPath)
                If sPath = "" Then Exit Sub 'abort
                
            End Select
            If m_cItem.nType = vbwFile Or m_cItem.nType = vbwFolder Then
                sPath = pGetRelativePath(sPath)
            End If
            m_cItem.sPath = sPath
            LSet m_cItems(.ItemData(.Selected)) = m_cItem
            If m_cItem.nType = vbwLiveFolder Then
                pRefreshVirtualTree (.Selected)
            End If
            'sPaths(.ItemData(.Selected)) = sType & sPath
            bChanged = True
        End With
    Case "PrjPopOpenFileDevPad"
        OpenSelectedFile False
    Case "PrjPopOpenFile"
        OpenSelectedFile True
    Case "PrjPopRenameItem"
        Debug.Print "*** This doesn't work! Why?!***"
        tvProject.LabelEdit (tvProject.Selected)
    Case "PrjPopRemoveItem"
        DeleteItem
    Case "CodePopInsert"
        lvwSymbols_DblClick lvwSymbols.SelectedRow, 1
    Case "CodePopNewEntry"
        frmEntry.DisplayEntry True, "", "", 0, 0, 0
    Case "CodePopEditEntry"
        If IsLoaded("frmEntry") Then
            frmEntry.bNoUnload = True
            If frmEntry.ValidateEntry = False Then Exit Sub
        End If
        ShowCurrentEntry
    Case "CodePopNewFolder"
        NewCategory
    Case "CodePopRename"
        sKey = lvwSymbols.RowKey(lvwSymbols.SelectedRow)
        If sKey <> "&^UP" Then
            lRow = lvwSymbols.SelectedRow
            sResult = cDialog.InputBox("Please enter a new name", "Rename Item", lvwSymbols.CellText(lRow, 1))
            If sResult <> "" Then
                If Left$(sKey, 1) = "&" Then 'folder
                    Name App.Path & "\_insert" & sOldPath & "\" & lvwSymbols.CellText(lRow, 1) As App.Path & "\_insert" & sOldPath & "\" & sResult
                    lvwSymbols.CellText(lRow, 1) = sResult
                    lvwSymbols.RowKey(lRow) = "&" & sResult
                Else 'item
                    With m_cInserts(sKey)
                        ChangeEntry sResult, .sValue, .lStart, .lLen, lRow
                    End With
                End If
            End If
        End If
    Case "CodePopDelete"
        sKey = lvwSymbols.RowKey(lvwSymbols.SelectedRow)
        If sKey <> "&^UP" Then
            If Left$(sKey, 1) = "&" Then
                DeleteCategory True
            Else
                '"Are you sure you want to delete the selected item?"
                If cDialog.ShowYesNo(LoadResString(1274) & LoadResString(1269) & LoadResString(1270), False) = No Then Exit Sub
                For i = lvwSymbols.RowKey(lvwSymbols.SelectedRow) To m_lInsertsCount - 1 Step 1
                    LSet m_cInserts(i) = m_cInserts(i + 1)
                Next
                m_lInsertsCount = m_lInsertsCount - 1
                lvwSymbols.RemoveRow lvwSymbols.SelectedRow
                'ChangePath sOldPath
            End If
        End If
    Case "DocPopSave", "DocPopSaveAs", "DocPopClose", "DocPopExplore"
        If sKey = "DocPopClose" Then
            lvwAllDocs.Redraw = False
            bNoUpdate = True
        End If
        
        For i = 1 To lvwAllDocs.Rows Step 1
            If lvwAllDocs.CellSelected(i, 1) Then
                frmMainForm.bCancelClose = False
                Set frmDoc = cDocuments.ItemByID(lvwAllDocs.CellTag(i, 1))
                If Not frmDoc Is Nothing Then
                    Select Case sKey
                    Case "DocPopSave" 'save
                        frmDoc.Save
                    Case "DocPopSaveAs" 'saveas
                        frmDoc.SaveAs
                    Case "DocPopClose" 'close
                        frmDoc.Close
                    Case "DocPopExplore" 'explore
                        sPath = GetFolder(frmDoc.FileName)
                        If sPath <> "" Then
                            'open explorer using ShellExecute
                            ShellFunc sPath, , "explore"
                        End If
                    End Select
                    If frmMainForm.bCancelClose Then Exit For
                End If
            End If
        Next
        If sKey = "DocPopClose" Then
            bNoUpdate = False
            UpdateDocs
            lvwAllDocs.Redraw = True
        End If
    End Select
End Sub
Private Sub ListDocs()
    On Error Resume Next
    Dim Form As IDevPadDocument
    Dim lIndex As Long
    Dim sFileName As String
    Dim i As Long
    'lCount = 1
    lvwAllDocs.Redraw = False
    lvwAllDocs.Clear False
    For i = 1 To cDocuments.Count
'    For Each Form In Forms
'        If Form.Name = "frmDocument" Then
        Set Form = cDocuments.Item(i)
        
            sFileName = Form.FileName
            lvwAllDocs.AddRow

            If Form.Saved = False Then
                lvwAllDocs.CellText(i, 2) = "Not Saved"
                lvwAllDocs.CellIcon(i, 1) = IndexForKey("CODE")
            Else
                lvwAllDocs.CellText(i, 2) = sFileName 'sFileName
                lvwAllDocs.CellIcon(i, 1) = SetIcon(sFileName)
            End If
            lvwAllDocs.CellText(i, 1) = Form.DocumentCaption
            
            lvwAllDocs.CellTag(i, 1) = Form.DocID

            'lCount = lCount + 1
'        End If
    Next
    lvwAllDocs.Redraw = True
End Sub



Private Sub tbrMain_ButtonClick(ByVal lButton As Long)
    Dim sKey As String
    sKey = tbrMain.ButtonKey(lButton)
    lCurrentTab = lButton + 1
    tvProject.Visible = (sKey = "Project") And bBusy = False
    lblMsg.Visible = (sKey = "Project")
    lvwAllDocs.Visible = (sKey = "Docs")
    If (sKey = "Docs") Then ListDocs
    lvwSymbols.Visible = (sKey = "Insert")
    txtSymbolPath.Visible = (sKey = "Insert")

    picProjects_Resize
    If Visible Then
        On Error Resume Next
        Select Case sKey
        Case "Project"
            tvProject.SetFocus
        Case "Docs"
            lvwAllDocs.SetFocus
        Case "Insert"
            lvwSymbols.SetFocus
        End Select
    End If
    
    tbrMain.ButtonChecked("Docs") = (sKey = "Docs")
    tbrMain.ButtonChecked("Project") = (sKey = "Project")
    tbrMain.ButtonChecked("Insert") = (sKey = "Insert")
    SetCaption
    
End Sub

'Private Sub tbsTabs_TabClick(ByVal lTab As Long)
'
'    Dim i As Integer
'    Dim sKey As String
'    sKey = tbsTabs.TabKey(lTab)
'
'    tvProject.Visible = (sKey = "Project") And bBusy = False
'    lblMsg.Visible = (sKey = "Project")
'    lvwAllDocs.Visible = (sKey = "Docs")
'    If (sKey = "Docs") Then ListDocs
'    lvwSymbols.Visible = (sKey = "Insert")
'    txtSymbolPath.Visible = (sKey = "Insert")
'
'    picProjects_Resize
'    If Visible Then
'        On Error Resume Next
'        Select Case sKey
'        Case "Project"
'            tvProject.SetFocus
'        Case "Docs"
'            lvwAllDocs.SetFocus
'        Case "Insert"
'            lvwSymbols.SetFocus
'        End Select
'    End If
'    SetCaption
'End Sub

Private Sub OpenSelectedFile(bDefault As Boolean)
    Dim nType As Integer
    Dim sPath As String
    Dim i As Integer
    Dim lSelected As Long
    Dim sTag As String
    Dim m_cItem As ItemInfo
    If lCurrentTab = 1 Then
        With tvProject
            lSelected = .Selected
            If .ItemKey(lSelected) = "Project" Then Exit Sub
            If .ItemData(lSelected) = 0 Then Exit Sub
           
            If .ItemImage(lSelected) <> IndexForKey("FolderClosed") And .ItemImage(lSelected) <> IndexForKey("LIVEFOLDER") Then
               ' LSet m_cItem = m_cItems(.ItemData(.Selected))
                'sTag = sPaths(.ItemData(lSelected))
                'sType = Left$(sTag, 1)
                If .ItemData(lSelected) = -2 Then
                    ' a live file...
                    sPath = .ItemKey(lSelected)
                Else
                    nType = m_cItems(.ItemData(.Selected)).nType
                    sPath = m_cItems(.ItemData(.Selected)).sPath
                    If nType = vbwEmail Then sPath = "mailto:" & sPath
                End If
                If nType = vbwEmail Or nType = vbwURL Then bDefault = True
                If nType = vbwFile Or nType = vbwFolder Then sPath = pGetAbsolutePath(sPath)
                pOpenFile sPath, bDefault
            End If
        End With
    End If
End Sub
Private Function pGetRelativePath(sFile As String) As String
Dim sFilePath As String
Dim sFileName As String
Dim sNewPath As String
Dim sProjectRoot As String
    sProjectRoot = ProjectRoot
    If sProjectPath = "" Then
        pGetRelativePath = sFile
    Else
        sFilePath = GetFolder(sFile) & "\"
        sFileName = GetCaption(sFile)
        If InStr(1, sFilePath, ProjectRoot) Then
            'referenced file is above project folder...
            sFilePath = Replace(sFilePath, sProjectRoot, ".\")
        Else
            
            Do
                ' go through the actual location of the project
                ' each time the current directory does not match
                ' the actual file's path, substitute a ..\
                If InStr(1, sFilePath, sProjectRoot) = 0 Then
                    sNewPath = "..\" & sNewPath
                    sProjectRoot = Left$(sProjectRoot, InStrRev(sProjectRoot, "\") - 1)
                    DoEvents
                Else
                    ' then, get the remaining half of the path
                    ' ie \Downloads\VB\vbexplorer\, and all the ..\'s
                    sFilePath = Replace(sFilePath, sProjectRoot, "")
                    If sFilePath = "" Then
        
                    Else
                        sFilePath = Right$(sFilePath, Len(sFilePath) - 1) '& "\"
                    End If
                    ' finally, remove the first \ from sPath, and add the filename
                    sFilePath = sNewPath & sFilePath '& GetCaption(sFile)
                    Exit Do
                End If
            Loop
        End If
        'remove any trailing \
        If Right$(sFilePath, 1) = "\" Then sFilePath = Left$(sFilePath, Len(sFilePath) - 1)
        pGetRelativePath = sFilePath & "\" & sFileName
    End If
    
End Function
Private Function pGetAbsolutePath(ByVal sFile As String) As String
Dim sFileFolder As String
    If Left$(sFile, 2) = ".\" Then
        pGetAbsolutePath = ProjectRoot & Right$(sFile, Len(sFile) - 2)
    ElseIf Mid$(sFile, 2, 2) = ":\" Then
        pGetAbsolutePath = sFile
    ElseIf Left$(sFile, 3) = "..\" Then
        sFileFolder = ProjectRoot ' folder
        sFileFolder = Left$(sFileFolder, Len(sFileFolder) - 1)
        Do While Left$(sFile, 3) = "..\"
            ' up another folder
            If InStrRev(sFileFolder, "\") <> 0 Then sFileFolder = Left$(sFileFolder, InStrRev(sFileFolder, "\") - 1)
            sFile = Right$(sFile, Len(sFile) - 3)
        Loop
        pGetAbsolutePath = sFileFolder & "\" & sFile
    Else
        pGetAbsolutePath = sFile
    End If
End Function
Private Sub pOpenFile(sFile As String, bDefault As Boolean)
    If bDefault Then
        ShellFunc sFile
    ElseIf IsProject(sFile) Then
        OpenProject sFile
    Else
        cDocuments.LoadFile (sFile)
    End If
End Sub

Private Sub tvProject_AfterLabelEdit(hItem As Long, NewText As String, Cancel As Boolean)
Dim sRoot As String
Dim sNewRoot As String

    If tvProject.ItemImage(hItem) = IndexForKey("FolderClosed") Then
        'we will replace this with the new root
        sRoot = tvProject.ItemFullPath(hItem)
        m_cFolders(tvProject.ItemData(hItem)).sName = NewText
        sNewRoot = tvProject.ItemFullPath(hItem)
        sNewRoot = Left$(sNewRoot, InStrRev(sNewRoot, "\")) & NewText
        'correct the paths of all the items
        pCorrectPaths hItem, sRoot, sNewRoot
        bChanged = True
    ElseIf tvProject.ItemData(hItem) > 0 Then
        m_cItems(tvProject.ItemData(hItem)).sName = NewText
        bChanged = True
    Else
        Cancel = True
    End If
End Sub
Private Sub pCorrectPaths(hStart As Long, sRoot As String, sNewRoot As String)
Dim hCurItem As Long
Dim hLastItem As Long
Dim sTempRoot As String

    hCurItem = hStart
    'update data for each moved item
    Do
        If tvProject.ItemImage(hCurItem) = IndexForKey("FolderClosed") Then
            sTempRoot = Replace(tvProject.ItemFullPath(hCurItem), sRoot, sNewRoot)
            tvProject.ItemKey(hCurItem) = sTempRoot
            If Left$(sTempRoot, 8) = "Project\" Then sTempRoot = Right$(sTempRoot, Len(sTempRoot) - 8)
            If InStrRev(sTempRoot, "\") <> 0 Then sTempRoot = Left$(sTempRoot, InStrRev(sTempRoot, "\") - 1)
            If sTempRoot = m_cFolders(tvProject.ItemData(hCurItem)).sName Then sTempRoot = ""
            m_cFolders(tvProject.ItemData(hCurItem)).sPath = sTempRoot
        Else 'item
            sTempRoot = Replace(tvProject.ItemFullPath(hCurItem), sRoot, sNewRoot)
            sTempRoot = Left$(sTempRoot, InStrRev(sTempRoot, "\") - 1)
            If Left$(sTempRoot, 8) = "Project\" Then sTempRoot = Right$(sTempRoot, Len(sTempRoot) - 8)
            If sTempRoot = "Project" Then sTempRoot = ""
            m_cItems(tvProject.ItemData(hCurItem)).sFolder = sTempRoot
        End If
        'hDragItem=itemchild(
        'save it
        hLastItem = hCurItem
        hCurItem = tvProject.ItemChild(hLastItem) 'get its child
        If hCurItem = 0 And hCurItem <> hStart Then
            hCurItem = tvProject.ItemNextSibling(hLastItem) 'next item along
        End If
    Loop While hCurItem <> 0
End Sub
Private Sub tvProject_BeforeLabelEdit(hItem As Long, Cancel As Boolean)
'    If m_tProjectType <> tDevPad Then
'        Cancel = True
    If tvProject.ItemData(hItem) <= 0 Then
        Cancel = True
    End If
End Sub

Private Sub tvProject_Click(X As Long, Y As Long, RightButton As Boolean)
    If RightButton And bItemClicked = False Then
        ShowPopup "mnuProject", Me
    End If
    bItemClicked = False
End Sub

Private Sub tvProject_DragBegin(ByVal hItem As Long)
    'save the drag item
    hDragItem = hItem
End Sub

Private Sub tvProject_DragEnd(MoveItem As Boolean, hDropItem As Long)
Dim sRoot As String
Dim sNewRoot As String
Dim vResult As ShowYesNoResult

    If tvProject.ItemData(hDropItem) = -1 Or tvProject.ItemData(hDropItem) = -2 Or tvProject.ItemData(hDragItem) = -1 Then 'live folder item
        hDropItem = -1
        Exit Sub
    ElseIf tvProject.ItemData(hDropItem) = 0 Or tvProject.ItemImage(hDropItem) = IndexForKey("FolderClosed") Then
        'avoid errors for the check below
    ElseIf m_cItems(tvProject.ItemData(hDropItem)).nType = vbwLiveFolder Then
        hDropItem = -1
        Exit Sub
    End If
    If tvProject.ItemImage(hDropItem) <> IndexForKey("FolderClosed") Then
        If tvProject.ItemImage(hDropItem) <> IndexForKey("FILE_VBP") Then
            hDropItem = tvProject.ItemParent(hDropItem)
        End If
    End If
    If tvProject.ItemData(hDragItem) = -2 Then
        vResult = cDialog.ShowYesNo("Create a shortcut to this file?", True)
        If vResult = Yes Then
            AddItem tvProject.ItemKey(hDropItem), tvProject.ItemText(hDragItem), tvProject.ItemKey(hDragItem), vbwFile
        End If
        hDropItem = -1
        Exit Sub
    End If
    MoveItem = True
    'we will replace this with the new root
    sRoot = tvProject.ItemFullPath(hDragItem)
    'If sRoot = "" Then sRoot = m_cItems(tvProject.ItemData(hDragItem))
    sRoot = Left$(sRoot, InStrRev(sRoot, "\") - 1)
    sNewRoot = tvProject.ItemKey(hDropItem)
    If sRoot = sNewRoot Then
        hDropItem = -1
        Exit Sub
    End If
    If tvProject.ItemImage(hDragItem) = IndexForKey("LIVEFOLDER") Then
        DoEvents
        If Left$(sNewRoot, 8) = "Project\" Then sNewRoot = Right$(sNewRoot, Len(sNewRoot) - 8)
        If sNewRoot = "Project" Then sNewRoot = ""
        m_cItems(tvProject.ItemData(hDragItem)).sFolder = sNewRoot
    Else
        bChanged = True
        'order has changed...
        bOrderChanged = True
        'get the root path...
        
        'correct the paths of all the items
        pCorrectPaths hDragItem, sRoot, sNewRoot
    End If
End Sub
Public Property Get DevPadProject() As Boolean
    DevPadProject = (m_tProjectType = tDevPad)
End Property

Private Sub tvProject_ItemClick(hItem As Long, RightButton As Boolean)
    Dim sType As String
    Dim lpPos As POINTAPI
    Dim lPos As Long

    If RightButton Then
        tvProject.Selected = hItem
        bItemClicked = True

        'If m_tProjectType = tDevPad Then
            If tvProject.ItemKey(hItem) = "Project" Or tvProject.ItemData(hItem) = 0 Or tvProject.ItemImage(hItem) = IndexForKey("FolderClosed") Then
                If tvProject.ItemData(hItem) <> -1 Then
                    bItemClicked = False
                End If
                Exit Sub
            End If
            BuildPopupMenu 2
            If tvProject.ItemData(hItem) = -2 Then
                'a virtual item...
                HideProjectMenuItem "PrjPopSep1"
                HideProjectMenuItem "PrjPopEditPath"
                HideProjectMenuItem "PrjPopRenameItem"
                HideProjectMenuItem "PrjPopRemoveItem"
            Else
                Select Case m_cItems(tvProject.ItemData(hItem)).nType
                Case 0 ' file
    
                Case 1 ' folder
                    HideProjectMenuItem "PrjPopOpenFileDevPad"
                    SetProjectMenuCaption "PrjPopOpenFile", "Open Folder"
                    SetProjectMenuCaption "PrjPopEditPath", "Edit Folder Location"
                Case 2 ' URL
                    ' hide open in DevPad
                    HideProjectMenuItem "PrjPopOpenFileDevPad"
                    SetProjectMenuCaption "PrjPopOpenFile", "Open In Browser"
                    SetProjectMenuCaption "PrjPopEditPath", "Edit URL"
                Case 3 ' Email
                    ' hide open in DevPad
                    HideProjectMenuItem "PrjPopOpenFileDevPad"
                    SetProjectMenuCaption "PrjPopOpenFile", "Write Email"
                    SetProjectMenuCaption "PrjPopEditPath", "Edit Email Address"
                Case vbwLiveFolder
                    HideProjectMenuItem "PrjPopOpenFileDevPad"
                    HideProjectMenuItem "PrjPopOpenFile"
                    HideProjectMenuItem "PrjPopSep1"
                End Select
            End If
'        Else
'            BuildPopupMenu 2
'            If tvProject.ItemKey(hItem) = "Project" Or tvProject.ItemImage(hItem) = IndexForKey("FolderClosed") Then
'                ' is root item or category
'                If tvProject.ItemText(hItem) Like "*.vbp" Then
'                    ' is project
'                    HideProjectMenuItems
'                ElseIf tvProject.ItemText(hItem) Like "*.vbg" Then
'                    ' is project group
'                    HideProjectMenuItems
'                Else
'                    bItemClicked = False
'                    DeletePopupMenu 2
'                    Exit Sub
'                End If
'            Else
'                sType = "1"
'                ' is file
'                HideProjectMenuItem "PrjPopRenameItem"
'                HideProjectMenuItem "PrjPopOpenFile"
'            End If
'        End If
       ' If sType <> "" Then frmMainForm.ctlPopMenu.MenuDefault( = True
        ShowPopup "mnuProjectPopup", Me
        DeletePopupMenu 2
        
'        GetCursorPos lpPos
'        ScreenToClient hWnd, lpPos
'        frmMainForm.ctlPopMenu.ShowPopupMenu Me, "mnuProjectPopup", lpPos.x * Screen.TwipsPerPixelX, lpPos.y * Screen.TwipsPerPixelY
'        DeletePopupMenu 2
       ' frmMainForm.m_cPopMenu.ShowPopupAbsolute lpPos.x, lpPos.y
    End If
End Sub

Private Sub HideProjectMenuItem(sKey As String)
    frmMainForm.ctlPopMenu.RemoveItem sKey
End Sub
Private Sub SetProjectMenuCaption(sKey As String, sCaption As String)
Dim lIndex As Long
    'lIndex = frmMainForm.m_cPopMenu.IndexForKey(sKey)
    frmMainForm.ctlPopMenu.Caption(sKey) = sCaption
End Sub
Private Sub HideProjectMenuItems()
    With frmMainForm.ctlPopMenu
        .RemoveItem "PrjPopOpenFileDevPad"
        .RemoveItem "PrjPopOpenFile"
        .RemoveItem "PrjPopEditPath"
        .RemoveItem "PrjPopSep8"
    End With
End Sub

Private Sub picProjects_Resize()

On Error Resume Next
    With picProjects
        ' Project View
        If lCurrentTab = 1 Then
            tvProject.Move 0, 0, .ScaleWidth, .ScaleHeight '- 30
        ElseIf lCurrentTab = 4 Then
            lvwAllDocs.Move 0, 0, .ScaleWidth, .ScaleHeight
            If lvwAllDocs.Width > (lvwAllDocs.ColumnWidth(1) + 50) * Screen.TwipsPerPixelX Then _
                lvwAllDocs.ColumnWidth(2) = (lvwAllDocs.Width / Screen.TwipsPerPixelX) - lvwAllDocs.ColumnWidth(1) - 20
        ElseIf lCurrentTab = 2 Then
            txtSymbolPath.Width = .ScaleWidth
            txtSymbolPath.Left = 0
            txtSymbolPath.Top = 0
            lvwSymbols.Top = txtSymbolPath.Top + txtSymbolPath.Height + 40
            lvwSymbols.Width = .ScaleWidth
            lvwSymbols.Left = 0
            lvwSymbols.ColumnWidth(1) = (lvwSymbols.Width / Screen.TwipsPerPixelX) - 22
            lvwSymbols.Height = .ScaleHeight - (txtSymbolPath.Height + 30) - 15
        End If
        lblMsg.Width = .ScaleWidth
    End With
End Sub

Private Property Get ProjectName() As String
    ProjectName = pProjectInfo.sProjectName
End Property
Public Property Get ProjectPath() As String
    ProjectPath = sProjectPath
End Property
Private Property Let ProjectName(ByVal New_ProjectName As String)
    pProjectInfo.sProjectName = New_ProjectName
End Property

Public Function NewProject() As Boolean

On Error GoTo ErrHandler
    ' Add the root item
    If ClearProject(False) = False Then
        Exit Function
    End If
    tvProject.NoDragDrop = False
    m_tProjectType = tDevPad
'    m_bVBProject = False
'    m_bVBProjectGroup = False
'
    tvProject.Clear
    tvProject.Add 0&, FirstChild, "Project", LoadResString(1032), IndexForKey("FILE_VBP")  '24
    tvProject.Sorted("Project") = False 'True

    sProjectPath = ""
    pProjectInfo.sProjectName = GetSetting(REG_KEY, "Settings", "ProjectName", LoadResString(1032)) & " 1"
    pProjectInfo.sProjectAuthor = GetSetting(REG_KEY, "Settings", "ProjectAuthor", "Unknown")
    bChanged = False
    NewProject = True
    bProjectOpen = True
    SetCaption
    If Visible = False Then ShowWindow False
    Exit Function
ErrHandler:
    cDialog.ErrHandler Err, Error, "Project.NewProject"
End Function
Public Sub SaveProject(Optional bSaveAs As Boolean)
    SaveProjectFile sProjectPath, bSaveAs
End Sub

Public Sub AddFolder()
    Dim sResult As String
    sResult = cDialog.BrowseForFolder
    If sResult <> Empty Then
        NewItem sResult, GetCaption(sResult), vbwFolder
    End If
End Sub
Public Sub AddInternet(bURL As Boolean)
    Dim sURL As String
    Dim sPrompt As String
    If bURL Then
        sPrompt = "Hyperlink"
    Else
        sPrompt = "Email Address"
    End If
    sURL = cDialog.InputBox("Please enter a new " & sPrompt, "Add " & sPrompt)
    If sURL <> Empty Then
        If bURL Then
            NewItem sURL, sURL, vbwURL
        Else
            NewItem sURL, sURL, vbwEmail
        End If
    End If
End Sub


Public Function NewFolder(Optional sNewFolder As String = "", Optional bNoError As Boolean, Optional sParent As String = "") As String 'node
    'Dim sParent As Variant
   ' Dim nodParent As Node
    Dim sTest As String
    Dim bCreating As Boolean
    
    ' reset the file pos
    ' add the file, and the rest of the sing
    If sNewFolder = Empty Then
        bCreating = True
        sNewFolder = cDialog.InputBox("Enter the new folder name", "New Folder")
        If sNewFolder = Empty Then Exit Function
    End If
   ' Dim frmD As Node
    sParent = pGetValidParent(sParent)

    Dim lNewItem As Long
    Dim sFullPath As String
    Dim i As Long
    sFullPath = sParent & "\" & sNewFolder
    If tvProject.IsValidNewKey(sFullPath) Then
        lNewItem = tvProject.Add(sParent, AlphabeticalChild, sFullPath, sNewFolder, IndexForKey("FolderClosed"), IndexForKey("FolderOpen"))
        ReDim Preserve m_cFolders(1 To m_lFolderCount + 1)
        m_lFolderCount = m_lFolderCount + 1
        'LSet m_cItems(m_lFolderCount) = vData
       ' If Left$(sParent, 4) = "Project" Then sParent = Right$(sParent, Len(sParent) - 5)
        m_cFolders(m_lFolderCount).sPath = TrimPath(sParent)
        m_cFolders(m_lFolderCount).sName = sNewFolder
        tvProject.ItemData(lNewItem) = m_lFolderCount

    End If
    NewFolder = sFullPath
    If bCreating Then
        tvProject.SetFocus
        tvProject.EnsureVisible lNewItem
        'tvproject
       ' tvProject.ItemSelected(tvProject.Selected) = False
        tvProject.Selected = lNewItem
    End If
    bChanged = True
   
End Function

Public Function AddItem(ByVal sParent As String, ByVal sText As String, ByVal sPath As String, Optional ByVal ShortcutType As ShortcutTypes = vbwFolder, Optional bRealPath As Boolean) As Long
    Dim lNewItem As Long
    Dim sFullPath As String
    Dim tData As ItemInfo
    sFullPath = sParent & "\" & sText
    If tvProject.IsValidNewKey(sFullPath) Then
        ' Set the object
        lNewItem = tvProject.Add(sParent, LastChild, sFullPath, sText)
        'Set nodNew = tvProject.Nodes.Add(sParent, tvwChild, , sText)
        With tvProject
            tData.nType = ShortcutType
            tData.sFolder = sParent
            tData.sName = sText
            If ShortcutType = vbwFile Or ShortcutType = vbwFolder Then
                tData.sPath = pGetRelativePath(sPath)
            Else
                tData.sPath = sPath
            End If
            ' set the parent
            AddItemData lNewItem, ShortcutType, tData
            .ItemImage(lNewItem) = GetImage(ShortcutType, sPath)
        End With
        AddItem = lNewItem
    Else
        AddItem = tvProject.ItemIndex(sFullPath)
    End If
End Function
Public Sub AddFiles()
    OpenFile True
End Sub
Public Function NewItem(sPath As String, Optional sCaption As String, Optional ByVal ShortcutType As ShortcutTypes = vbwFile, Optional sParent As String) As Boolean
    Dim sFolderName

    If sCaption = Empty Then sCaption = GetCaption(sPath)
    sParent = pGetValidParent(sParent)
    AddItem sParent, sCaption, sPath, ShortcutType
    tvProject.EnsureVisible (tvProject.ItemIndex(sParent & "\" & sCaption))
    NewItem = True
    bChanged = True
End Function
Private Function pGetValidParent(ByVal sParent As String) As String
    If sParent = "" Then
        If tvProject.Selected = 0 Or tvProject.ItemParent(tvProject.Selected) = 0 Then
            sParent = "Project"
        Else
            If tvProject.ItemImage(tvProject.Selected) <> IndexForKey("FOLDERCLOSED") Then
                sParent = tvProject.ItemKey(tvProject.ItemParent(tvProject.Selected))
            Else
                sParent = tvProject.ItemKey(tvProject.Selected)
            End If
        End If
    End If
    pGetValidParent = sParent
End Function
    
Public Function DeleteItem() As Boolean
    Dim sTest As String
    Dim sPrompt As String
    Dim sItemName As String
    With tvProject
        If .Selected = 0 Then Exit Function
        'On Error GoTo errhandler
        If .ItemKey(.Selected) = "Project" Then
            cDialog.ErrHandler vbObjectError + 1012, "Cannot delete root item", "Project.DeleteItem"
            DeleteItem = False
            Exit Function
        ElseIf .ItemImage(.Selected) = IndexForKey("FolderClosed") Then
            ' Current item is folder
            sPrompt = LoadResString(1272) '"folder and all its contents"
        Else
            ' Current item is file
            sPrompt = LoadResString(1271) '"file"
        End If
        '"Are you sure you want to delete the current "
        If cDialog.ShowYesNo(LoadResString(1274) & LoadResString(1269) & sPrompt & "?", False) = No Then
            DeleteItem = False
            ' No we do not want to
            Exit Function
        End If
        
        'move up items
        If .ItemImage(.Selected) = IndexForKey("FolderClosed") Then
            'is a folder...
            pRemoveArrayItem True, .ItemData(.Selected)
        ElseIf .ItemData(.Selected) > 0 Then
            pRemoveArrayItem False, .ItemData(.Selected)
        End If
        .Remove .Selected
    End With
    bChanged = True
    DeleteItem = True
End Function
Private Sub pRemoveArrayItem(bFolder As Boolean, lIndex As Long)
Dim i As Long
    If bFolder Then
        For i = lIndex To m_lFolderCount - 1
            LSet m_cFolders(i) = m_cFolders(i + 1)
        Next
        m_lFolderCount = m_lFolderCount - 1
        If m_lFolderCount = 0 Then
            Erase m_cFolders()
        Else
            ReDim Preserve m_cFolders(1 To m_lFolderCount)
        End If
    Else
        For i = lIndex To m_lItemCount - 1
            LSet m_cItems(i) = m_cItems(i + 1)
        Next
        m_lItemCount = m_lItemCount - 1
        If m_lItemCount = 0 Then
            Erase m_cItems()
        Else
            ReDim Preserve m_cItems(1 To m_lItemCount)
        End If
    End If
End Sub

Public Function ClearProject(bRaiseCloseEvent As Boolean) As Boolean
    If OKToClose = True Then
        On Error Resume Next
       ' tvProject.Nodes.Clear
        tvProject.Clear
        Reset
        ClearProject = True
        bProjectOpen = False
        SetCaption
        
        'Erase sPaths
        'ReDim sPaths(0 To 0)
        SaveSetting REG_KEY, "Settings", "LastProject", ""
        ReDim m_cFolders(0 To 0)
        ReDim m_cItems(0 To 0)
        Erase m_cFolders
        Erase m_cItems
        bOrderChanged = False
        bChanged = False
        m_lFolderCount = 0
        m_lItemCount = 0
        sProjectPath = ""
    Else
        ClearProject = False
    End If
End Function



Private Function SaveProjectAs() As Boolean
    Dim sFilter As String
    Dim i As Long
    Dim sOldPath As String
    'get the filter
    sFilter = "DevPad Project|*.dpp" '& GetProjectFilter()
    sOldPath = sProjectPath
    If cDialog.ShowOpenSaveDialog(True, "Save " & pProjectInfo.sProjectName & " As...", sFilter, , , 2) = False Then
        SaveProjectAs = False
    Else
        With CmDlg
            'save the filenames...
            sProjectPath = .FileName

            If m_tProjectType <> tDevPad And .FilterIndex = 1 Then
                'wasn't devpad project, but now it is
                m_tProjectType = tDevPad
            Else
                'no change...
            End If
            SaveSetting REG_KEY, "Settings", "LastProject", sProjectPath
            frmMainForm.UpdateFileMenu .FileName, 2
            SetCaption
            If sOldPath = "" Then
                SetStatusBar "Updating To Relative Paths..."
                'update items with new relative path
                For i = 1 To m_lItemCount
                    With m_cItems(i)
                        If .nType = vbwFile Or .nType = vbwFolder Then
                            .sPath = pGetRelativePath(.sPath)
                        End If
                    End With
                Next
                SetStatusBar ""
            End If
        End With
        SaveProjectAs = True
    End If
End Function
' gets the correct image for items, folders, urls etc
Private Function GetImage(ByVal nType As Long, ByVal sPath As String) As Long
    Select Case nType
    Case vbwFile
        ' set the image
        GetImage = SetIcon(sPath)
        
    Case vbwLiveFolder
        GetImage = IndexForKey("LIVEFOLDER")
    Case vbwFolder
'        If IsDrive(sPath) Then
'            GetImage = IndexForKey("Drive")
'        Else
            GetImage = IndexForKey("FolderShortcut")
        'End If
    Case vbwURL
        GetImage = IndexForKey("HYPERLINK")
    Case vbwEmail
        GetImage = IndexForKey("EMAIL")
    Case Else
        GetImage = "0"
    End Select
End Function

Private Function SetIcon(ByVal sFileName As String, Optional bCheck As Boolean = True) As Long
    Dim sExtension As String
    Dim lIndex As Long
    Dim hIcon As Long
    
    ' Get the extension
    sExtension = GetExtension(sFileName)
    ' What extension?
    'Select Case LCase$(sExtension)
'    Case "htm", "html"
'        SetIcon = IndexForKey("FILE_HTM")
'    Case "vbp"
'        SetIcon = IndexForKey("FILE_VBP")
'    Case "c", "cpp", "cxx", "tli"
'        SetIcon = IndexForKey("FILE_CSOURCE")
'    Case "h", "hxx", "tlh", "inl"
'        SetIcon = IndexForKey("FILE_CHEADER")
'    Case "rc", "res"
'        SetIcon = IndexForKey("FILE_RES")
    'Case Else
    If bCheck Then
        If Dir(sFileName) = "" Then
            SetIcon = IndexForKey("ALERT")
        End If
    End If
    If SetIcon = 0 Then
        On Error Resume Next
        lIndex = vbalMain.ItemIndex("FILE_" & UCase$(sExtension))
        If Err = 0 And lIndex <> 0 Then
            SetIcon = lIndex
        Else
            Dim SH_INFO As SHFILEINFO
            SHGetFileInfo sFileName, 0, SH_INFO, Len(SH_INFO), SHGFI_ICON + SHGFI_SMALLICON
            hIcon = SH_INFO.hIcon
            If hIcon <> 0 Then
                vbalMain.AddFromHandle hIcon, 1, "FILE_" & UCase$(sExtension)
                SetIcon = IndexForKey("FILE_" & UCase$(sExtension))
            Else
                SetIcon = IndexForKey("Misc")
            End If
        End If
    End If
    'End Select
End Function

' is a folder?
Private Function IsFolder(sFileName As String) As Boolean
    Select Case Len(sFileName) - InStrRev(sFileName, ".")
    Case 1, 2, 3, 4
        ' valid extension
        IsFolder = False
    Case Else
        IsFolder = True
    End Select
End Function
' updates the list of open documents
Public Sub UpdateDocs()
    If bNoUpdate Then Exit Sub
    'If tbsTabs.TabKey(tbsTabs.SelectedTab) = "Docs" Then ListDocs
    If lCurrentTab = 4 Then ListDocs
    'If (tbsTabs.SelectedItem.key = "Docs") Then ListDocs
End Sub


Public Sub NewCategory()
On Error GoTo ErrHandler
    Dim sNewCategory As String
    Dim intFileNum As Integer
    sNewCategory = cDialog.InputBox("Please enter a name for the new category", "New Category")
    If sNewCategory = "" Then Exit Sub
    MkDir App.Path & "\_insert" & sOldPath & "\" & sNewCategory
    ChangePath sOldPath, True
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Project.NewCategory"
End Sub
Public Sub DeleteCategory(Optional bSelected As Boolean)
On Error GoTo ErrHandler
    If sOldPath <> "" And sOldPath <> "\" Or bSelected Then
        Dim sPath As String
        sPath = sOldPath
        If bSelected Then sPath = sPath & "\" & lvwSymbols.CellText(lvwSymbols.SelectedRow, 1)
        'delete the folder (displays standard explorer prompt)
        DeleteFile App.Path & "\_insert" & sPath
        'unfortunately, we have no way of knowing if it was cancelled or not...
        'assume it wasn't!
        If bSelected Then
            ChangePath sOldPath, True
        Else
            UpFolder
        End If
    Else
        cDialog.ShowWarning "You cannot delete the root folder", "Project.DeleteCategory"
    End If
    DoEvents
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Project.DeleteCategory"
End Sub
'Public Function AddDirectory(ByVal sFolder, bIncSubDirs As Boolean, bCreateDirs As Boolean, sFilter As String, bOverwrite As Boolean, sParent As String)
'    bBusy = True
'    bCancel = False
'    WalkFilesAndDirs sFolder, bIncSubDirs, bCreateDirs, sFilter, bOverwrite, sParent 'tvProject.ItemKey(tvProject.Selected)
'    bBusy = False
'    bCancel = False
'End Function
'
'Private Function WalkFilesAndDirs(ByVal sStartDir As String, bIncSubDirs As Boolean, bCreateDirs As Boolean, sFilter As String, bOverwrite As Boolean, sStartItem As String)  'node
'    Dim lpFindFileData As WIN32_FIND_DATA, lFileHdl  As Long
'    Dim sTemp As String, sTemp2 As String, lRet As Long, iLastIndex  As Integer
'    Dim sPath As String
'    Dim nodTemp As String ' node
'    Dim sKey As String
'    Dim sCurDir As String
'    On Error GoTo ErrHandler
'    ' add trailing \ to start directory if required
'    If Right$(sStartDir, 1) <> "\" Then sStartDir = sStartDir & "\"
'   ' Set nodStart = tvProject.SelectedItem
'   ' nodStart.Selected = True
'    sKey = sStartItem 'tvProject.ItemKey(lStartItem) '.key
'    sStartDir = sStartDir & sFilter
'   ' lblMsg.Caption = "Indexing..."
'    ' get a file handle
'    lFileHdl = FindFirstFile(sStartDir, lpFindFileData)
'    If lFileHdl <> -1 Then
'        Do Until lRet = ERROR_NO_MORE_FILES
'            sPath = Left$(sStartDir, Len(sStartDir) - 4) & "\"
'            SetStatusBar "Indexing " & sPath
'            DoEvents
'            If bCancel Then Exit Do
'            ' if it is a directory
'            If (lpFindFileData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = vbDirectory Then
'                'Strip off null chars and format the sing
'                sTemp = StripTerminator(lpFindFileData.cFileName)
'                ' make sure it is not a reference
'                If sTemp <> "." And sTemp <> ".." Then
'                    'add it to the tree view. Store its path as its Key
'                    'frmExample.lstFolders.AddItem sTemp
'                    If bCreateDirs Then
'                        If bIncSubDirs Then
'                            sCurDir = NewFolder(sTemp, True, sKey)
'                        Else
'                            AddItem sKey, sTemp, sPath & sTemp, vbwFolder
'                        End If
'                    Else
'                        ' remain where we started
'                        sCurDir = sKey 'lStartItem
'                    End If
'                    If bIncSubDirs Then
'                        ' then walk this dir
'                        WalkFilesAndDirs sPath & sTemp, True, bCreateDirs, sFilter, bOverwrite, sCurDir
'                    End If
'                End If
'            ' if it is a file
'            Else
'                sTemp = StripTerminator(lpFindFileData.cFileName)
'                AddItem sKey, sTemp, sPath & sTemp, vbwFile
'            End If
'            ' based on the file handle iterate through all files and dirs
'            lRet = FindNextFile(lFileHdl, lpFindFileData)
'            If lRet = 0 Then Exit Do
'        Loop
'    End If
'    SetStatusBar
'    ' close the file handle
'    lRet = FindClose(lFileHdl)
'    Exit Function
'ErrHandler:
'    cDialog.ErrHandler Err, Error, "Project.WalkFilesAndDirs"
'End Function
Public Sub AddLiveFolder()
On Error GoTo ErrHandler
    Dim sFolder As String
    Dim sParent As String
    sFolder = cDialog.BrowseForFolder(ProjectRoot)
    If sFolder <> "" Then
        sParent = pGetValidParent(sParent)
        pAddLiveFolder sFolder, GetCaption(sFolder), sParent
    End If
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Project.AddLiveFolder"
End Sub
Private Function pAddLiveFolder(sPath As String, sName As String, sParent As String) As Long
'    Dim bIncSubFolders As Boolean
'    Dim bCreateDirs As Boolean
'    Dim bOverwrite As Boolean
'    Dim sFilter As String
'    'Dim lstItem As ListItem
'    Dim bOptions As Variant
'    Dim sOptions As String
'    Dim i As Long
'    With frmSelectFilesFolders
'        SendMessage tvProject.hWnd, WM_SETREDRAW, False, 0&
'        tvProject.Visible = False
'        lblMsg.Refresh
'        Screen.MousePointer = 11
'        .Hide
'        For i = 1 To .lvwFiles.Rows
'            sOptions = .lvwFiles.CellTag(i, 1) ' lstItem.Tag
'
'            bOverwrite = (Mid$(sOptions, 3, 1) = "1") ' bOptions(3)
'            If .lvwFiles.CellIcon(i, 1) = IndexForKey("FolderClosed") Then
'                ' is a folder
'                bIncSubFolders = (Mid$(sOptions, 1, 1) = "1") 'blnOptions(1)
'                bCreateDirs = (Mid$(sOptions, 2, 1) = "1") 'blnOptions(2)
'                AddDirectory .lvwFiles.CellText(i, 1), bIncSubFolders, bCreateDirs, .lvwFiles.CellText(i, 3), bOverwrite, .lvwFiles.CellText(i, 4)
'            Else
'
'                ' is a file
'    AddItem "PRoject", "xxxx", "fred", vbwFile
'    AddItem "PRoject", "aaaa", "fred2", vbwFile
    Dim lItem As Long
    lItem = AddItem(sParent, sName, sPath, vbwLiveFolder)
    With tvProject
        .ItemPlusMinus(lItem) = True
        .Sorted(lItem) = False
    End With
    pAddLiveFolder = lItem
    bChanged = True
'            End If
'        Next
'        ShowProjectTV
'        SendMessage tvProject.hWnd, WM_SETREDRAW, True, 0&
'        Unload frmSelectFilesFolders
'        Screen.MousePointer = 0
'    End With
End Function
Public Sub SelectTab(sTabName As String)
   ' tbsTabs.SelectTab (sTabName)
    tbrMain.RaiseButtonClick tbrMain.ButtonIndex(sTabName)
    'Call tbsTabs_TabClick(tbsTabs.IndexForTab(sTabName))
End Sub
'Public Sub ViewVBProjectSource()
'
'    If OKToClose Then
'        If sProjectPath = "" Or LCase$(sProjectPath) = "()newproject" Then
'            cDialog.ErrHandler , "This project has not been saved", "ShowProjectSource", , , True
'            Exit Sub
'        End If
'        cDialog.ErrHandler , "You will need to reload the project if the project source has changed", "ShowProjectSource", , , True
'
'        LoadFile sProjectPath
'    End If
'End Sub

Private Sub tvProject_ItemDblClick(hItem As Long)
    OpenSelectedFile False
End Sub

Private Sub tvProject_ItemExpandingCancel(hItem As Long, ExpandType As vbwTreeView.ExpandTypeConstants, Cancel As Boolean)
Dim sPath As String
    If tvProject.ItemData(hItem) = 0 Then Exit Sub
    If tvProject.ItemData(hItem) = -1 Then
        sPath = tvProject.ItemKey(hItem) & "\"
    ElseIf tvProject.ItemImage(hItem) = IndexForKey("FOLDERCLOSED") Then
        m_cFolders(tvProject.ItemData(hItem)).bExpanded = Not (ExpandType = Collapse)
        Exit Sub
    ElseIf m_cItems(tvProject.ItemData(hItem)).nType = vbwLiveFolder Then
        sPath = m_cItems(tvProject.ItemData(hItem)).sPath & "\"
    End If
    If sPath <> "" Then
    'If tvProject.ItemData(hItem) = -1 Or tvProject.ItemData(hItem) = -2 And ExpandType <> Collapse Then
        If tvProject.ItemChild(hItem) = 0 Then
            'no existing items
            
            'Dim sTemp2 As String', lRet As Long, iLastIndex  As Integer
            'Dim sPath As String
            'Dim nodTemp As String ' node
            'Dim sKey As String
            'Dim sCurDir As String
            'Dim sStartDir As String
            'Dim lItem As Long
            Dim bHasFolders As Boolean
            Dim bHasFiles As Boolean
            tvProject.Visible = False
            ' add \ to start directory if required
            'sStartDir = m_cItems(tvProject.ItemData(hItem)).sPath & "\"
            sPath = pGetAbsolutePath(sPath)
            bHasFolders = WalkDir(sPath, True, hItem)
            bHasFiles = WalkDir(sPath, False, hItem)
            If bHasFiles = False And bHasFolders = False Then
                tvProject.ItemPlusMinus(hItem) = False
            End If
            tvProject.Visible = True
        End If
        Cancel = False
    End If
End Sub
Private Function pAddLiveItem(bFile As String, sText As String, sPath As String, sParent As String) As Long
    Dim lItem As Long
    With tvProject
        If tvProject.IsValidNewKey(sPath) Then
            ' Set the object
            lItem = .Add(sParent, LastChild, sPath, sText)
            If bFile Then
                .ItemData(lItem) = -2
                .ItemImage(lItem) = SetIcon(sPath, False)
            Else
                .ItemData(lItem) = -1
                .ItemImage(lItem) = IndexForKey("FOLDERCLOSED")
                .ItemSelectedImage(lItem) = IndexForKey("FOLDEROPEN")
                .ItemPlusMinus(lItem) = True
                .Sorted(lItem) = False
            End If
            pAddLiveItem = lItem
        End If
    End With
End Function
Private Function WalkDir(sStartDir As String, bFolder As Boolean, hItem As Long) As Boolean
Dim bHasFiles As Boolean
Dim lpFindFileData As WIN32_FIND_DATA, lFileHdl  As Long
Dim lRet As Long
Dim lItem As Long
Dim sTemp As String
    If bFolder Then
        lpFindFileData.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY
    End If
' get a file handle
    lFileHdl = FindFirstFile(sStartDir & "*.*", lpFindFileData)
    If lFileHdl <> -1 Then
        Do Until lRet = ERROR_NO_MORE_FILES
            'sPath = Left$(sStartDir, Len(sStartDir) - 4) & "\"
            'SetStatusBar "Indexing " & sPath
            'DoEvents
            'If bCancel Then Exit Do
            ' if it is a directory
            If (lpFindFileData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = vbDirectory Then
                'Strip off null chars and format the sing
                sTemp = StripTerminator(lpFindFileData.cFileName)
                ' make sure it is not a reference
                If sTemp <> "." And sTemp <> ".." And bFolder Then
                    WalkDir = True
                    'add it to the tree view. Store its path as its Key
                    'frmExample.lstFolders.AddItem sTemp
                    lItem = pAddLiveItem(False, sTemp, sStartDir & sTemp, tvProject.ItemKey(hItem))
                    tvProject.ItemPlusMinus(lItem) = True
                End If
            ' if it is a file
            ElseIf bFolder = False Then
                WalkDir = True
                sTemp = StripTerminator(lpFindFileData.cFileName)
                'AddItem tvProject.ItemKey(hItem), sTemp, sStartDir & sTemp, vbwTempLiveFile
                pAddLiveItem True, sTemp, sStartDir & sTemp, tvProject.ItemKey(hItem)
            End If
            ' based on the file handle iterate through all files and dirs
            lRet = FindNextFile(lFileHdl, lpFindFileData)
            If lRet = 0 Then Exit Do
        Loop
    End If
    ' close the file handle
    lRet = FindClose(lFileHdl)
End Function


Private Sub ShowProjectTV()
    tvProject.Visible = (lCurrentTab = 1) '(tbsTabs.TabKey(tbsTabs.SelectedTab) = "Project")
End Sub
Friend Sub AddItemTag(ByVal Item As Variant, ByVal sData As String)
    lPathsCount = UBound(sPaths) + 1
    ReDim Preserve sPaths(0 To lPathsCount)
    sPaths(lPathsCount) = sData
    tvProject.ItemData(Item) = lPathsCount
   ' lPathsCount = lPathsCount + 1
End Sub
Private Sub AddItemData(ByVal lItem As Long, ByVal nType As Integer, ByRef vData As ItemInfo)
   ' lPathsCount = UBound(sPaths) + 1
   ' ReDim Preserve sPaths(0 To lPathsCount)
  '  If nType = vbwFolder Then
    'Dim sPath As String
    vData.sFolder = TrimPath(vData.sFolder)
    ReDim Preserve m_cItems(1 To m_lItemCount + 1)
    m_lItemCount = m_lItemCount + 1
    LSet m_cItems(m_lItemCount) = vData
 '   Debug.Print m_cItems(m_lItemCount).sPath
    tvProject.ItemData(lItem) = m_lItemCount
   ' lPathsCount = lPathsCount + 1
End Sub
Private Function TrimPath(ByVal sPath As String) As String
    If sPath = "Project" Then
        TrimPath = ""
    ElseIf Left$(sPath, 8) = "Project\" Then
        TrimPath = Right$(sPath, Len(sPath) - 8)
    Else
        TrimPath = sPath
    End If
    
End Function
Friend Function ReadItemTag(ByVal Item As Variant) As String
    ReadItemTag = sPaths(tvProject.ItemData(Item))
End Function
Private Sub VBProjects_Init()
    'If cVBProjects Is Nothing Then Set cVBProjects = New clsVBProjects
End Sub

Public Function SaveInsertFile(ByVal sFile As String)
Dim i As Long
Dim nFileNum As Integer
Dim sHeader As String
Dim nHeaderLen As Integer
On Error Resume Next
    Kill sFile
    Err.Clear
On Error GoTo ErrHandler
    nFileNum = FreeFile
    Open sFile For Binary Access Write Lock Read As #nFileNum

    Put #nFileNum, , "devpadinsert"
    sHeader = App.Major & App.Minor & "." & App.Revision
    nHeaderLen = Len(sHeader)
    Put #nFileNum, , nHeaderLen
    Put #nFileNum, , sHeader
    Put #nFileNum, , m_lInsertsCount
    For i = 1 To m_lInsertsCount
        Put #nFileNum, , m_cInserts(i)
    Next i
    Close #nFileNum
    Exit Function
ErrHandler:
    cDialog.ErrHandler Err, Error, "Project.SaveInsertFile"
    Close #nFileNum
End Function
Public Function ReadInsertFile(ByVal sFile As String) As Boolean
Dim i As Long
Dim nFileNum As Integer
Dim sBuf As String
Dim lCount As Long
Dim bFail As Boolean
Dim sError As String
Dim sHeader As String
Dim nHeaderLen As Integer

On Error GoTo ErrHandler

    nFileNum = FreeFile
    
    Open sFile For Binary Access Read Lock Write As #nFileNum
    Erase m_cInserts
    m_lInsertsCount = 0
    sBuf = Space$(12)
    Get #nFileNum, , sBuf
    If sBuf = "devpadinsert" Then
        
    
        Get #nFileNum, , nHeaderLen
        sHeader = Space(nHeaderLen)
        Get #nFileNum, , sHeader
        Get #nFileNum, , m_lInsertsCount
        If m_lInsertsCount > 0 Then
            'm_iStoreCount = lCount
            ReDim m_cInserts(1 To m_lInsertsCount)
            
            For i = 1 To m_lInsertsCount
               Get #nFileNum, , m_cInserts(i)
            Next i
        End If
        Close #nFileNum
    Else
        Close #nFileNum
        'cDialog.ErrHandler 9, "The file " & sFile & " is an invalid format, or does not exist. It will be overwritten", "Project.ReadInsertFile"
        SaveInsertFile (sFile)
    End If
    ReadInsertFile = True
    
    Exit Function
ErrHandler:
    If Err = 76 Then
        If sFile <> App.Path & "\_insert\index.ins" Then
            sFile = App.Path & "\_insert\index.ins"
            Resume
        Else
            cDialog.ErrHandler Err, "Unable to find code library", "Project.ReadInsertFile"
        End If
    End If
    cDialog.ErrHandler Err, Error, "Project.ReadInsertFile"
    Close #nFileNum
    
End Function
Public Sub InsertSymbols()
    ChangePath "\Symbols"
    SelectTab "Insert"
End Sub
Private Sub ChangePath(sPath As String, Optional bRefresh As Boolean = False)
On Error GoTo ErrHandler
    If sPath = sOldPath And bRefresh = False Then Exit Sub
    ' Save Changes
    If sOldPath <> "-1" And sPath <> "\nosave" Then
        'save
        SaveInsertFile (App.Path & "\_insert" & sOldPath & "\index.ins")
    End If
    If sPath = "\nosave" Then sPath = ""
    'if we are exiting, don't re-load
    If bClosing Then Exit Sub
    If sPath = "\" Then
        sPath = ""
    ElseIf sPath <> "" Then
        If Left$(sPath, 2) = "\\" Then sPath = Right$(sPath, Len(sPath) - 1)
        If Left$(sPath, 1) <> "\" Then sPath = "\" & sPath
    End If
    'load item
    If LoadInsertPath(sPath) Then
        sOldPath = sPath
        txtSymbolPath.Text = sPath
    End If
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Project.ChangePath"
End Sub
Private Function LoadInsertPath(sPath As String) As Boolean
Dim lRow As Long
Dim i As Long
Dim lIndex As Long
    
    With lvwSymbols
        
'        If sPath = "\Symbols" Then
'            .Clear
'        'load the file
'            For i = 1 To 255
'                lRow = .Rows + 1
'                .AddRow , i
'                .CellText(lRow, 1) = Chr$(i)
'                '.CellIcon(lRow, 1) = IndexForKey("FILE_TXT")
'            Next
'            SortInsertList
'            LoadInsertPath = True
        If ReadInsertFile(App.Path & "\_insert" & sPath & "\index.ins") Then
            .Redraw = False
            .Clear
            .SortObject.Clear
            'add the 'up' item
            If sPath <> "" And sPath <> "\" Then
                lRow = .Rows + 1
                .AddRow , "&^UP"
                .CellText(lRow, 1) = "Up"
                .CellIcon(lRow, 1) = IndexForKey("FOLDERUP")
            End If
            'add the folders
            AddFolders (App.Path & "\_insert" & sPath)
            'add the items
            lIndex = IndexForKey("FILE_TXT")
            For i = 1 To m_lInsertsCount
                lRow = .Rows + 1
                .AddRow , i  'm_cInserts(i).sName
                .CellText(lRow, 1) = m_cInserts(i).sName
                .CellIcon(lRow, 1) = lIndex
            Next
'            For i = 1 To 255
'                lRow = .Rows + 1
'                .AddRow , i
'                .CellText(lRow, 1) = Chr$(i)
'                '.CellIcon(lRow, 1) = IndexForKey("FILE_TXT")
'            Next
            'm_lInsertsCount = 255
            If .Rows <> 0 Then
                .CellSelected(1, 1) = True
            End If
            
            SortInsertList
            .Redraw = True
            LoadInsertPath = True
        End If
    End With
End Function

Private Sub AddFolders(sPath As String)
    Dim lpFindFileData As WIN32_FIND_DATA, lFileHdl  As Long
    Dim sTemp As String, lRet As Long
    Dim lRow As Long
    On Error Resume Next
    
    ' get a file handle
    lFileHdl = FindFirstFile(sPath & "\*.*", lpFindFileData)
    If lFileHdl <> -1 Then
        Do Until lRet = ERROR_NO_MORE_FILES
            DoEvents
            'if it is a folder
            If (lpFindFileData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = vbDirectory Then
                'Strip off null chars and format the sing
                sTemp = StripTerminator(lpFindFileData.cFileName)
                ' make sure it is not a reference
                If sTemp <> "." And sTemp <> ".." Then
                    lRow = lvwSymbols.Rows + 1
                    'add it to the listview
                    lvwSymbols.AddRow , "&" & sTemp
                    lvwSymbols.CellText(lRow, 1) = sTemp
                    lvwSymbols.CellIcon(lRow, 1) = IndexForKey("FOLDERCLOSED")
                End If
            End If
            lRet = FindNextFile(lFileHdl, lpFindFileData)
            If lRet = 0 Then Exit Do
        Loop
    End If
    ' close the file handle
    lRet = FindClose(lFileHdl)
End Sub

Private Sub tvProject_KeyDown(KeyCode As Long)

    Select Case KeyCode
'    Case VK_DELETE 'del
'        DeleteItem
    Case 1179764
        If tvProject.ItemData(tvProject.Selected) = -1 Or tvProject.ItemImage(tvProject.Selected) = IndexForKey("LIVEFOLDER") Then pRefreshVirtualTree (tvProject.Selected)
'    Case VK_RETURN  'enter
'        tvProject_ItemDblClick (tvProject.Selected)
    End Select
End Sub

Private Sub txtSymbolPath_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ChangePath txtSymbolPath.Text
    End If
End Sub
Private Sub UpFolder()
Dim lPos As Long
Dim sNewPath As String
    'go up
    lPos = InStrRev(sOldPath, "\")
    sNewPath = Left$(sOldPath, lPos - 1)
    ChangePath (sNewPath)
End Sub












'Public Function SaveProjectFile(ByVal sFile As String, ByVal bSaveAs As Boolean)
'Dim i As Long
'Dim nFileNum As Integer
'Dim sHeader As String
'Dim nHeaderLen As Integer
'Dim lFolderIndex As Long
'Dim hItem As Long
'Dim hLastItem As Long
'Dim lLiveFolderIndex As Long
'    nFileNum = FreeFile
'    If sProjectPath = "()NEWPROJECT" Or bSaveAs Then
'        If SaveProjectAs = False Then Exit Function
'        sFile = sProjectPath
'    End If
'
'
'    If m_tProjectType <> tDevPad Then
'        VBProjects_Init
'        m_cProjectEx.SaveProject sProjectPath
'        'cVBProjects.SaveVBProject sProjectPath, m_bVBProjectGroup
'    Else
'        frmMainForm.ctlProg.Visible = True
'        frmMainForm.ctlProg.Text = "Saving Project..."
'        frmMainForm.ctlProg.Value = 0
'        'SetStatusBar "Saving Project..."
'        On Error Resume Next
'        'erase the existing file
'        Kill sFile
'        Err.Clear
'        On Error GoTo ErrHandler
'        'open the file...
'        Open sFile For Binary Access Write Lock Read As #nFileNum
'        'output identity header
'        Put #nFileNum, , "devpadproject"
'        sHeader = App.Major & App.Minor & "." & App.Revision
'        nHeaderLen = Len(sHeader)
'        Put #nFileNum, , nHeaderLen
'        Put #nFileNum, , sHeader
'        'save project information
'        With pProjectInfo
'            Put #nFileNum, , Len(.sProjectName)
'            Put #nFileNum, , .sProjectName
'            Put #nFileNum, , Len(.sProjectAuthor)
'            Put #nFileNum, , .sProjectAuthor
'            Put #nFileNum, , Len(.sProjectDescription)
'            Put #nFileNum, , .sProjectDescription
'        End With
'        If bOrderChanged = False Then
'            'output folders
'            Put #nFileNum, , m_lFolderCount
'            For i = 1 To m_lFolderCount
'                Put #nFileNum, , m_cFolders(i)
'                frmMainForm.ctlProg.Value = (i / (m_lFolderCount + m_lItemCount)) * 100
'            Next i
'
'        Else
'            'Dim i As Long
'            lFolderIndex = IndexForKey("FOLDERCLOSED")
'            lLiveFolderIndex = IndexForKey("LIVEFOLDER")
'            'output folders
'            Put #nFileNum, , m_lFolderCount
'
'            'we have to do things the long way... to ensure we
'            'list folders in correct order
'            hItem = tvProject.ItemChild("Project")
'            Dim hLastItem2 As Long
'            Do While hItem <> 0
'                If tvProject.ItemImage(hItem) = lFolderIndex And tvProject.ItemData(hItem) <> -1 Then
'                    i = i + 1
'                    'is a folder
'                    'output its item
'                    Put #nFileNum, , m_cFolders(tvProject.ItemData(hItem))
'                    frmMainForm.ctlProg.Value = (i / (m_lFolderCount + m_lItemCount)) * 100
'                End If
'                'save it
'                hLastItem = hItem
'                hItem = tvProject.ItemChild(hLastItem) 'get its child
'                If hItem = 0 Or tvProject.ItemImage(hLastItem) = lLiveFolderIndex Then
'                    hItem = tvProject.ItemNextSibling(hLastItem) 'next item along
'                    If hItem = 0 Then
'                        hLastItem2 = hLastItem
'                        Do While hItem = 0
'                            hItem = tvProject.ItemParent(hLastItem2)
'                            If tvProject.ItemKey(hItem) = "Project" Then
'                                'Exit Do
'                                GoTo ExitLoop
'                            End If
'                            hLastItem2 = hItem
'                            hItem = tvProject.ItemNextSibling(hItem)
'                            'If tvProject.ItemData(hItem) = -1 Then hItem = 0
'                            'If hItem = 0 Then Stop
'                        Loop
'                       ' hLastItem = hItem
'                    End If
'                    'If hItem = 0 Then
'                    '    hItem = tvProject.ItemParent
'                End If
'            Loop
'ExitLoop:
'            If i <> m_lFolderCount Then Stop
'        End If
'        'output items
'        Put #nFileNum, , m_lItemCount
'        For i = 1 To m_lItemCount
'            Put #nFileNum, , m_cItems(i)
'            frmMainForm.ctlProg.Value = ((m_lFolderCount + i) / (m_lFolderCount + m_lItemCount)) * 100
'        Next i
'        'tidy up
'        Close #nFileNum
'    End If
'    bChanged = False
'    bOrderChanged = False
'    frmMainForm.ctlProg.Visible = False
'    Exit Function
'ErrHandler:
'    cDialog.ErrHandler Err, Error, "Project.SaveInsertFile"
'    Close #nFileNum
'
'End Function
Public Function SaveProjectFile(ByVal sFile As String, ByVal bSaveAs As Boolean) As Boolean
Dim i As Long
Dim nFileNum As Integer
Dim sHeader As String
Dim nHeaderLen As Integer
Dim lFolderIndex As Long
Dim hItem As Long
Dim hLastItem As Long
Dim lLiveFolderIndex As Long
Dim sOut As String
    nFileNum = FreeFile
    If sProjectPath = "" Or bSaveAs Or m_tProjectType <> tDevPad Then
        If m_tProjectType <> tDevPad Then
            '"Developers Pad does not currently support saving projects in this format"
            cDialog.ErrHandler vbObjectError + 2000, LoadResString(1287), "ProjectVC.SaveProject"
        End If
        If SaveProjectAs = False Then Exit Function
        sFile = sProjectPath
    End If

    If m_tProjectType <> tDevPad Then
'        VBProjects_Init
'        'not supported...
'        'm_cProjectEx.SaveProject (sProjectPath)
'        '"Developers Pad does not currently support saving projects in this format"
'        cDialog.ErrHandler vbObjectError + 2000, LoadResString(1287), "ProjectVC.SaveProject"
'        SaveProjectAs
'        'cVBProjects.SaveVBProject sProjectPath, m_bVBProjectGroup
    Else
        frmMainForm.ctlProg.Visible = True
        frmMainForm.ctlProg.Text = "Saving Project..."
        frmMainForm.ctlProg.Value = 0
        'SetStatusBar "Saving Project..."
        On Error Resume Next
        'erase the existing file
        Kill sFile
        Err.Clear
        On Error GoTo ErrHandler
        'open the file...
        Open sFile For Output Lock Read Write As #nFileNum
        sOut = "<DevelopersPadProject>" & vbCrLf & _
               "  <Settings" & vbCrLf & _
               "    Name = """ & pProjectInfo.sProjectName & """" & vbCrLf & _
               "    Author = """ & pProjectInfo.sProjectAuthor & """" & vbCrLf & _
               "    Version = ""1.0""" & vbCrLf & _
               "    Root = """ & m_sProjectRoot & """" & vbCrLf & _
               "  >" & vbCrLf & _
               "    <Description>" & vbCrLf & _
               "    " & pProjectInfo.sProjectDescription & vbCrLf & _
               "    </Description>" & vbCrLf & _
               "  </Settings>"
        Print #nFileNum, sOut
        Print #nFileNum, "  <Folders>"
        If bOrderChanged = False Then
            For i = 1 To m_lFolderCount
                sOut = "    <Folder" & vbCrLf & _
                         "      Name = """ & m_cFolders(i).sName & """" & vbCrLf & _
                         "      Path = """ & m_cFolders(i).sPath & """" & vbCrLf & _
                         "      Expanded = """ & IIf(m_cFolders(i).bExpanded, "1", "0") & """" & vbCrLf & _
                         "    />"
                Print #nFileNum, sOut
                frmMainForm.ctlProg.Value = (i / (m_lFolderCount + m_lItemCount)) * 100
            Next i
        Else
            lFolderIndex = IndexForKey("FOLDERCLOSED")
            lLiveFolderIndex = IndexForKey("LIVEFOLDER")
            'we have to do things the long way... to ensure we
            'list folders in correct order
            hItem = tvProject.ItemChild("Project")
            Dim hLastItem2 As Long
            Do While hItem <> 0
                If tvProject.ItemImage(hItem) = lFolderIndex And tvProject.ItemData(hItem) <> -1 Then
                    i = i + 1
                    'is a folder
                    'output its item
                    With m_cFolders(tvProject.ItemData(hItem))
                        sOut = "    <Folder" & vbCrLf & _
                         "      Name = """ & .sName & """" & vbCrLf & _
                         "      Path = """ & .sPath & """" & vbCrLf & _
                         "    />"
                    End With
                    Print #nFileNum, sOut
                    frmMainForm.ctlProg.Value = (i / (m_lFolderCount + m_lItemCount)) * 100
                End If
                'save it
                hLastItem = hItem
                hItem = tvProject.ItemChild(hLastItem) 'get its child
                If hItem = 0 Or tvProject.ItemImage(hLastItem) = lLiveFolderIndex Then
                    hItem = tvProject.ItemNextSibling(hLastItem) 'next item along
                    If hItem = 0 Then
                        hLastItem2 = hLastItem
                        Do While hItem = 0
                            hItem = tvProject.ItemParent(hLastItem2)
                            If tvProject.ItemKey(hItem) = "Project" Then
                                'finished...
                                GoTo ExitLoop
                            End If
                            hLastItem2 = hItem
                            hItem = tvProject.ItemNextSibling(hItem)
                        Loop
                    End If
                End If
            Loop
ExitLoop:
            'If i <> m_lFolderCount Then Stop
        End If
        Print #nFileNum, "  </Folders>"
        Print #nFileNum, "  <Files>"
        For i = 1 To m_lItemCount
            With m_cItems(i)
                sOut = "    <File" & vbCrLf & _
                     "      Name = """ & .sName & """" & vbCrLf & _
                     "      Folder = """ & .sFolder & """" & vbCrLf & _
                     "      RelPath = """ & .sPath & """" & vbCrLf & _
                     "      Type = """ & .nType & """" & vbCrLf & _
                     "    />"
            End With
            Print #nFileNum, sOut
            frmMainForm.ctlProg.Value = (i / (m_lFolderCount + m_lItemCount)) * 100
        Next i
        Print #nFileNum, "  </Files>"
        Print #nFileNum, "</DevelopersPadProject>"
        'tidy up
        Close #nFileNum
    End If
    bChanged = False
    bOrderChanged = False
    frmMainForm.ctlProg.Visible = False
    SaveProjectFile = True
    Exit Function
ErrHandler:
    cDialog.ErrHandler Err, Error, "Project.SaveInsertFile"
    Close #nFileNum

End Function
Private Sub SetProjectType(sFileName As String)
    Dim sExtension As String
    'get the file's extension, and set project mode
    'accordingly
    sExtension = LCase$(GetExtension(sFileName))
    Select Case sExtension
    Case "dpp"
        m_tProjectType = tDevPad
    Case "vbp"
        m_tProjectType = tVB
    Case "vbg"
        m_tProjectType = tVBGroup
    Case "dsp"
        m_tProjectType = tVC
'    Case "csproj"
'        m_tProjectType = tCnet
    Case "csproj", "vbproj"
        m_tProjectType = tVBnet
    End Select
End Sub
Private Function GetProjectFilter() As String
    Select Case m_tProjectType
    Case tDevPad
    Case tVB
        GetProjectFilter = "|VB Project|*.vbp"
    Case tVBGroup
        GetProjectFilter = "|VB Project Group|*.vbg"
    Case tVC
        GetProjectFilter = "|Visual C++ Projects|*.dsp"
    Case tVBnet
        GetProjectFilter = "|Visual Studio.net Projects|*.vbproj;*.csproj"
'    Case tCnet
'        GetProjectFilter = "|C# Projects|*.csproj"
    End Select
End Function

Private Sub ProjectExInit()
    If Not m_cProjectEx Is Nothing Then
        If m_cProjectEx.ProjectType <> m_tProjectType Then
            'need to change...
            Set m_cProjectEx = Nothing
        End If
    End If
    If m_cProjectEx Is Nothing Then
        Select Case m_tProjectType
        Case tDevPad
        Case tVB
            Set m_cProjectEx = New clsProjectVB
        Case tVBGroup
            Set m_cProjectEx = New clsProjectVBGroup
        Case tVC
            Set m_cProjectEx = New clsProjectVC
        Case tVBnet
            Set m_cProjectEx = New clsProjectVBnet
'        Case tCnet
'            Set m_cProjectEx = New clsProjectCnet
        End Select
    End If
End Sub
'Public Function OpenProject(ByVal sFile As String) As Boolean
'Dim rc As RECT
'Dim mhwnd As Long
'Dim i As Long
'Dim nFileNum As Integer
'Dim sBuf As String
'Dim lCount As Long
'Dim bFail As Boolean
'Dim sError As String
'Dim sHeader As String
'Dim nHeaderLen As Integer
'Dim lFolderIcon As Long
'Dim lFolderOpenIcon As Long
'Dim sName As String
'Dim sParent As String
'Dim lNewItem As Long
'Dim lLen As Long
'    On Error GoTo ErrHandler
'    If ClearProject(False) = False Then Exit Function
'    sProjectPath = sFile
'
'    SendMessage tvProject.hWnd, WM_SETREDRAW, False, 0&
'    With tvProject
'        .Clear
'  '      .Remove "Project"
'        .Add 0&, FirstChild, "Project", LoadResString(1032), IndexForKey("FILE_VBP")  '24
'        SetProjectType sProjectPath
'        If m_tProjectType <> tDevPad Then
'            .NoDragDrop = True
'           ' tvProject.Add 0&, FirstChild, "Project", LoadResString(1032), IndexForKey("FILE_VBP") '24
'
'            ProjectExInit
'            m_cProjectEx.LoadProject sProjectPath
'            'ImportProject (sProjectPath)
'            pProjectInfo.sProjectName = GetCaption(sProjectPath)
'           ' m_bVBProject = True
'           ' m_bVBProjectGroup = (GetExtension(sProjectPath) = "vbg")
'           ' GoTo General
'
'          ' Exit Function
'        Else
'            .NoDragDrop = False
'
'            frmMainForm.ctlProg.Visible = True
'            frmMainForm.ctlProg.Text = "Loading Project..."
'
'            nFileNum = FreeFile
'            Open sFile For Binary Access Read Lock Write As #nFileNum
'            Erase m_cFolders
'            Erase m_cItems
'            m_lFolderCount = 0
'            m_lItemCount = 0
'
'            sBuf = Space$(13)
'            Get #nFileNum, , sBuf
'            If sBuf = "devpadproject" Then
'                Get #nFileNum, , nHeaderLen
'                sHeader = Space(nHeaderLen)
'                Get #nFileNum, , sHeader
'                'we don't worry about version info at the moment!
'                'If sHeader...
'
'                With pProjectInfo
'                    'read len
'                    Get #nFileNum, , lLen
'                    'fill buffer
'                    .sProjectName = Space(lLen)
'                    'read into buffer!
'                    Get #nFileNum, , .sProjectName
'                    'and so on...
'                    Get #nFileNum, , lLen
'                    .sProjectAuthor = Space(lLen)
'                    Get #nFileNum, , .sProjectAuthor
'                    Get #nFileNum, , lLen
'                    .sProjectDescription = Space(lLen)
'                    Get #nFileNum, , .sProjectDescription
'                End With
'                lFolderIcon = IndexForKey("FolderClosed")
'                lFolderOpenIcon = IndexForKey("FolderOpen")
'                'tvProject.Redraw
'                Get #nFileNum, , m_lFolderCount
'                If m_lFolderCount > 0 Then
'                    ReDim m_cFolders(1 To m_lFolderCount)
'                    For i = 1 To m_lFolderCount
'                        Get #nFileNum, , m_cFolders(i)
'                        ' folders
'                        sParent = m_cFolders(i).sPath
'                        If sParent = "" Then
'                            sParent = "Project"
'                        Else
'                            sParent = "Project\" & sParent
'                        End If
'                        sName = m_cFolders(i).sName
'                        If sName = "" Then Err.Raise vbObjectError + 2001, , "Invalid Project File Format"
'                        'add to treeview
''                        If tvProject.IsValidNewKey(sParent) Then
''                            'unfortunately, we have to allow for this!
''                            'and I haven't decided what to do yet!
''                            Stop
''                           ' tvProject.Add sParent, AlphabeticalChild, sParent, GetCaption(sParent), lFolderIcon, lFolderOpenIcon
''                        End If
'                      '  If sName = "Downloads" Then Stop
'                        lNewItem = tvProject.Add(sParent, AlphabeticalChild, sParent & "\" & sName, sName, lFolderIcon, lFolderOpenIcon)
'                        tvProject.ItemData(lNewItem) = i
'                        frmMainForm.ctlProg.Value = (i / (m_lFolderCount)) * 100
'                    Next i
'                End If
'                Get #nFileNum, , m_lItemCount
'                If m_lItemCount > 0 Then
'                    ReDim m_cItems(1 To m_lItemCount)
'                    For i = 1 To m_lItemCount
'                        Get #nFileNum, , m_cItems(i)
'                        ' folders
'                        sParent = m_cItems(i).sFolder
'                        If sParent = "" Then
'                            sParent = "Project"
'                        Else
'                            sParent = "Project\" & sParent
'                        End If
'
'                        sName = m_cItems(i).sName
'                        'add to treeview
'                        'tvProject.Add sParent, AlphabeticalChild, sParent & "\" & sName, sName, lFolderIcon, lFolderOpenIcon
'                        'lNewItem = tvProject.Add(sParent, LastChild, sParent & "\" & sName, sName, GetImage(m_cItems(i).nType, m_cItems(i).sPath))
'
'                        lNewItem = tvProject.Add(sParent, LastChild, "K" & CStr(i), sName, GetImage(m_cItems(i).nType, m_cItems(i).sPath))
'                        If m_cItems(i).nType = vbwLiveFolder Then tvProject.ItemPlusMinus(lNewItem) = True
'
'                        tvProject.ItemData(lNewItem) = i
'                        frmMainForm.ctlProg.Value = (i / (m_lItemCount)) * 100
'                    Next i
'                End If
'
'                Close #nFileNum
'            Else
'                Close #nFileNum
'                'cDialog.ErrHandler 9, "The file " & sFile & " is an invalid format, or does not exist. It will be overwritten", "Project.ReadInsertFile"
'                'SaveInsertFile (sFile)
'            End If
'
'        End If
'    End With
'Tidy:
'    frmMainForm.ctlProg.Visible = False
'    SendMessage tvProject.hWnd, WM_SETREDRAW, True, 0&
'
'    tvProject.Refresh
'
'    tvProject.Tag = " " & LoadResString(1032) & " - " & pProjectInfo.sProjectName
'    tvProject.ItemExpanded("Project") = True
'
'    bProjectOpen = True
'    SaveSetting REG_KEY, "Settings", "LastProject", sProjectPath
'    bChanged = False
'    frmMainForm.UpdateFileMenu sProjectPath, True
'    If Visible = False Then ShowWindow False
'    SetCaption
'    SelectTab "Project"
'    Exit Function
'ErrHandler:
'    cDialog.ErrHandler Err, Error, "Project.ReadInsertFile"
'    'Resume Next
'   Close #nFileNum
'   GoTo Tidy
'End Function
Public Function OpenProject(ByVal sFile As String) As Boolean
Dim i               As Long
Dim lFolderIcon     As Long
Dim lFolderOpenIcon As Long
Dim sName           As String
Dim sParent         As String
Dim lNewItem        As Long
Dim cXML            As clsXML
Dim lNodeIndex      As Long
Dim lParentIndex    As Long
Dim cNode           As Node
Dim lExpandFolders() As Long
Dim lExpandFolderCount As Long

    On Error GoTo ErrHandler
    If ClearProject(False) = False Then Exit Function
    ProjectRoot = ""
    sProjectPath = sFile
    
    SendMessage tvProject.hWnd, WM_SETREDRAW, False, 0&
    
    With tvProject
        .Clear
        'create root item
        .Add 0&, FirstChild, "Project", LoadResString(1032), IndexForKey("FILE_VBP")  '24
        'get the project type
        SetProjectType sProjectPath
        If m_tProjectType <> tDevPad Then
            'disable drop and drag
            '.NoDragDrop = True
            ProjectExInit
            m_cProjectEx.LoadProject sProjectPath
            pProjectInfo.sProjectName = GetCaption(sProjectPath)
            m_sProjectRoot = GetFolder(sProjectPath)
        Else
            If Dir$(sFile) = "" Then Err.Raise 53, "Project.Open"
            'Open "c:\fffff.htm" For Input As #1
            
            'set up the progress bar
            frmMainForm.ctlProg.Visible = True
            frmMainForm.ctlProg.Text = "Parsing XML..."
            'do the XML
            Set cXML = New clsXML
            With cXML
                .ParseXMLFile sFile
                frmMainForm.ctlProg.Text = "Creating Project Items..."
                lNodeIndex = cXML.Node("Settings").Index
                If lNodeIndex = 0 Then Err.Raise vbObjectError + 1002, "Project.Open", "Invalid Project File"

                With pProjectInfo
                    'fill buffer
                    .sProjectName = cXML.NodeAttribute(lNodeIndex, "Name")
                    .sProjectAuthor = cXML.NodeAttribute(lNodeIndex, "Author")
                    .sProjectDescription = cXML.NodeChild(lNodeIndex, "Description").Value
                    ProjectRoot = cXML.NodeAttribute(lNodeIndex, "Root")
                End With
                'get the image indexes...
                lFolderIcon = IndexForKey("FolderClosed")
                lFolderOpenIcon = IndexForKey("FolderOpen")
                'get the folder node
                lParentIndex = .Node("Folders").Index
                'get the num
                m_lFolderCount = .NodeChildCount(lParentIndex, "Folder")
                If m_lFolderCount > 0 Then
                    ReDim m_cFolders(1 To m_lFolderCount)
                    For i = 1 To m_lFolderCount
                        lNodeIndex = .NodeChildIndex(lParentIndex, "Folder", i)
                        
                        sName = .NodeAttribute(lNodeIndex, "Name")
                        With m_cFolders(i)
                            .sPath = cXML.NodeAttribute(lNodeIndex, "Path")
                            .sName = sName
                            '.bExpanded = -CBool(cXML.NodeAttribute(lNodeIndex, "Expanded"))
                            ' folders
                            sParent = .sPath
                            If sParent = "" Then
                                sParent = "Project"
                            Else
                                sParent = "Project\" & sParent
                            End If
                        End With
                        lNewItem = tvProject.Add(sParent, LastChild, sParent & "\" & sName, sName, lFolderIcon, lFolderOpenIcon)
                        tvProject.ItemData(lNewItem) = i
                        'what shall I do about this?
                        If .NodeAttribute(lNodeIndex, "Expanded") = "1" Then
                            'add to list of folders to expand
                            lExpandFolderCount = lExpandFolderCount + 1
                            ReDim Preserve lExpandFolders(1 To lExpandFolderCount)
                            lExpandFolders(lExpandFolderCount) = lNewItem
                        End If
                        frmMainForm.ctlProg.Value = (i / (m_lFolderCount)) * 100
                    Next
                End If
                'get the files node
                lParentIndex = .Node("Files").Index
                m_lItemCount = .NodeChildCount(lParentIndex, "File")
                
                If m_lItemCount > 0 Then
                    ReDim m_cItems(1 To m_lItemCount)
                    For i = 1 To m_lItemCount
                        lNodeIndex = .NodeChildIndex(lParentIndex, "File", i)
                        With m_cItems(i)
                            .nType = cXML.NodeAttribute(lNodeIndex, "Type")
                            .sFolder = cXML.NodeAttribute(lNodeIndex, "Folder")
                            .sName = cXML.NodeAttribute(lNodeIndex, "Name")
                            .sPath = cXML.NodeAttribute(lNodeIndex, "RelPath")
                            'parent...
                            sParent = .sFolder
                            If sParent = "" Then
                                sParent = "Project"
                            Else
                                sParent = "Project\" & sParent
                            End If
                            sName = .sName
                        End With

                        'add to treeview
                        lNewItem = tvProject.Add(sParent, LastChild, "K" & CStr(i), sName, GetImage(m_cItems(i).nType, pGetAbsolutePath(m_cItems(i).sPath)))
                        If m_cItems(i).nType = vbwLiveFolder Then tvProject.ItemPlusMinus(lNewItem) = True

                        tvProject.ItemData(lNewItem) = i
                        frmMainForm.ctlProg.Value = (i / (m_lItemCount)) * 100
                    Next i
                End If
            End With
            'expand the folders
            For i = 1 To lExpandFolderCount Step 1
                .ItemExpanded(lExpandFolders(i)) = True
            Next i
        End If
    End With
    
Tidy:
    

    
    frmMainForm.ctlProg.Visible = False
    SendMessage tvProject.hWnd, WM_SETREDRAW, True, 0&
    
    tvProject.Refresh
    
    tvProject.Tag = " " & LoadResString(1032) & " - " & pProjectInfo.sProjectName
    tvProject.ItemExpanded("Project") = True

    bProjectOpen = True
    SaveSetting REG_KEY, "Settings", "LastProject", sProjectPath
    bChanged = False
    frmMainForm.UpdateFileMenu sProjectPath, 2
    If Visible = False Then ShowWindow False
    SetCaption
    SelectTab "Project"
    Exit Function
ErrHandler:
    If Err = -2147208504 Then
        sParent = "Project"
        Resume
    Else
        cDialog.ErrHandler Err, Error, "Project.ReadInsertFile"
    End If
    GoTo Tidy
End Function

Public Sub AddAllFiles()
    Dim i As Long
    For i = 1 To cDocuments.Count
        If cDocuments.Item(i).Saved = True Then
            'add the current file to the project
            NewItem cDocuments.Item(i).FileName, cDocuments.Item(i).DocumentCaption
        End If
    Next
End Sub
Private Sub pRefreshVirtualTree(hItem As Long)
Dim bState As Boolean
    'we need to update virtual items
    bState = tvProject.ItemExpanded(hItem)
    tvProject.RemoveChildren (hItem)
    tvProject.ItemExpanded(hItem) = bState
End Sub
Private Function IndexForKey(sKey As String) As Long
   ' Dim i As Long
   If sKey = "-1" Then
    IndexForKey = -1
   Else
    IndexForKey = vbalMain.ItemIndex(UCase$(sKey))
   End If
End Function
