VERSION 5.00
Object = "{1F5999A2-1D0B-11D4-82CF-004005AAE138}#6.1#0"; "VBWTBA~1.OCX"
Object = "{017E002E-D7CC-11D2-8E21-44B10AC10000}#22.0#0"; "VBWGRID.OCX"
Begin VB.Form frmOpen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "1002"
      Height          =   360
      Left            =   5055
      TabIndex        =   2
      Top             =   3615
      Width           =   1200
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "1244"
      Default         =   -1  'True
      Height          =   360
      Left            =   5055
      TabIndex        =   1
      Top             =   3180
      Width           =   1200
   End
   Begin VB.ComboBox cboPath 
      Height          =   315
      ItemData        =   "frmOpen.frx":000C
      Left            =   3375
      List            =   "frmOpen.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   90
      Width           =   2565
   End
   Begin VB.PictureBox picFavLabel 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   90
      ScaleHeight     =   300
      ScaleWidth      =   1050
      TabIndex        =   5
      Top             =   4200
      Width           =   1050
   End
   Begin VB.PictureBox picOpen 
      BorderStyle     =   0  'None
      Height          =   3540
      Index           =   0
      Left            =   945
      ScaleHeight     =   3540
      ScaleWidth      =   6345
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   825
      Width           =   6345
      Begin VB.ListBox lstTemplate 
         Height          =   840
         Left            =   45
         TabIndex        =   12
         Top             =   330
         Width           =   2400
      End
      Begin VB.TextBox txtPreview 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   45
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   9
         Text            =   "frmOpen.frx":0010
         Top             =   1425
         Width           =   4560
      End
      Begin VB.Label lblDescription 
         Height          =   1050
         Left            =   2505
         TabIndex        =   11
         Top             =   330
         Width           =   2070
      End
      Begin VB.Label lblLabel 
         Caption         =   "1165"
         Height          =   300
         Index           =   3
         Left            =   75
         TabIndex        =   10
         Top             =   75
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Preview..."
         Height          =   180
         Left            =   60
         TabIndex        =   8
         Top             =   1170
         Width           =   1725
      End
   End
   Begin VB.PictureBox picOpen 
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
      Height          =   3780
      Index           =   1
      Left            =   555
      ScaleHeight     =   3780
      ScaleWidth      =   6345
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1065
      Width           =   6345
   End
   Begin DevPad.ctlFrame ctlFrame1 
      Height          =   390
      Left            =   45
      TabIndex        =   13
      Top             =   60
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   688
      Begin vbwTBar.cToolbar tbrMain 
         Index           =   0
         Left            =   870
         Top             =   150
         _ExtentX        =   741
         _ExtentY        =   318
      End
      Begin vbwTBar.cToolbarHost tbhMain 
         Height          =   210
         Index           =   0
         Left            =   165
         TabIndex        =   15
         Top             =   105
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   370
         BorderStyle     =   0
      End
   End
   Begin vbwTBar.cToolbar tbrMain 
      Index           =   1
      Left            =   6015
      Top             =   420
      _ExtentX        =   741
      _ExtentY        =   318
   End
   Begin vbwTBar.cToolbarHost tbhMain 
      Height          =   330
      Index           =   1
      Left            =   6030
      TabIndex        =   16
      Top             =   90
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   582
      BorderStyle     =   0
   End
   Begin VB.PictureBox picOpen 
      BorderStyle     =   0  'None
      Height          =   3420
      Index           =   2
      Left            =   105
      ScaleHeight     =   3420
      ScaleWidth      =   6345
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   630
      Width           =   6345
      Begin vbAcceleratorGrid.vbalGrid lvwFiles 
         Height          =   2475
         Left            =   30
         TabIndex        =   3
         Top             =   150
         Width           =   6180
         _ExtentX        =   10901
         _ExtentY        =   4366
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
         DefaultRowHeight=   15
      End
   End
   Begin VB.Label lblHeader 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   1350
      TabIndex        =   14
      Top             =   135
      Width           =   1485
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   390
      Left            =   1290
      Top             =   60
      Width           =   4695
   End
End
Attribute VB_Name = "frmOpen"
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
Private Const WM_ACTIVATE = &H6

Private WithEvents cDlg As clsHookDlg
Attribute cDlg.VB_VarHelpID = -1
Private cFlatCombo      As clsFlatCombo
Private cFavourites     As clsFavourites
Private bLoaded         As Boolean
Private m_bAddToProject As Boolean
Private lCurrentTab     As Long
Private m_lInitTab      As Long
Private bDlgCancel      As Boolean

Public Property Let AddToProject(bNew As Boolean)
    'sets whether we want to add the selected file(s)
    'to the current project instead
    m_bAddToProject = bNew
End Property

'*** Control Events ***
Private Sub cboPath_Click()
Dim sInit As String
    With cDlg
        sInit = .ItemCaption(FILENAME_TEXTBOX)
        'goto the selected path...
        'fill the textbox with the path
        .ItemCaption(FILENAME_TEXTBOX) = cboPath.Text
        'simulate pressing open
        .SimulateOpen
        'we are now in the specified folder... clear textbox
        .ItemCaption(FILENAME_TEXTBOX) = sInit
    End With
End Sub

Private Sub cDlg_DialogOK(bCancel As Boolean)
    bDlgCancel = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyTab And Shift = vbCtrlMask Then
        'next item
        tbrMain(0).RaiseButtonClick IIf(lCurrentTab = 2, 0, lCurrentTab + 1)
    ElseIf KeyCode = vbKeyTab And Shift = (vbCtrlMask + vbShiftMask) Then
        'last item
        tbrMain(0).RaiseButtonClick IIf(lCurrentTab = 0, 2, lCurrentTab - 1)
    ElseIf KeyCode = vbKeyF1 Then
        cDialog.ShowHelpTopic 11, hWnd
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdOpen_Click
    ElseIf KeyAscii = vbKeyEscape Then
        cmdCancel_Click
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        'don't want to close dialog this way!
        Cancel = -1
        cDlg.Simulatecancel
    End If
End Sub

Private Sub lstTemplate_Click()
Dim vInfo As TemplateInfo
On Error GoTo ErrHandler
    If lstTemplate.Text = "" Then Exit Sub
    'load the template into the preview textbox
    txtPreview = LoadTextFile(App.Path & "\_templates\" & lstTemplate.Text & ".txt")
    'get the info for that template
    vInfo = GetTemplateInfo(lstTemplate.Text & ".txt")
    'get its description
    lblDescription.Caption = vInfo.sDescription
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Open.Template_Click", "Filename: " & App.Path & "\_templates\" & lstTemplate.Text & ".txt"
End Sub

Private Sub cmdCancel_Click()
    'set cancel flag
    bDlgCancel = True
    'cancel dialog
    cDlg.Simulatecancel
End Sub

Private Sub cmdOpen_Click()
    'set cancel flag
    bDlgCancel = False
    'perform correct operation, depending on selected tab
    Select Case tbrMain(0).ButtonKey(lCurrentTab)
    Case "Open" '"Open"
        cDlg.SimulateOpen
    Case "Recent" '"Recent"
        cDlg.ItemCaption(FILENAME_TEXTBOX) = lvwFiles.CellText(lvwFiles.SelectedRow, 2)
        cDlg.SimulateOpen
    Case "New" '"New"
        'we want to cancel the common open dialog, because we have
        'selected a template from our dialog
        cDlg.Simulatecancel
    End Select
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
    'Init the favourites list
    Set cFavourites = New clsFavourites
    'Load the resource strings...
    LoadResStrings Controls
    'make the combo flat
    Set cFlatCombo = New clsFlatCombo
    cFlatCombo.Attach cboPath.hWnd
    cFlatCombo.BackColor = &H808080
    'build the toolbars
    With tbrMain(0)
        'build the toolbar
        'set image list
        .ImageSource = CTBExternalImageList
        .SetImageList frmMainForm.vbalMain, CTBImageListNormal
        'create tb
        .CreateToolbar 16
        'add buttons
        .AddButton "New", IndexForKey("NEWTEXT"), , , Empty, CTBCheckGroup, "New"
        .AddButton "Open", IndexForKey("OPEN"), , , Empty, CTBCheckGroup, "Open"
        .AddButton "Recent", IndexForKey("HISTORY"), , , Empty, CTBCheckGroup, "Recent"
    End With
    
    With tbhMain(0)
        'init toolbar host
        'no border
        .BorderStyle = etbhBorderStyleNone
        'set width
        .Width = tbrMain(0).ToolbarWidth * Screen.TwipsPerPixelX + 60
        .Left = 15
        .Top = 30
        'set height
        .Height = tbrMain(0).ToolbarHeight * Screen.TwipsPerPixelY
        ctlFrame1.Height = .Height + 60
        ctlFrame1.Width = .Width + 30
        'capture main toolbar
        .Capture tbrMain(0)
    End With
    
    With tbrMain(1)
        'build the toolbar
        'set image list
        .ImageSource = CTBExternalImageList
        .SetImageList frmMainForm.vbalMain, CTBImageListNormal
        'create tb
        .CreateToolbar 16
        'add buttons
        .AddButton "Add To Favourites", IndexForKey("NEWFOLDER"), , , Empty, CTBNormal, "New"
    End With
    With tbhMain(1)
        'init toolbar host
        'set width
        .Width = tbrMain(1).ToolbarWidth * Screen.TwipsPerPixelX + 60
        'set height
        .Height = tbrMain(1).ToolbarHeight * Screen.TwipsPerPixelY
        'capture main toolbar
        .Capture tbrMain(1)
    End With
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Open.Load"
End Sub



Private Sub lvwFiles_ColumnClick(ByVal lCol As Long)
    'sort the Recent files list by the clicked column
    SortDocsList lCol
End Sub

Private Sub lvwFiles_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    cmdOpen_Click
End Sub




Private Sub tbrMain_ButtonClick(Index As Integer, ByVal lButton As Long)
    If Index = 0 Then
        Dim i As Integer
        'show and enable the selected tab's controls
        'and hide and disable all others
        For i = 0 To tbrMain(0).ButtonCount - 1
            If i = lButton Then
                picOpen(i).Left = 135
            Else
                picOpen(i).Left = -20000
            End If
        Next
        'set the label
        lblHeader.Caption = tbrMain(0).ButtonKey(lButton)
        'ensure the buttonis checked
        tbrMain(0).ButtonChecked(lButton) = True
        'save button
        lCurrentTab = lButton
        'show favourites part if we have the open tab
        picFavLabel.Visible = (lButton = 1)
        cboPath.Visible = (lButton = 1)
        tbhMain(1).Visible = (lButton = 1)
        'set open button to show open/new
        cmdOpen.Caption = IIf(lButton = 0, LoadResString(1245), LoadResString(1244)) '1245=&New;1244=&Open
    Else
        'new fav button clicked
        'add the current folder to the favourites list
        cFavourites.AddFavourite cDlg.CurrentFolder, cboPath
    End If
End Sub

'*** Common Dialog Code ***
Public Sub Init(bNew As Boolean, Optional lTab As Long = -1)
    On Error GoTo ErrHandler
    If Visible Then Exit Sub  'I managed to get this happen once, so its not going to happen again!
    
    Dim iFileCount  As Long   'number of files returned by dialog
    Dim sFiles()    As String 'array of selected files
    Dim sDir        As String 'the path containing the selected files
    Dim i           As Long   'counter

    tbrMain(0).ButtonEnabled("New") = Not m_bAddToProject
    'default to open tab
    If lTab = -1 Then lTab = 1 '1
    m_lInitTab = lTab
    'don't have a path selected
    cboPath.ListIndex = -1
    LoadRecentDocs
    'reset flag
    bDlgCancel = True
    If cDlg Is Nothing Then
        'we haven't initialized class yet
        Set cDlg = New clsHookDlg
        'make everything appear in the write ZOrder
        For i = 0 To 2
            'ensure picbox in correct pos
            picOpen(i).Move 135, 500
            picOpen(i).ZOrder
        Next i
        picOpen(lTab).ZOrder
        cmdOpen.ZOrder
        cmdCancel.ZOrder
        lstTemplate.Clear
        'add all the templates to the template list
        AddAllFilesInDir "_templates", lstTemplate
        'select the first item, if there is a template in the list
        If lstTemplate.ListCount <> 0 Then lstTemplate.ListIndex = 0
        'set the classes properties
        With cDlg
            'we want the open dlg to be contained by picOpen(1)
            .ContainerhWnd = picOpen(1).hWnd
            'it is a custom dialog
            .CustomDialog = True
            'this is the parent form
            .ParentFormhWnd = hWnd
            'set owneer form
            .hWnd = frmMainForm.hWnd
        End With
    End If
    'for some reason, the open tab has to be active when we show
    'the open dialog... we can then switch to the correct tab
    'in the cDlg_Show event
    tbrMain(0).RaiseButtonClick 1
    'reload
    cFavourites.LoadFavourites cboPath, Me
    With cDlg
        'we want errors!
        .CancelError = True
        .FileName = ""
        'set the flags
        .Flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST Or OFN_ALLOWMULTISELECT Or OFN_EXPLORER Or OFN_HIDEREADONLY
        'set the initial dir
        .InitDir = GetSetting(REG_KEY, "Settings", "DefaultFilePath", CurDir$)
        'use default filter
        .Filter = sFileFilter
        'set the filter index from last selected item
        .FilterIndex = GetSetting(REG_KEY, "Settings", "DefaultFileType", "1")  '1
        'display the dialog
        .ShowOpen
        'return multifilename info
        .ParseMultiFileName sDir, sFiles(), iFileCount
        'get the filename and filetitle
        .FileName = StripTerminator(.FileName)
        .FileTitle = StripTerminator(.FileTitle)
        'save the favourites, filterindex and last folder
        cFavourites.SaveFavourites cboPath
        SaveSetting REG_KEY, "Settings", "DefaultFileType", .FilterIndex
        SaveSetting REG_KEY, "Settings", "DefaultFilePath", GetFolder(.FileName)

        If (iFileCount = 1) Then 'only one file
            If m_bAddToProject Then
                'add to project
                frmProject.NewItem .FileName, .FileTitle, vbwFile
            Else
                Select Case .FilterIndex
                Case lProjectFilter 'Project filter index
                    'display the project window
                    frmMainForm.ShowProjectWindow
                    'open the file as a project
                    frmProject.OpenProject .FileName
                Case lProjectFilter + 1 'Workspace filter index
                    cWorkspace.Load .FileName
                Case 1 'All Supported Files index
                    LoadFileDefault .FileName, False
                Case Else
                    'load the file
                    '.FilterIndex = (lProjectFilter + 2) if we want to force text
                    cDocuments.LoadFile .FileName, .FilterIndex = (lProjectFilter + 2)
                End Select
            End If
        Else
            If .FilterIndex = lProjectFilter Then
                'can't allow multiple projects at once
                cDialog.ShowWarning "You can only open one Developers Pad Project at a time", "Open.Init"
            Else
                'we are closing/opening multiple... we don't want
                'LoadFile to change the status bar
                bClosingMultiple = True
                'do it ourselves once..
                SetStatusBar "Opening files...."
                For i = 1 To iFileCount
                    If m_bAddToProject Then
                        'add the file to the current project
                        frmProject.NewItem sDir & "\" & sFiles(i), , vbwFile
                    Else
                        ' open as file
                        cDocuments.LoadFile sDir & "\" & sFiles(i), .FilterIndex = (lProjectFilter + 1)
                    End If
                Next i
                'reset flag
                bClosingMultiple = False
                'update the list of windows...
                frmMainForm.UpdateWindowList
                'reset statusbar
                SetStatusBar
            End If
        End If
    End With
    'save the favourites!
    cFavourites.SaveFavourites cboPath
    SendMessage cDlg.hWnd, WM_ACTIVATE, 1, 0&
    Exit Sub
ErrHandler:
    If (Err.Number <> 20001) Then
        cDialog.ErrHandler Err, Err.Description, "Core.OpenFile"
    Else
        'dialog cancelled
        If tbrMain(0).ButtonKey(lCurrentTab) = "New" And bDlgCancel = False Then  'new
            'new tab is selected, and the cancel button has not been pressed...
            'load the document with the selected template
            'frmMainForm.LoadNewDoc , lstTemplate.Text & ".txt"
            cDocuments.New , lstTemplate.Text & ".txt"
            'we still want to save the favourites
            cFavourites.SaveFavourites cboPath
        End If
        SendMessage cDlg.hWnd, WM_ACTIVATE, 1, 0&
    End If
End Sub

Private Sub cDlg_DialogClose()
    'hooked cmdlg has been closed
    'hide this dialog too
    Hide
End Sub

Private Sub cDlg_InitDialog(ByVal hDlg As Long)
    With cDlg
        picOpen(1).Width = .DialogWidth * Screen.TwipsPerPixelX '+ 500 '1100
        picOpen(1).Height = (.DialogHeight * Screen.TwipsPerPixelX)
        Width = picOpen(1).Width + 200 '180
        Height = picOpen(1).Height + 500 '225

        'set the dialogs font
        cDlg.SetFont "Tahoma", hdc
        'if we have loaded the dialog before,
        'we don't need to do anything else
        If bLoaded Then Exit Sub
        'set flag
        bLoaded = True
        'position our command buttons relative to
        'the cmdlg's buttons

        .SetItemPos IDOK, cmdOpen.hWnd, 0
        .SetItemPos IDCANCEL, cmdCancel.hWnd, 0

    End With
End Sub
Private Sub cDlg_Show()
    DoEvents
    'obey the classes commands!
    'display this dialog
    Show , frmMainForm
    'switch to correct tab
    tbrMain(0).RaiseButtonClick m_lInitTab
End Sub

'*** Recent Documents ***
Private Sub LoadRecentDocs()
    Dim cItems  As Collection
    Dim lRow    As Long
    Dim i       As Long
    'lists the recent documents in the list control...
    With lvwFiles
        'clear the list
        .Clear True
        'add the columns
        .AddColumn "File", "File", , , 64
        .AddColumn "Folder", "Folder", , , 217
        .AddColumn "Accessed", "Accessed", , , 110
        'get the history collection from the main form
        Set cItems = frmMainForm.FileHistoryItems
        If cItems.Count <> 0 Then
            'if there are items load them
            For i = 1 To cItems.Count
                'add a new row
                .AddRow
                'get the rows new index
                lRow = .Rows
                'set its filetitle
                .CellText(lRow, 1) = GetCaption(cItems(i))
                'its full path
                .CellText(lRow, 2) = cItems(i)
                If Left$(cItems(i), 2) <> "\\" Then
                    'don't bother if on a network
                    On Error Resume Next
                    'and its modified date
                    .CellText(lRow, 3) = FileDateTime(cItems(i))
                    If Err Then .RemoveRow lRow
                End If
                
            Next
            'select the first item
            .SelectedRow = 1
        End If
    End With
End Sub

Private Sub SortDocsList(lCol As Long)
    'sorts the recent document list
    Static lLastCol As Long
    With lvwFiles.SortObject
        'clear sort object
        .Clear
        'set the sort column
        .SortColumn(1) = lCol
        'set the order... if this column was clicked last time,
        'then reverse the order
        .SortOrder(1) = IIf(lLastCol <> lCol, 1, 2) 'CCLOrderAscending
        If lCol = 3 Then 'date column
            .SortType(1) = CCLSortDate
        Else
            .SortType(1) = CCLSortString
        End If
        'do the sort
        lvwFiles.Sort
        'save the last column
        If lLastCol = lCol Then
            'reset flag so we will go back to Ascending
            'order next time
            lLastCol = 0
        Else
            lLastCol = lCol
        End If
    End With
End Sub
