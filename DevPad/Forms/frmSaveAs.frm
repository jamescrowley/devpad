VERSION 5.00
Object = "{1F5999A2-1D0B-11D4-82CF-004005AAE138}#6.1#0"; "VBWTBA~1.OCX"
Begin VB.Form frmSaveAs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save As..."
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSaveAs.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboPath 
      Height          =   315
      ItemData        =   "frmSaveAs.frx":000C
      Left            =   3255
      List            =   "frmSaveAs.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   105
      Width           =   2565
   End
   Begin VB.PictureBox picSave 
      BorderStyle     =   0  'None
      Height          =   3555
      Left            =   135
      ScaleHeight     =   3555
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   540
      Width           =   6195
   End
   Begin DevPad.ctlFrame ctlFrame1 
      Height          =   390
      Left            =   75
      TabIndex        =   1
      Top             =   60
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   688
      Begin VB.Image imgSave 
         Height          =   240
         Left            =   90
         Picture         =   "frmSaveAs.frx":0010
         Top             =   75
         Width           =   240
      End
      Begin VB.Image imgOpen 
         Height          =   240
         Left            =   105
         Picture         =   "frmSaveAs.frx":015A
         Top             =   60
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin vbwTBar.cToolbar tbrMain 
      Left            =   5895
      Top             =   435
      _ExtentX        =   741
      _ExtentY        =   318
   End
   Begin vbwTBar.cToolbarHost tbhMain 
      Height          =   345
      Left            =   5910
      TabIndex        =   3
      Top             =   105
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   609
      BorderStyle     =   0
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   390
      Left            =   570
      Top             =   75
      Width           =   5295
   End
End
Attribute VB_Name = "frmSaveAs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Private Const WM_ACTIVATE = &H6

Private WithEvents cDlg As clsHookDlg
Attribute cDlg.VB_VarHelpID = -1
Private cFlatCombo      As clsFlatCombo
Private bLoaded         As Boolean
Private cFavourites     As clsFavourites
Private m_frmOwner      As Form
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
'*** Dialog Hook ***
Private Sub cDlg_DialogClose()
    On Error Resume Next
    'do what we are told... hide the dialog
    Hide
End Sub
Private Sub cDlg_DialogOK(bCancel As Boolean)
    If GetExtension(cDlg.SelectedFile) = "" Then
        'until I can decide what to do for a default extension
        'for each file group... ask them to specify one
        
        '1260="You have not entered an file extension. Save without any extension?"
        If cDialog.ShowYesNo(LoadResString(1260), False) = No Then bCancel = True
    End If
End Sub
Private Sub cDlg_InitDialog(ByVal hDlg As Long)
    'set the dialog font
    cDlg.SetFont "Tahoma", hdc
    picSave.Width = cDlg.DialogWidth * Screen.TwipsPerPixelX '+ 500 '1100
    picSave.Height = cDlg.DialogHeight * Screen.TwipsPerPixelX
    Width = picSave.Width + 200 '180
    Height = picSave.Height + 500 '225
End Sub
Public Function Init(Optional bDefault As Boolean = True, Optional frmForm As IDevPadDocument, Optional sFilter As String = "", Optional sTitle As String = "", Optional sFileName As String = "", Optional bSave As Boolean = True, Optional Owner As Form = Nothing, Optional nFilter As Long = -1) As Boolean
    Set m_frmOwner = Owner
    If m_frmOwner Is Nothing Then Set m_frmOwner = frmMainForm

    On Error GoTo ErrHandler
    'set default values...
    If sFilter = "" Then sFilter = sFileFilter
    If sTitle <> "" Then
        Caption = sTitle
    ElseIf bSave Then
        Caption = "Save..."
    Else
        Caption = "Open..."
    End If
    
    If cDlg Is Nothing Then
        'initialize the hookdlg class
        Set cDlg = New clsHookDlg
        'ensure buttons stay on top
        With cDlg
            'container is picSave
            .ContainerhWnd = picSave.hWnd
            'we are hosting a custom dialog
            .CustomDialog = True
            'this is the parent form
            .ParentFormhWnd = hWnd
        End With
    End If
    'owner is frmMainForm
    cDlg.hWnd = m_frmOwner.hWnd
    'load the favourite folders...
    cFavourites.LoadFavourites cboPath, Me
    With cDlg
        'we want errors
        .CancelError = True
        'set flags
        .Flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST Or OFN_EXPLORER Or OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT

        'set the initial dir
        .InitDir = GetSetting(REG_KEY, "Settings", "DefaultFilePath", CurDir$)
        .Filter = sFilter
        'set filter index
        If nFilter = -1 Then
            .FilterIndex = GetSetting(REG_KEY, "Settings", "DefaultFileType", "1") '1
        Else
            .FilterIndex = nFilter
        End If
        'set the filename...
        .FileName = sFileName
    End With
    If bDefault Then
        'perform default action... save
        With frmForm
            'set the default extension to the form's current one
            If frmForm.FileName <> "" Then cDlg.DefaultExt = GetExtension(frmForm.FileName)
            'display the correct icon
            imgSave.Visible = True
            imgOpen.Visible = False
            ' show dialog
            cDlg.ShowSave
            'retreive the filename
            sFileName = cDlg.FileName
            ' Save the file
            .SaveDocument sFileName
            'add to file menu
            frmMainForm.UpdateFileMenu sFileName, 1
            frmMainForm.UpdateWindowList
            'set the forms caption
            .DocumentCaption = cDlg.FileTitle
            'succeeded
            Init = True
        End With
    Else
        With cDlg
            'otherwise
            'display open or save..?
            'display the correct icon
            imgSave.Visible = (bSave)
            imgOpen.Visible = Not (bSave)
            If bSave Then
                .ShowSave
            Else
                .ShowOpen
            End If
            'save the result to the CmDlg struct
            CmDlg.FileName = .FileName
            CmDlg.FileTitle = .FileTitle
            CmDlg.FilterIndex = .FilterIndex
            'succeeded
            Init = True
        End With
    End If
    'save the favourites
    cFavourites.SaveFavourites cboPath
    'save filepath
    SaveSetting REG_KEY, "Settings", "DefaultFilePath", GetFolder(cDlg.FileName)
    SendMessage cDlg.hWnd, WM_ACTIVATE, 1, 0&
    Exit Function
ErrHandler:
    If (Err.Number <> 20001) Then
        cDialog.ErrHandler Err, Err.Description, "Core.OpenFile"
    Else
        'dialog cancelled
        Init = False
        SendMessage cDlg.hWnd, WM_ACTIVATE, 1, 0& ' frmMainForm.SetFocus
    End If
End Function

Private Sub cDlg_Show()
    On Error Resume Next
    'cmdlg displayed... show this dialog
    Show , m_frmOwner
    If Err = 373 Then Show 'This interaction between compiled and design environment components is not supported.
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then cDialog.ShowHelpTopic 14, hWnd
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cDlg.SimulateOpen
    ElseIf KeyAscii = vbKeyEscape Then
        cDlg.Simulatecancel
    End If
End Sub

Private Sub Form_Load()
    'initialize favourites
    Set cFavourites = New clsFavourites
    'load resource strings
    LoadResStrings Controls
    'make the combo flat
    Set cFlatCombo = New clsFlatCombo
    cFlatCombo.Attach cboPath.hWnd
    cFlatCombo.BackColor = &H808080
    'build the toolbars
    With tbrMain
        'build the toolbar
        'set image list
        .ImageSource = CTBExternalImageList
        .SetImageList frmMainForm.vbalMain, CTBImageListNormal
        'create tb
        .CreateToolbar 16
        'add buttons
        .AddButton "Add To Favourites", IndexForKey("NEWFOLDER"), , , Empty, CTBNormal, "New"
    End With
    With tbhMain
        'init toolbar host
        'set width
        .Width = tbrMain.ToolbarWidth * Screen.TwipsPerPixelX + 60
        'set height
        .Height = tbrMain.ToolbarHeight * Screen.TwipsPerPixelY
        'capture main toolbar
        .Capture tbrMain
    End With
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Open.Load"
End Sub

Private Sub tbrMain_ButtonClick(ByVal lButton As Long)
    'new fav button clicked
    'add the current folder to the favourites list
    cFavourites.AddFavourite cDlg.CurrentFolder, cboPath
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        'don't want to close dialog this way!
        Cancel = -1
        cDlg.Simulatecancel
    End If
End Sub

