VERSION 5.00
Object = "{5C0E11AE-2C8C-4C35-BC7A-D9B469D5DE4D}#6.1#0"; "VBWTRE~2.OCX"
Begin VB.Form frmAddIns 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Developers Pad Add-Ins"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddIns.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin vbwTreeView.TreeView tvwAddIns 
      Height          =   2340
      Left            =   105
      TabIndex        =   13
      Top             =   555
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   4128
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
      ItemHeight      =   0
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "1293"
      Height          =   345
      Left            =   3915
      TabIndex        =   12
      Top             =   3150
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "1047"
      Height          =   1065
      Left            =   3015
      TabIndex        =   6
      Top             =   1830
      Width           =   2880
      Begin VB.CheckBox chkShowOnTb 
         Caption         =   "1292"
         Height          =   252
         Left            =   90
         TabIndex        =   9
         Top             =   735
         Width           =   2130
      End
      Begin VB.CheckBox chkShowOnMenu 
         Caption         =   "1291"
         Height          =   252
         Left            =   90
         TabIndex        =   8
         Top             =   495
         Width           =   2070
      End
      Begin VB.CheckBox chkLoadStartup 
         Caption         =   "1290"
         Height          =   252
         Left            =   90
         TabIndex        =   7
         Top             =   240
         Width           =   1410
      End
   End
   Begin DevPad.ctlFrame ctlFrame1 
      Height          =   390
      Left            =   90
      TabIndex        =   4
      Top             =   75
      Width           =   2820
      _ExtentX        =   4974
      _ExtentY        =   688
      Begin VB.Label lblHeader1 
         BackStyle       =   0  'Transparent
         Caption         =   "1289"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   105
         TabIndex        =   5
         Top             =   90
         Width           =   1620
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "1001"
      Default         =   -1  'True
      Height          =   345
      Left            =   5235
      TabIndex        =   0
      Top             =   3150
      Width           =   1215
   End
   Begin VB.Label lblAuthor 
      Height          =   255
      Left            =   3825
      MouseIcon       =   "frmAddIns.frx":000C
      TabIndex        =   11
      Top             =   1485
      Width           =   2595
   End
   Begin VB.Label Label1 
      Caption         =   "1170"
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
      Left            =   3030
      TabIndex        =   10
      Top             =   1485
      Width           =   840
   End
   Begin VB.Label lblDescription 
      Caption         =   "Description...."
      Height          =   780
      Left            =   3045
      TabIndex        =   3
      Top             =   540
      Width           =   3435
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   4
      X1              =   120
      X2              =   6465
      Y1              =   3030
      Y2              =   3030
   End
   Begin VB.Label lblHeader2 
      BackStyle       =   0  'Transparent
      Caption         =   "Class Spy"
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
      Left            =   3105
      TabIndex        =   2
      Top             =   165
      Width           =   3030
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      Caption         =   "1240"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   120
      MouseIcon       =   "frmAddIns.frx":0316
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3210
      Width           =   360
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3000
      Top             =   90
      Width           =   3495
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   5
      X1              =   105
      X2              =   6465
      Y1              =   3045
      Y2              =   3030
   End
End
Attribute VB_Name = "frmAddIns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bIgnore As Boolean
Private lImgFolder As Long
Private lImgFolderOpen As Long
Private lImgAddIn As Long
Private cFlatOpt(2) As New clsFlatOpt

Private Sub chkShowOnMenu_Click()
    If bIgnore Then Exit Sub
    With cAddIns.Info(AddInIndex)
        If chkShowOnMenu.Value = 0 Then
            If frmMainForm.ctlPopMenu.MenuExists("ToolsAddIn" & AddInIndex) Then frmMainForm.ctlPopMenu.RemoveItem ("ToolsAddIn" & AddInIndex)
        Else
            frmMainForm.ctlPopMenu.InsertItem .Name, "ToolsSep2", "ToolsAddIn" & AddInIndex
        End If
    End With
    UpdateInfo
End Sub
Private Sub chkLoadStartup_Click()
    UpdateInfo
End Sub
Private Sub chkShowOnTB_Click()
    UpdateInfo
End Sub
Private Sub UpdateInfo()
    Dim cInfo As ADDININFO
    If bIgnore Then Exit Sub
    cInfo = cAddIns.Info(AddInIndex)
    With cInfo
        .ShowInTB = CBool(chkShowOnTb)
        .ShowInMenu = CBool(chkShowOnMenu)
        .LoadAtStartup = CBool(chkLoadStartup)
    End With
    cAddIns.UpdateInfo AddInIndex, cInfo
End Sub

Private Sub cmdInstall_Click()
    ShellFunc App.Path & "\addins.ini"
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()

On Error GoTo ErrHandler
    Dim i As Long

    With tvwAddIns
        .LabelEditing = True
        .Lines = True
        .PlusMinus = True
        .Rootlines = False
        .ShowSelected = True
        .NoDragDrop = True
    End With
    
    LoadResStrings Controls
    With frmMainForm.vbalMain
        tvwAddIns.hImageList = .hIml
        lImgFolder = .ItemIndex("FOLDERCLOSED")
        lImgFolderOpen = .ItemIndex("FOLDEROPEN")
        lImgAddIn = .ItemIndex("FILE_BAT")
    End With
    cFlatOpt(0).Attach chkLoadStartup.hWnd
    cFlatOpt(1).Attach chkShowOnMenu.hWnd
    cFlatOpt(2).Attach chkShowOnTb.hWnd
    
    For i = 1 To cAddIns.InfoCount
        AddTreeViewItem cAddIns.Info(i).Category, cAddIns.Info(i).Name, i
    Next i

    With tvwAddIns
        If .Count > 0 Then
            If .ItemHasChildren("Misc") Then
                .Selected = .ItemChild("Misc")
                .EnsureVisible .ItemChild("Misc")
            Else
                .Selected = .ItemIndex("Misc")
                .EnsureVisible "Misc"
            End If
        Else
            tvwAddIns.Add 0, AlphabeticalChild, "Misc", "Misc", lImgFolder
        End If
        tvwAddIns_SelChanged
    End With
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, "An error occured whilst loading the Add-Ins: " & Error
End Sub

Private Sub AddTreeViewItem(ByVal sCategory As String, ByVal sName As String, ByVal sKey As String)
Dim hParent As Long
    If sCategory = "" Then sCategory = "Misc"
    If tvwAddIns.IsValidNewKey(sCategory) = True Then
        hParent = tvwAddIns.Add(0&, AlphabeticalChild, sCategory, sCategory, lImgFolder)
    Else
        hParent = tvwAddIns.ItemIndex(sCategory)
    End If
    tvwAddIns.Add hParent, AlphabeticalChild, "F" & sKey, sName, lImgAddIn
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Erase cFlatOpt()
End Sub

Private Sub lblAuthor_Click()
    If lblAuthor.Tag <> "" Then ShellFunc lblAuthor.Tag
End Sub

Private Sub lblHelp_Click()
    cDialog.ShowHelpTopic 6, hWnd
End Sub

Private Property Get AddInIndex() As Long
Dim sKey As String
    sKey = tvwAddIns.ItemKey(tvwAddIns.Selected)
    AddInIndex = CLng(Right$(sKey, Len(sKey) - 1))
End Property
'
Private Sub tvwAddIns_ItemDblClick(hItem As Long)
    If tvwAddIns.ItemImage(tvwAddIns.Selected) = lImgAddIn Then
        If cAddIns.Add(AddInIndex) Then cAddIns.Tool(AddInIndex).ShowDialog
    End If
End Sub

Private Sub tvwAddIns_ItemExpand(hItem As Long, ExpandType As vbwTreeView.ExpandTypeConstants)
    If ExpandType = Collapse Then
        tvwAddIns.ItemImage(hItem) = lImgFolder
    Else
        tvwAddIns.ItemImage(hItem) = lImgFolderOpen
    End If
End Sub

Private Sub tvwAddIns_SelChanged()
    Dim bValid As Boolean
    bValid = (tvwAddIns.ItemImage(tvwAddIns.Selected) = lImgAddIn)
    Frame1.Enabled = bValid
    lblAuthor.Enabled = bValid
    If bValid = False Then
        lblDescription.Caption = ""
        lblAuthor.Caption = ""
        lblHeader2.Caption = tvwAddIns.ItemText(tvwAddIns.Selected)
    Else
        With cAddIns.Info(AddInIndex)
            lblDescription.Caption = .Description
            lblHeader2.Caption = .Name
            lblAuthor.Caption = .Author
            lblAuthor.Tag = .Website
            lblAuthor.MousePointer = IIf(.Website = "", vbDefault, vbCustom)
            lblAuthor.ForeColor = IIf(.Website = "", &H80000012, &HFF0000)

            bIgnore = True
            chkShowOnTb.Value = Abs(.ShowInTB)
            chkLoadStartup.Value = Abs(.LoadAtStartup)
            chkShowOnMenu.Value = Abs(.ShowInMenu)
            'work around for flatopt class bug...
            'a BM_SETCHECK message doesn't seem to be sent
            'when changing the value property
            chkLoadStartup.Refresh
            chkShowOnTb.Refresh
            chkShowOnMenu.Refresh
            bIgnore = False
        End With
    End If
End Sub

