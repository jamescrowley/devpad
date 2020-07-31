VERSION 5.00
Object = "{1F5999A2-1D0B-11D4-82CF-004005AAE138}#6.0#0"; "vbwTbar_6.ocx"
Object = "{017E002E-D7CC-11D2-8E21-44B10AC10000}#22.0#0"; "vbwGrid.ocx"
Begin VB.Form frmFindResults 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find Results"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFindResults.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNew 
      Cancel          =   -1  'True
      Caption         =   "1225"
      Height          =   345
      Left            =   4095
      TabIndex        =   9
      Top             =   2175
      Width           =   1200
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "1013"
      Height          =   345
      Left            =   4095
      TabIndex        =   8
      Top             =   75
      Width           =   1200
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "1018"
      Height          =   345
      Left            =   4095
      TabIndex        =   7
      Top             =   495
      Width           =   1200
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "1017"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4095
      TabIndex        =   6
      Top             =   915
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox picResults 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   2520
      Index           =   0
      Left            =   30
      ScaleHeight     =   2520
      ScaleWidth      =   3990
      TabIndex        =   0
      Top             =   75
      Width           =   3990
      Begin VB.TextBox txtReplace 
         Height          =   285
         Index           =   0
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   360
         Width           =   2745
      End
      Begin VB.TextBox txtFind 
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   2745
      End
      Begin vbAcceleratorGrid.vbalGrid lvwResults 
         Height          =   1695
         Index           =   0
         Left            =   0
         TabIndex        =   5
         Top             =   750
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   2990
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
      Begin VB.Label lblReplace 
         Caption         =   "1103"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblFind 
         Caption         =   "1202"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox picHolder 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   5355
      TabIndex        =   10
      Top             =   2610
      Width           =   5355
      Begin vbwTBar.cToolbar tbrMain 
         Left            =   1035
         Top             =   45
         _ExtentX        =   741
         _ExtentY        =   318
      End
      Begin vbwTBar.cToolbarHost tbhMain 
         Height          =   210
         Left            =   210
         TabIndex        =   11
         Top             =   75
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   370
         BorderStyle     =   0
      End
   End
End
Attribute VB_Name = "frmFindResults"
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
Private iTabCount   As Integer 'number of loaded tabs...
Private bReplaced() As Boolean 'array containing whether the result set was a replace result set
Private iCurrentTab As Integer 'active selected tab

'Note: This form is waiting to be made dockable...
'      that is, once I have worked out how to let two windows
'      dock on the same side at once!

'returns the number of results currently loaded
Public Property Get ResultsCount() As Integer
    ResultsCount = iTabCount
End Property
'returns the current tab
Public Property Get CurrentTab() As Integer
    CurrentTab = iCurrentTab
End Property
'Clears the current result set
Private Sub cmdClear_Click()
    'clear the listview
    lvwResults(iCurrentTab).Clear
    'reset tag
    lvwResults(iCurrentTab).Tag = ""
    'reset find text etc
    ResetPane iCurrentTab
    SetCaption
End Sub
Private Sub ResetPane(lTab As Integer)
    'clear find/replace textboxes
    txtFind(lTab).Text = ""
    txtReplace(lTab).Text = ""
    txtReplace(lTab).Tag = "" 'reset prompt tag
    txtReplace(lTab).Locked = False
End Sub
Private Sub cmdClose_Click()
    'hide the form... don't unload!
    Hide
End Sub

Private Sub cmdNew_Click()
    'create a new result set
    NewResultPane True
End Sub

Private Sub cmdReplace_Click()
    Dim sReplace       As String
    Dim sFind          As String
    Dim lChange        As Long
    Dim lReplaceLen    As Long
    Dim i              As Long
    Dim bItemReplaced  As Boolean
    Dim sTag           As String
    Dim sText          As String
    Dim lPos           As Long
    'display an error if no replace text has been set, just in case...
    '(but only once!)
    If txtReplace(iCurrentTab).Tag <> "1" And txtReplace(iCurrentTab) = "" Then
        Select Case cDialog.ShowYesNo(LoadResString(1104), True) ' no text set. replace with nothing?
        Case Yes
            'we have shown the error once, don't bother the user again!
            txtReplace(iCurrentTab).Tag = "1"
        Case No, Cancelled
            'abort...
            Exit Sub
        End Select
    End If
    
    'we are wanting to replace the text
    If cmdReplace.Caption = LoadResString(1013) Then ' "&Replace"
        'set the find/replace text...
        sFind = txtFind(iCurrentTab).Text
        sReplace = txtReplace(iCurrentTab).Text
        'we are going to have replaced this item...
        bItemReplaced = True
        'now that we have done one replace, we don't want
        'to let the user change the replace text
        txtReplace(iCurrentTab).Locked = True
    Else
        'we have already replaced... so we are actually
        'looking for the replace text.
        sFind = txtReplace(iCurrentTab).Text
        sReplace = txtFind(iCurrentTab).Text
        'we are going to have not replaced this item!
        bItemReplaced = False
    End If
    With ActiveDoc
        'check that the selection matches the find text...
        If LCase$(.SelText) = LCase$(sFind) Then
            'make sure the results window doesn't lose the focus
            '.NoEnterFocus = True
            'measure the change in length...
            'this is needed to update the pos other items in the result set
            lChange = Len(sReplace) - Len(sFind)
            'clear sel text if it contains a character return...
            'otherwise the richedit doesn't always delete it
            If InStr(1, .SelText, vbCrLf) Then .SelText = ""
            lPos = .SelStart + Len(sReplace)
            'then replace with new text
            .SelText = sReplace
            If lChange <> 0 Then
                'we need to update list of positions
                With lvwResults(iCurrentTab)
                    'don't waste time redrawing
                    .Redraw = False
                    sTag = .CellTag(.SelectedRow, 1)
                    For i = 1 To .Rows
                        sText = .CellText(i, 2)
                        If sText > lPos And .CellTag(i, 1) = sTag Then
                            'adjust value accordingly
                            .CellText(i, 2) = .CellText(i, 2) + lChange
                        End If
                    Next
                    'we can redraw again
                    .Redraw = True
                End With
            End If
           ' .NoEnterFocus = False
        Else
            'invalid selection.. display error
            cDialog.ShowWarning LoadResString(1105), "FindResults.Replace" 'invalid selection
            Exit Sub
        End If
    End With
    With lvwResults(iCurrentTab)
        ' now replaced... update row tag to say this
        .CellTag(.SelectedRow, 2) = IIf(bItemReplaced, "R", "")
        ' change selection to update buttons etc...
        lvwResults_SelectionChange iCurrentTab, .SelectedRow, -1
    End With
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then cDialog.ShowHelpTopic 8, hWnd
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        'don't unload the form if the user has tried to close it... just hide it!
        Hide
        Cancel = -1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
    'loop through the tabs...
    For i = 1 To iTabCount - 1
        txtFind(i).Visible = False
        txtReplace(i).Visible = False
        'we need to do this to stop VB crashing!
        'don't ask me why...
        'something to do with thin3d border, and a control array...
        SetParent txtFind(i).hWnd, 0
        SetParent txtReplace(i).hWnd, 0
        SetWindowPos txtFind(i).hWnd, 0, 0, 0, 0, 0, 0
        SetWindowPos txtReplace(i).hWnd, 0, 0, 0, 0, 0, 0
    Next
End Sub

Private Sub lvwResults_ColumnClick(Index As Integer, ByVal lCol As Long)
    'sort the results list according to column
    SortDocsList lCol
End Sub
Private Sub SortDocsList(lCol As Long)
    Static lLastCol As Long
    With lvwResults(iCurrentTab).SortObject
        'clear sort object
        .Clear
        'set its column
        .SortColumn(1) = lCol
        'and the order... reverse order if we are sorting
        'same column as last time
        .SortOrder(1) = IIf(lLastCol <> lCol, 1, 2) 'CCLOrderAscending
        If lCol = 1 Then
            'if we are sorting by first column, sort by string
            .SortType(1) = CCLSortString
        Else
            'otherwise, sort numerically
            .SortType(1) = CCLSortNumeric
        End If
        'do sort
        lvwResults(iCurrentTab).Sort
        If lLastCol = lCol Then
            'reset so we go back to ascending order next time
            lLastCol = 0
        Else
            'save the last column
            lLastCol = lCol
        End If
    End With
End Sub
Private Sub lvwResults_SelectionChange(Index As Integer, ByVal lRow As Long, ByVal lCol As Long)
    Dim frmDoc As IDevPadDocument
    'Dim Form As Form
    Dim bActive As Boolean
    Dim lLen As Long
    Dim lStart As Long
    Dim i As Long
    Dim sTag As String
    'could be -1 if triggered by my code!
    If lRow = -1 Then lRow = lvwResults(Index).SelectedRow
    If lRow = 0 Then Exit Sub
    
    sTag = lvwResults(iCurrentTab).CellTag(lRow, 1)
    
    If IsNumeric(sTag) Then
        'get the open doc...
        Set frmDoc = cDocuments.ItemByID(CLng(sTag))
    Else
        'load the file...
        Set frmDoc = cDocuments.LoadFile(sTag)
    End If
    
    If frmDoc Is Nothing Then
        cDialog.ShowWarning LoadResString(1241), "Results.SelChange" 'that doc is no longer open
        Exit Sub
    End If
    'any need to activate?
    If frmDoc.DocID = ActiveDoc.DocID Then bActive = True

    With frmDoc
        'activate the form if it isn't active
        If frmDoc.DocID <> ActiveDoc.DocID Then .SetFocus
        'wait a sec...!
        DoEvents
        'get the start pos
        lStart = lvwResults(iCurrentTab).CellText(lRow, 2) - 1
        
        If lvwResults(iCurrentTab).CellTag(lRow, 2) = "R" Then
            'item has been replaced...
            'get length accordingly
            lLen = Len(txtReplace(iCurrentTab))
            'and set replace button to 'undo'
            cmdReplace.Caption = LoadResString(1016) '"&Undo"
        Else
            'item has not been replaced...
            'get length accordingly
            lLen = Len(txtFind(iCurrentTab))
            'and set replace button to 'replace'
            cmdReplace.Caption = LoadResString(1013) '"&Replace"
        End If
        'set the richtextbox's selection
        .SelStart = lStart
        .SelLength = lLen
        '.SetSelection lStart, lStart + lLen
        DoEvents
        'return focus to here
        If Visible Then SetFocus
        'update form caption to reflect item change
        SetCaption
    End With
End Sub
Private Sub SetCaption()
    'sets the form's caption...
    '"Find Results -" & lvwResults(iCurrentTab).Rows & " items found"
    Caption = LoadResString(1106) & lvwResults(iCurrentTab).Rows & LoadResString(1107) ' " items found"
    'if the current tab has already been replaced then
    ' append 'and replaced' to caption
    If bReplaced(iCurrentTab) Then Caption = Caption & LoadResString(1108) '" & replaced"
    'Next, give info on the selected row
    'Caption = Caption & " - Item " & lRow & " selected"
    Caption = Caption & LoadResString(1115) & lvwResults(iCurrentTab).SelectedRow & LoadResString(1110)
    cmdReplace.Enabled = (lvwResults(iCurrentTab).Rows <> 0)
End Sub

Private Sub Form_Load()
    'load resource strings
    LoadResStrings Controls
    'build toolbar
    'load the standard toolbar
    With tbrMain
        'assign image list
        .ImageSource = CTBExternalImageList
        .SetImageList frmMainForm.vbalMain, CTBImageListNormal
        .CreateToolbar 16, True, True, True
        '.Wrappable = True
'        .AddButton "New", IndexForKey("NEWTEXT"), , , , CTBNormal Or CTBAutoSize, "NEW"
'        .AddButton "Rename", IndexForKey("EDIT"), , , , CTBNormal Or CTBAutoSize, "RENAME"
'        .AddButton "-", , , , , CTBSeparator, "SEP"
    End With
    'set up standard toolbar host
    With tbhMain
        .BorderStyle = etbhBorderStyleNone
        .Height = tbrMain.ToolbarHeight * Screen.TwipsPerPixelY
        'capture Main toolbar
        .Capture tbrMain
    End With
    'create a new result pane
    NewResultPane
End Sub
'adds a new result pane
Private Sub NewResultPane(Optional bPrompt As Boolean = False)
    Dim sText As String
    Dim i As Long
    'add new tab
    'add the toolbar items
    sText = "Results " & iTabCount + 1
    If bPrompt Then
        'prompt for text
        '"Please enter a name for the new result set"
        sText = cDialog.InputBox(LoadResString(1261), LoadResString(1262), sText)
        If sText = "" Then Exit Sub
    End If
    'add the button, and adjust toolbar height
    
    
    tbrMain.AddButton sText, IndexForKey("PAD"), , , sText, CTBNormal Or CTBAutoSize Or CTBCheckGroup, "DOC"
    Debug.Print tbrMain.ToolbarWidth
    tbhMain.Width = picHolder.ScaleWidth
   ' tbrMain.ListStyle = False
    tbrMain.ResizeToolbar
    picHolder.Height = tbrMain.ToolbarHeight2 * Screen.TwipsPerPixelX
    tbhMain.Move 0, 0, picHolder.ScaleWidth, picHolder.ScaleHeight
    
  '  If tbrMain.ToolbarWidth * Screen.TwipsPerPixelX > tbhMain.Width Then
        For i = 1 To tbrMain.ButtonCount
            tbrMain.ButtonTextEllipses(i) = True
            tbrMain.ButtonTextWrap(i) = False
            'tbrMain.ButtonWidth(i) = tbrMain.ButtonWidth(i) / 2
        Next i
  '  End If
    'adjust form height
    Height = 3000 + picHolder.Height
    If iTabCount <> 0 Then
        'load new controls to array
        Load lvwResults(iTabCount)
        Load picResults(iTabCount)
        Load txtFind(iTabCount)
        Load txtReplace(iTabCount)
        Load lblReplace(iTabCount)
        Load lblFind(iTabCount)
        'set the thin3d borders
        SetThin3DBorder txtFind(iTabCount).hWnd
        SetThin3DBorder txtReplace(iTabCount).hWnd
        
        'we need to use api for these...
        SetParent txtFind(iTabCount).hWnd, picResults(iTabCount).hWnd
        SetParent txtReplace(iTabCount).hWnd, picResults(iTabCount).hWnd
        'and can use vb for these...
        'don't ask me why!
        Set lvwResults(iTabCount).Container = picResults(iTabCount)
        Set lblReplace(iTabCount).Container = picResults(iTabCount)
        Set lblFind(iTabCount).Container = picResults(iTabCount)
        'show all the new items...
        lvwResults(iTabCount).Visible = True
        picResults(iTabCount).Visible = True
        txtFind(iTabCount).Visible = True
        txtReplace(iTabCount).Visible = True
        lblReplace(iTabCount).Visible = True
        lblFind(iTabCount).Visible = True
        'reset find/replace text
        ResetPane iTabCount
    End If
    'resize 'replaced' array
    ReDim Preserve bReplaced(0 To iTabCount)
    'disable this wretched thing...
    'it keeps resetting itself!
    'lvwResults(iTabCount).HeaderHotTrack = False
    'add the columns
    lvwResults(iTabCount).AddColumn , LoadResString(1111), , , 138  '"Found In"
    lvwResults(iTabCount).AddColumn , LoadResString(1112), , , 55 ' "Pos"
    lvwResults(iTabCount).AddColumn , LoadResString(1113), , , 39 '"Line"
    tbrMain.RaiseButtonClick iTabCount
    'increment tab count
    iTabCount = iTabCount + 1
End Sub
'creates a new result log
Public Sub NewResults(sText As String, sReplace As String, bReplace As Boolean, iResultTab As Integer)
    If iTabCount <= iResultTab Then
        'requested pane is out of range...
        'create new result pane
        NewResultPane
        'set result tab to the new one
        iResultTab = iTabCount - 1
    End If
    'select that tab
    tbrMain.RaiseButtonClick (iResultTab)
    'clear that result set
    lvwResults(iCurrentTab).Clear
    'set the find and replace text
    txtFind(iCurrentTab).Text = sText
    txtReplace(iCurrentTab).Text = sReplace
    txtReplace(iCurrentTab).Locked = False
    'set whether this result set has been replaced
    bReplaced(iCurrentTab) = bReplace
    'no redrawing...
    lvwResults(iCurrentTab).Redraw = False
End Sub
Public Sub AddItem(ByVal lStart As Long, ByVal sCaption As String, ByVal ID As Long, ByVal lLine As Long, Optional sFileName As String = "")
    'adds an item to the result set
    Dim lIndex As Long
    On Error Resume Next
    With lvwResults(iCurrentTab)
        lIndex = .Rows + 1
        'add a new row
        .AddRow
        'get the document's caption
        .CellText(lIndex, 1) = sCaption 'frmDocument.Caption
        'it's filename
        If sFileName <> "" Then
            .CellTag(lIndex, 1) = sFileName
        Else
            .CellTag(lIndex, 1) = ID
        End If
        'the start pos
        .CellText(lIndex, 2) = lStart + 1
        'whether the row has been replaced or not
        '(we default to the whole result sets value at first)
        .CellTag(lIndex, 2) = IIf(bReplaced(iCurrentTab), "R", "")
        'get the line the text was found on
        .CellText(lIndex, 3) = lLine 'frmDocument.GetLineFromChar(lStart) + 1
    End With
End Sub
'Complete adding new results
Public Sub Complete()
    'redraw
    lvwResults(iCurrentTab).Redraw = True
    'update the form's caption
    SetCaption
    'show the form...
    If Visible = False Then Show , frmMainForm
    'select the tab for this result set
    tbrMain.RaiseButtonClick iCurrentTab
    'tbsResults_TabClick (iCurrentTab + 1)
End Sub

Private Sub tbrMain_ButtonClick(ByVal lButton As Long)
    Dim i As Integer
    ' show and enable the selected tab's controls
    ' and hide and disable all others
    iCurrentTab = lButton
    For i = 0 To tbrMain.ButtonCount - 1
        If i = lButton Then
            picResults(i).Left = 75
            picResults(i).Enabled = True
            picResults(i).ZOrder (0)
        Else
            picResults(i).Left = -20000
            picResults(i).Enabled = False
        End If
    Next
    'ensure correct item is checked
    tbrMain.ButtonChecked(iCurrentTab) = True
    'select the first item if there are rows in the set, and there is
    'no selected row
    If lvwResults(iCurrentTab).Rows <> 0 And lvwResults(iCurrentTab).SelectedRow = 0 Then lvwResults(iCurrentTab).CellSelected(1, 1) = True      ' = 1
    'lvwResults_SelectionChange iCurrentTab, -1, -1
    'update the form's caption
    SetCaption
End Sub

