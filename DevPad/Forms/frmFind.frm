VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   2910
   ClientLeft      =   2655
   ClientTop       =   3555
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
   Icon            =   "frmFind.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2910
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4455
      Top             =   1515
   End
   Begin VB.ListBox lstDocs 
      Height          =   960
      Left            =   270
      Style           =   1  'Checkbox
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   435
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2295
      Index           =   0
      Left            =   90
      ScaleHeight     =   2295
      ScaleWidth      =   3930
      TabIndex        =   12
      Top             =   60
      Width           =   3930
      Begin VB.TextBox txtFind 
         Height          =   495
         Left            =   720
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         ToolTipText     =   "Text to Find. Press Ctrl+Tab to enter the tab character."
         Top             =   0
         WhatsThisHelpID =   20000
         Width           =   3045
      End
      Begin VB.PictureBox picExpandFind 
         BorderStyle     =   0  'None
         DrawStyle       =   5  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         MousePointer    =   8  'Size NW SE
         Picture         =   "frmFind.frx":000C
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   14
         ToolTipText     =   "Left click to expand Find box. Right click to restore to original size."
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txtReplace 
         Height          =   495
         Left            =   720
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         ToolTipText     =   "Replace Text. Press Ctrl+Tab to enter the tab character."
         Top             =   585
         Width           =   3045
      End
      Begin VB.PictureBox picExpandReplace 
         BorderStyle     =   0  'None
         DrawStyle       =   5  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         MousePointer    =   8  'Size NW SE
         Picture         =   "frmFind.frx":0156
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   13
         ToolTipText     =   "Left click to expand Replace box. Right click to restore to original size."
         Top             =   960
         Width           =   255
      End
      Begin DevPad.vbwFlatButton cmdMoreFind 
         Height          =   285
         Left            =   3780
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   210
         Width           =   150
         _ExtentX        =   265
         _ExtentY        =   503
      End
      Begin DevPad.vbwFlatButton cmdMoreReplace 
         Height          =   285
         Left            =   3780
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   810
         Width           =   150
         _ExtentX        =   265
         _ExtentY        =   503
      End
      Begin VB.Frame fraOptions 
         Caption         =   "1047"
         Height          =   1170
         Left            =   15
         TabIndex        =   15
         Top             =   1080
         Width           =   3855
         Begin DevPad.vbwFlatButton cmdWhere 
            Height          =   285
            Left            =   3615
            TabIndex        =   22
            Top             =   240
            Width           =   150
            _ExtentX        =   265
            _ExtentY        =   503
         End
         Begin VB.ComboBox cboWhere 
            Height          =   315
            ItemData        =   "frmFind.frx":02A0
            Left            =   1680
            List            =   "frmFind.frx":02B0
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   210
            Width           =   1890
         End
         Begin VB.CheckBox chkOutput 
            Caption         =   "1095"
            Height          =   252
            Left            =   1680
            TabIndex        =   6
            ToolTipText     =   "Finds the specified text after the cursor position"
            Top             =   840
            Width           =   1920
         End
         Begin VB.CheckBox chkWholeWord 
            Caption         =   "1093"
            Height          =   252
            Left            =   120
            TabIndex        =   4
            ToolTipText     =   "Find the whole word only"
            Top             =   840
            Width           =   1452
         End
         Begin VB.CheckBox chkMatchCase 
            Caption         =   "1092"
            Height          =   252
            Left            =   120
            TabIndex        =   3
            ToolTipText     =   "Case Sensitivity"
            Top             =   600
            Width           =   1455
         End
         Begin VB.CheckBox chkFindOnwards 
            Caption         =   "1094"
            Height          =   252
            Left            =   1680
            TabIndex        =   5
            ToolTipText     =   "Finds the specified text after the cursor position"
            Top             =   600
            Width           =   1815
         End
         Begin DevPad.vbwFlatButton cmdLog 
            Height          =   285
            Left            =   3615
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   795
            Width           =   150
            _ExtentX        =   265
            _ExtentY        =   503
         End
         Begin VB.Label lblLabel 
            Caption         =   "1091"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   255
            Width           =   1485
         End
      End
      Begin VB.Label lblFind 
         Caption         =   "1089"
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblReplace 
         Caption         =   "1090"
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "1014"
      Height          =   345
      Left            =   4080
      TabIndex        =   10
      Top             =   915
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "1013"
      Height          =   345
      Left            =   4080
      TabIndex        =   9
      Top             =   495
      Width           =   1200
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "1011"
      Default         =   -1  'True
      Height          =   345
      Left            =   4080
      TabIndex        =   8
      Top             =   75
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "1000"
      Height          =   345
      Left            =   4080
      TabIndex        =   11
      Top             =   2505
      Width           =   1200
   End
   Begin VB.Image imgUp 
      Height          =   60
      Left            =   5055
      Picture         =   "frmFind.frx":02CC
      ToolTipText     =   "Hide Replace"
      Top             =   2310
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      Caption         =   "1240"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   135
      MouseIcon       =   "frmFind.frx":0386
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2550
      Width           =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   75
      X2              =   5280
      Y1              =   2415
      Y2              =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   2
      X1              =   75
      X2              =   5265
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Menu mnuLog 
      Caption         =   "&Log"
      Visible         =   0   'False
      Begin VB.Menu mnuLogResults 
         Caption         =   "Output To Results 1"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuLogResultsNew 
         Caption         =   "New..."
      End
   End
   Begin VB.Menu mnuFindTOP 
      Caption         =   "&FindHistory"
      Visible         =   0   'False
      Begin VB.Menu mnuFind 
         Caption         =   "&Selected Text"
         Index           =   0
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Current &Line"
         Index           =   1
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Current &Word"
         Index           =   2
      End
      Begin VB.Menu mnuFind 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFindHistory 
         Caption         =   "Empty"
         Enabled         =   0   'False
         Index           =   0
      End
   End
   Begin VB.Menu mnuReplaceTOP 
      Caption         =   "&ReplaceHistory"
      Visible         =   0   'False
      Begin VB.Menu mnuReplace 
         Caption         =   "&Selected Text"
         Index           =   0
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "Current &Line"
         Index           =   1
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "Current &Word"
         Index           =   2
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuReplaceHistory 
         Caption         =   "Empty"
         Enabled         =   0   'False
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmFind"
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
Private Const WM_SIZE = &H5
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTBOTTOMRIGHT = 17
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private bShowReplace        As Boolean
Private cFlatCombo()        As clsFlatCombo
Private cFlatOpt()          As clsFlatOpt

Private lCount              As Long 'number of occurances found...
Private bCancel             As Boolean 'allows us to skip out of long processes
Private bBusy               As Boolean 'tells us we are in a long process!
Private bChanged            As Boolean

'displays this dialog...
Public Sub ShowFind(bShowReplace As Boolean)
Dim sSelText As String
Dim bReset   As Boolean
    'abort if starting up...
    If bStartUp Then Exit Sub
    'if there is a document open, and there is
    'selected text, use that as default find text
'    If DocOpen Then
'        sSelText = ActiveDoc.SelText
'        If sSelText <> "" Then
'            If Len(sSelText) < 200 Then txtFind.Text = ActiveDoc.SelText
'            'cboWhere.ListIndex = 1
'        Else
'            bReset = True
'        End If
'    Else
'        bReset = True
'    End If
'    If bReset Then
'        txtFind.Text = ""
'        If cboWhere.ListIndex = 1 Then cboWhere.ListIndex = 0
'    End If
    'display form
    If Visible = False Then Show , frmMainForm
    If txtFind.Visible Then txtFind.SetFocus
    If bShowReplace Then SetReplace
End Sub
Private Function GetOptions() As FindConstants
    'build the options for searching in the richtextbox
    'match case?
    GetOptions = GetOptions Or (Abs(chkMatchCase) * rtfMatchCase)
    'whole word?
    GetOptions = GetOptions Or (Abs(chkWholeWord) * rtfWholeWord)
End Function
'load results window, and position below find window...
Private Sub LoadResultsWindow()
    '(only position if the form has not been loaded yet!)
    If IsLoaded("frmFindResults") = False Then
        Load frmFindResults
        frmFindResults.Top = Top + Height
        frmFindResults.Left = Left
    End If
End Sub
Private Sub cboWhere_Click()
    If cboWhere.ListIndex = 3 Then
        'If 'Selected Documents...' is chosen...
        ListDocs
    End If
    SetCaption
End Sub
Public Sub ListDocs()
    Dim i As Long
    'If 'Selected Documents...' is chosen...
    ' Clear the list
    lstDocs.Clear
    ' load the forms
    For i = 1 To cDocuments.Count
        lstDocs.AddItem cDocuments.Item(i).DocumentCaption
    Next
End Sub
'update command buttons....
Private Sub chkOutput_Click()
    SetCaption
End Sub
Private Sub SetCaption()
    'enable the arrow next to cboWhere if 'Selected Documents...' is chosen
    cmdWhere.Enabled = (cboWhere.ListIndex = 3)
    'enable Find onwards, provided 'Selected Text' is not chosen
    chkFindOnwards.Enabled = (cboWhere.ListIndex <> 1)
    If chkOutput.Value = 1 Then
        'we want to output to a log... so change Find to Find All
        cmdFind.Caption = LoadResString(1012) '"&Find All"
        If bShowReplace = True Then cmdReplace.Enabled = False
    Else
        'otherwise, reset to Find...
        cmdFind.Caption = LoadResString(1011) '"&Find"
    End If
    UpdateButtons
End Sub

Public Sub cmdFind_Click()
On Error GoTo ErrHandler
    Dim ActiveForm As IDevPadDocument
    Dim lPos        As Long
    Dim lStart      As Long
    'if we are searching in current selection,
    'we need to remember where the end of the sel is...
    Static lEnd     As Long

    'if there are no documents open, abort...
    If DocOpen = False Then Exit Sub
    If cmdFind.Enabled = False Then Exit Sub
    Set ActiveForm = ActiveDoc
    If cmdFind.Caption = LoadResString(1012) Then 'Find All
        ' we want to find all occurances
        FindAll False
        Exit Sub
    ElseIf cmdFind.Caption = LoadResString(1011) Then 'Find
        ' start find.... (first time!)
        ' add the search text to find history...
        frmMainForm.cFindHistory.AddNewItem txtFind.Text
        ' and re-load the find history...
        LoadFindReplaceHistory True
        ' get the first document in the search range
        If LoopDocs(True) = False Then
            FoundAll
            Exit Sub
        End If
        'get the start/end range from the options selected...
        GetSearchRange lStart, lEnd
    Else
        ' find next....
        lStart = ActiveForm.SelStart + ActiveForm.SelLength
    End If

    ' set find count to -1 (ignore)
    ' for individual finds, this is not used
    lCount = -1
    
FindNext:
    
    With ActiveForm
        '.Busy = True
        'use this so that we don't get the find form deactivated as the selection is made...
        '.NoEnterFocus = True
        'find the next occurance
        If lEnd = 0 Then lEnd = -1
        lPos = .Find(txtFind.Text, lStart, lEnd, GetOptions)
        'reset...
        '.NoEnterFocus = False
        '.Busy = False
        'SetFocus
    End With
    If (lPos >= 0) Then
        'we have found something...
        cmdFind.Caption = LoadResString(1085) '"&Find Next"
    Else
        'we have not found anything...
        If cboWhere.ListIndex = 2 Or cboWhere.ListIndex = 3 Then
            'All Documents or Specified Documents selected...
            'Go to next document...
            If LoopDocs(False) Then
                ' successfully activated next document
                ' try searching this one
                ActiveDoc.SelStart = 0
                ' reset these...
                lStart = 0
                lEnd = 0
                Set ActiveForm = ActiveDoc
                'go back to the beginning!
                GoTo FindNext
            Else
                'searched all documents....
                FoundAll
                lEnd = 0
            End If
        Else
            'finished searching current document....
            FoundAll
            lEnd = 0
        End If
    End If
    DoEvents
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Find.FindAll"
End Sub
Private Sub GetSearchRange(ByRef lStart As Long, ByRef lEnd As Long)
    If cboWhere.ListIndex = 1 Then
        ' search selected text
        ' set start and end positions...
        lStart = ActiveDoc.SelStart
        lEnd = ActiveDoc.SelLength + lStart
    Else
        If chkFindOnwards Then
            ' start from cursor position
            lStart = ActiveDoc.SelStart
        Else
            ' start from beginning...
            lStart = 0
        End If
        lEnd = 0
    End If
End Sub
'Displays a message box when searching has finished
Private Sub FoundAll()
Dim sMsg As String
Dim sWhere As String
    'disable update timer
    tmrUpdate.Enabled = False
    If cboWhere.ListIndex = 2 Or cboWhere.ListIndex = 3 Then
        sWhere = LoadResString(1087) ' specified documents
    Else
        sWhere = LoadResString(1088) ' current document
    End If
    ' re-enable the controls
    EnableItems True
    ' no more docs to search
    ' reset the find caption
    SetCaption
    'allow the activedoc to redraw
    On Error Resume Next
    ActiveDoc.Redraw = True
    ' 1086 = Finished searching
    sMsg = LoadResString(1086) & sWhere & "."
    ' 1229 = Found
    If lCount <> -1 Then sMsg = sMsg & " " & LoadResString(1229) & " " & lCount & " " & LoadResString(1230) & "."
    cDialog.ShowWarning sMsg, "Find.FoundAll", "Search Complete"
End Sub
'this proc allows us to loop through the documents in the
'selected range...
Private Function LoopDocs(bFirstDoc As Boolean) As Boolean
    Static i As Integer
    Static nStartDoc As Integer

    Dim bFindAnyDoc As Boolean
    Dim bValid As Boolean
    
    If cboWhere.ListIndex = 1 Or cboWhere.ListIndex = 0 Then
        ' search the active form or selection...
        LoopDocs = True
        Exit Function
    End If
    ' if we are searching a specified range of documents,
    ' we always want to start with the active document first,
    ' if possible... so we search for the active doc first,
    ' and then move on to any doc... (when bFindAnyDoc=True)
    bFindAnyDoc = False
    
    If bFirstDoc = True Then
        ' find the current documents item value, and start there.
        For i = 1 To cDocuments.Count
            ' we only want documents
            'If TypeOf Forms(i) Is frmDocument Then
            With cDocuments.Item(i)
                ' we are looking for a specific document
                If cboWhere.ListIndex = 3 Then
                    If .DocID = ActiveDoc.DocID And Not bFindAnyDoc Then
                        ' it is the current form
                        If IsChecked(.DocumentCaption) Then
                            ' it is checked. OK!
                            GoTo AcceptDoc
                        Else
                            ' start again
                            i = -1
                            ' doc is not checked. Find any doc
                            bFindAnyDoc = True
                            DoEvents
                        End If
                    ElseIf IsChecked(.DocumentCaption) And bFindAnyDoc Then
                        'it is checked, and we are not looking for
                        GoTo AcceptDoc
                    End If
                ElseIf .DocID = ActiveDoc.DocID Then
                    'it's the active doc... accept it
                    GoTo AcceptDoc
                End If
            End With
            'End If
        Next
        nStartDoc = -1
       ' SetFocus
        Exit Function
AcceptDoc:
       
        ' this document is checked, return it
        nStartDoc = i
        'activate the form
        With cDocuments.Item(i)
            '.NoEnterFocus = True
            SendMessage frmMainForm.GetMDIClienthWnd, WM_MDIACTIVATE, .DocumenthWnd, 0
            '.NoEnterFocus = False
        End With
        'reactivate this form...
        DoEvents
        Me.SetFocus
        'increment starting pos...
        i = i + 1
        LoopDocs = True
        Exit Function
    End If
    ' loop through all the documents
    If i > cDocuments.Count Then i = 1
    Do While i <> nStartDoc And nStartDoc <> -1
        'If Forms(i).Name = "frmDocument" Then
            bValid = True
            If cboWhere.ListIndex = 4 Then 'spec doc
                If Not IsChecked(cDocuments.Item(i).DocumentCaption) Then
                    ' if the current document is not checked then ignore it
                    bValid = False
                End If
            End If
            If bValid Then
                ' document found
                ' activate it
                On Error Resume Next
                cDocuments.Item(i).SetFocus
                DoEvents
                LoopDocs = True
                i = i + 1
                SetFocus
                Exit Function
            End If
        'End If
        If i = cDocuments.Count - 1 Then
            i = 1
        Else
            i = i + 1 ' next document
        End If
    Loop
    LoopDocs = False
    i = 0 ' reset
End Function
'returns whether the specified item is checked in lstDocs listbox
Private Function IsChecked(sText As String) As Boolean
    Dim i As Long
    'loop through all the items
    For i = 0 To lstDocs.ListCount - 1
        If lstDocs.Selected(i) And lstDocs.List(i) = sText Then
            'item is checked...
            IsChecked = True
            Exit Function
        End If
    Next
End Function
'Displays the list of log items
Private Sub cmdLog_Click()
Dim i As Long
    txtFind.SetFocus
    If IsLoaded("frmFindResults") Then
        'only bother listing if they have changed....
        If frmFindResults.ResultsCount <> mnuLogResults.Count Then
            For i = 1 To mnuLogResults.Count - 1
                Unload mnuLogResults(i)
            Next
            For i = 0 To frmFindResults.ResultsCount - 1 Step 1
                'load a new menu item...
                If i <> 0 Then Load mnuLogResults(i)
                'set it's caption
                mnuLogResults(i).Caption = frmFindResults.tbrMain.ButtonCaption(i) ' & i + 1 '- 1
                'make sure it is not checked
                mnuLogResults(i).Checked = False
            Next i
        End If
    End If
    'display the popup menu
    PopupMenu mnuLog, , cmdLog.Left + cmdLog.Width + picScreen(0).Left + fraOptions.Left, cmdLog.Top + picScreen(0).Top + fraOptions.Top - 20
End Sub

'Replace the current item
Private Sub cmdReplace_Click()
    If bShowReplace = False Then
        'if the replace textbox is not shown, display it...
        SetReplace
    Else
        Static sLastItem As String
        Dim bMatch As Boolean
        'abort if no documents are open
        If DocOpen = False Then Exit Sub
        With ActiveDoc
            'check that the selected text matches the find text
            If chkMatchCase.Value = 1 Then
                'Case sensitive...
                bMatch = (.SelText = txtFind.Text)
            Else
                'No case sensitivity...
                bMatch = (LCase$(.SelText) = LCase$(txtFind.Text))
            End If
            'and we have started our search...
            If bMatch And cmdFind.Caption <> LoadResString(1011) Then 'find
                'we don't want to lose the focus...
              '  .NoEnterFocus = True
                'update the text
                .SelText = txtReplace
                '.NoEnterFocus = False
                're-activate this window...
                SetFocus
                'if the last item we searched for isn't this one,
                'add it to the replace history
                If sLastItem <> txtReplace.Text Then
                    frmMainForm.cReplaceHistory.AddNewItem txtReplace.Text
                    're-load find/replace history
                    LoadFindReplaceHistory True
                    'save this...
                    sLastItem = txtReplace.Text
                End If
            End If
        End With
        'now find the next item...
        Call cmdFind_Click
    End If
End Sub
'returns the result tab that is selected...
Private Function GetResTab() As Long
    Dim i As Long
    'find the selected menu item..
    For i = 0 To mnuLogResults.Count - 1
        If mnuLogResults(i).Checked = True Then Exit For
    Next
    'increment the index by one
    'and return it...
    GetResTab = i '+ 1
End Function
'enable/disable all the items in the window...
'used when doing a long find/replace action
Private Sub EnableItems(bEnable As Boolean)
Dim i As Long
    'change the Close to Cancel, and Cancel to Close!
    cmdCancel.Caption = IIf(bEnable, LoadResString(1000), LoadResString(1002))
    For i = 0 To Controls.Count - 1
        If Controls(i).Name <> "cmdCancel" Then
            'provided the control isn't cmdCancel, and a valid type, disable it
            Select Case TypeName(Controls(i))
            Case "CommandButton", "TextBox", "vbwFlatButton", "CheckBox", "ComboBox"
                Controls(i).Enabled = bEnable
            End Select
        End If
    Next i
    'we are busy if we have been told to disable everything...
    bBusy = Not bEnable
    bCancel = False
End Sub
'This sub searches for more than one item at once!
Private Sub FindAll(bReplace As Boolean)
On Error GoTo ErrHandler
    Dim FindForm    As IDevPadDocument
    Dim lPos        As Long
    Dim lStart      As Long
    Dim lInitStart  As Long
    Dim lEnd        As Long
    Dim bIsCrLf     As Boolean
    Dim sCaption    As String
    Dim cCursor     As clsCursor
    'abort if no documents are open
    If DocOpen = False Then Exit Sub
    'if we are replacing, add the replace text to the history
    If bReplace Then frmMainForm.cReplaceHistory.AddNewItem txtReplace.Text
    'add find text to the history
    frmMainForm.cFindHistory.AddNewItem txtFind.Text
    're-load the history
    LoadFindReplaceHistory True
    ' find the first document in range...
    If LoopDocs(True) = False Then
        'no docs
        FoundAll
        Exit Sub
    End If
    'are we outputting results to result window?
    If chkOutput Then
        'load results window
        LoadResultsWindow
        'create a new result set...
        If bReplace Then
            frmFindResults.NewResults txtFind, txtReplace, True, GetResTab()
        Else
            frmFindResults.NewResults txtFind, "", False, GetResTab()
        End If
    End If
    'disable everything to prevent users clicking what
    'they shouldn't!
    EnableItems False
    'reset find counter
    lCount = 0
    tmrUpdate.Enabled = True
    'assign to current document
    Set FindForm = ActiveDoc
    ' set the end position
    GetSearchRange lStart, lEnd
    lInitStart = lStart
    ' change mouse cursor
    Set cCursor = New clsCursor
    cCursor.SetCursor vbHourglass
    ' gets around a bug in RichTextBox control. If the selected text
    ' contains a CrLf, then to delete it, you need to use SelText = ""
    ' then, replace it with the replace text.
    If InStr(1, txtFind, vbCrLf) <> 0 Then bIsCrLf = True
StartFind:
    With FindForm
        ' prevent redrawing in current doc..
        .Redraw = False
        Do
            DoEvents
            'Check to see if operation has been cancelled...
            If bCancel Then
                'abort...
                FoundAll
                Exit Do
            End If
            'find the next item...
            'what can we set lEnd to so the richtextbox ignores it,
            'rather than having to write different code ommiting it?!
            If lEnd = 0 Then lEnd = -1
            lPos = .Find(txtFind.Text, lStart, lEnd, GetOptions + rtfNoHighlight)
            If (lPos >= 0) Then
                'item found...
                If bReplace Then
                    'replace the text...
                    .SelStart = lPos
                    .SelLength = Len(txtFind)
                    'if it is a crlf, clear selection first...
                    If bIsCrLf Then .SelText = ""
                    'then change selection to replace text
                    .SelText = txtReplace
                    'increment start pos so we don't keep finding
                    'the same item!
                    lStart = lPos + Len(txtReplace) '+ 1
                Else
                    'increment start pos so we don't keep finding
                    'the same item!
                    lStart = lPos + Len(txtFind)
                End If
                'if we are outputting, add it to the results log
                If chkOutput Then frmFindResults.AddItem lPos, .DocumentCaption, .DocID, .LineFromChar(lStart)
                'increment the find count (displayed in msgbox at end)
                lCount = lCount + 1
            Else
                're-enable redrawing...
                .Redraw = True
               ' .NoEnterFocus = False
                'are we searching all documents/specified documents?
                If cboWhere.ListIndex = 2 Or cboWhere.ListIndex = 4 Then
                    'find next document in range
                    If LoopDocs(False) Then
                        ' successfully activated next document
                        ' try searching this one
                        .SelStart = 0
                        Set FindForm = ActiveDoc
                        ' reset start/end values...
                        lStart = 0
                        lEnd = 0
                        GoTo StartFind
                    Else
                        'searched everything... exit loop
                        FoundAll
                        Exit Do
                    End If
                Else
                    'restore selection, if needed
                    ActiveDoc.SelStart = lInitStart
                    'If cboWhere.ListIndex = 1 Then ActiveDoc.SelLength = lEnd - lInitStart
                    'searched everything... exit loop
                    FoundAll
                    Exit Do
                End If
            End If
        Loop
    End With
    If chkOutput Then
        'display the find/replace log
        frmFindResults.Complete
    End If
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Find.FindAll"
End Sub

Private Sub cmdReplaceAll_Click()
    'Replace all occurances...
    FindAll True
End Sub

Private Sub cmdWhere_Click()
    'Display listbox for choice of documents to search...
    lstDocs.Left = cmdWhere.Left + fraOptions.Left + picScreen(0).Left - lstDocs.Width
    lstDocs.Top = cmdWhere.Top + fraOptions.Top + picScreen(0).Top
    lstDocs.Visible = True
    'Ensure it has the focus... once it has lost the focus
    'it will hide itself
    lstDocs.SetFocus
End Sub

Private Sub Form_Activate()
Dim sText As String
    If (txtFind.Text = "" Or bChanged = False Or cmdFind.Caption = LoadResString(1011)) And DocOpen Then
        sText = ActiveDoc.SelText
        If InStr(1, sText, vbCrLf) = 0 And sText <> "" Then
            'adjust to new search text...
            txtFind.Text = sText
            bChanged = False
            'reset
            If cboWhere.ListIndex = 1 Then cboWhere.ListIndex = 0
        ElseIf sText <> "" Then
            'assume search selected text...
            cboWhere.ListIndex = 1
        Else
            'reset
            If cboWhere.ListIndex = 1 Then cboWhere.ListIndex = 0
        End If
    End If
End Sub

Private Sub lblHelp_Click()
    cDialog.ShowHelpTopic 7, hWnd
End Sub

Private Sub lstDocs_LostFocus()
    'popup control has lost the focus... hide it
    lstDocs.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If bBusy Then
        'if we are busy, abort unload
        Cancel = -1
    ElseIf UnloadMode = vbFormControlMenu Then
        'don't unload the form if the user has tried to close it... just hide it!
        Hide
        Cancel = -1
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'remove subclassing...
    DetachMessage Me, txtFind.hWnd, WM_SIZE
    DetachMessage Me, txtReplace.hWnd, WM_SIZE
    'remove thin 3d effect
    RestoreControls cFlatCombo, cFlatOpt
End Sub
Private Sub imgUp_Click()
    'restore to just displaying the find button...
    '(ie no replace!)
    SetFind
End Sub
'Subclassing for WM_SIZE message, which allows textbox
'to be resized....
Private Property Let ISubclass_MsgResponse(ByVal RHS As vbwSubClass.EMsgResponse)
End Property
Private Property Get ISubclass_MsgResponse() As vbwSubClass.EMsgResponse
ISubclass_MsgResponse = emrPreprocess
End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case iMsg
    Case WM_SIZE
        If hWnd = txtFind.hWnd Then
            'move the resizing arrows to the corner of the
            'find textbox as we resize...
            picExpandFind.Top = 0 + txtFind.Height - 195
            picExpandFind.Left = 720 + txtFind.Width - 210
        ElseIf hWnd = txtReplace.hWnd Then
            'move the resizing arrows to the corner of the
            'replace textbox as we resize...
            picExpandReplace.Top = 600 + txtReplace.Height - 195
            picExpandReplace.Left = 720 + txtReplace.Width - 210
        End If
    End Select
End Function

Private Sub mnuFind_Click(Index As Integer)
    'Command selected...
    'change the find text to the one clicked on the menu
    InsertAuto txtFind, mnuFind(Index).Caption
End Sub
Private Sub mnuFindHistory_Click(Index As Integer)
    'A recent item has been selected...
    'change the find text to the one clicked on the menu
    InsertAuto txtFind, mnuFindHistory(Index).Tag
End Sub
Private Sub mnuReplace_Click(Index As Integer)
    'Command selected...
    'change the replace text to the one clicked on the menu
    InsertAuto txtReplace, mnuReplace(Index).Caption
End Sub
Private Sub mnuReplaceHistory_Click(Index As Integer)
    'A recent item has been selected...
    'change the replace text to the one clicked on the menu
    InsertAuto txtReplace, mnuReplace(Index).Tag
End Sub
Private Sub mnuLogResults_Click(Index As Integer)
Dim i As Long
    'check the selected item...
    For i = 0 To mnuLogResults.Count - 1
        mnuLogResults(i).Checked = (i = Index)
    Next
    'uncheck the new item...
    mnuLogResultsNew.Checked = False
End Sub
Private Sub mnuLogResultsNew_Click()
    Dim i As Long
    'check the new item...
    mnuLogResultsNew.Checked = True
    'uncheck all the other items..
    For i = 0 To mnuLogResults.Count - 1
        mnuLogResults(i).Checked = False
    Next
End Sub
'Gets a specified bit of text...
Private Sub InsertAuto(txtText As TextBox, sString As String)
    Dim sText As String
    Select Case sString
    Case "&Selected Text"
        'return the selected text...
        sText = ActiveDoc.SelText
    Case "Current &Line"
        'the current line text...
        sText = ActiveDoc.LineText()
    Case "Current &Word"
        'the current word....
        sText = GetCurrentWord
    Case Else
        'the string we were sent...
        sText = sString
    End Select
    'update the textbox
    txtText.Text = sText
End Sub
'allows resizing of the find textbox...
Private Sub picExpandFind_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'release mouse capture...
    ReleaseCapture
    If Button = vbRightButton Then
        ' restore
        txtFind.Height = 495
        txtFind.Width = 3045
        picExpandFind_MouseDown -1, 0, 0, 0
    ElseIf Button = vbLeftButton Then
        ' reset other box's size...
        picExpandReplace_MouseDown vbRightButton, 0, 0, 0
        ' activate textbox...
        txtFind_GotFocus
        ' start resizing...
        SendMessage txtFind.hWnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, ByVal &O0
    End If
End Sub
'allows resizing of the replace textbox...
Private Sub picExpandReplace_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'release mouse capture...
    ReleaseCapture
    If Button = vbRightButton Then
        ' restore
        txtReplace.Height = 495
        txtReplace.Width = 3045
        picExpandReplace_MouseDown -1, 0, 0, 0
    ElseIf Button = vbLeftButton Then
        'reset other box's size...
        picExpandFind_MouseDown vbRightButton, 0, 0, 0
        txtReplace_GotFocus
        ' start resizing...
        SendMessage txtReplace.hWnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, ByVal &O0
    End If
End Sub

Private Sub tmrUpdate_Timer()
    'update status bar to give indication of progress!
    SetStatusBar LoadResString(1229) & " " & lCount & " " & LoadResString(1230) & "."
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyTab And Shift = vbCtrlMask Then
'        KeyCode = 0
'        txtFind.SetFocus
'        'use the indent string instead...
'        txtFind.SelText = ActiveDoc.IndentString
'    End If
End Sub

Private Sub cmdCancel_Click()
    If cmdCancel.Caption = LoadResString(1002) Then '"Cancel"
        'if we are busy, set cancel flag, and wait!
        bCancel = True
    Else
        'otherwise, hide the form...
        Hide
        'and reset caption
        SetCaption
    End If
End Sub

'Display the list of recent find items...
Private Sub cmdMoreFind_Click()
    PopupMenu mnuFindTOP, , cmdMoreFind.Left + cmdMoreFind.Width + picScreen(0).Left, cmdMoreFind.Top + picScreen(0).Top - 20
End Sub
'Display the list of recent replace items...
Private Sub cmdMoreReplace_Click()
    PopupMenu mnuReplaceTOP, , cmdMoreReplace.Left + cmdMoreReplace.Width + picScreen(0).Left, cmdMoreReplace.Top + picScreen(0).Top - 20
End Sub

Private Sub txtFind_GotFocus()
    'simulate change just in case mainform find combo
    'has been changed whilst we weren't looking!
    txtFind_Change
    'ensure textbox stays on top
    '(it might not be after replacebox has been resized)
    picExpandFind.ZOrder
    txtFind.ZOrder
End Sub
Private Sub txtReplace_GotFocus()
    'ensure textbox stays on top
    '(it might not be after findbox has been resized)
    picExpandReplace.ZOrder
    txtReplace.ZOrder
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
Dim i        As Long
Dim sSelText As String
    'capture WM_SIZE message for when resizing textbox..
    AttachMessage Me, txtFind.hWnd, WM_SIZE
    AttachMessage Me, txtReplace.hWnd, WM_SIZE
    
    'Attach thin-3d style combo boxes and checkboxes
    MakeControlsFlat Controls, cFlatCombo, cFlatOpt
    
    'populate find/replace history
    LoadFindReplaceHistory False
    'set to find only layout...
    SetFind
    'load resource strings
    LoadResStrings Controls
    'set selected item
    cboWhere.ListIndex = 0
    
    'simulate change to update controls...
    txtFind_Change
    
   ' picExpandFind_MouseDown -1, 0, 0, 0
   ' picExpandReplace_MouseDown -1, 0, 0, 0
   ' SetFind
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Find.Form_Load"
End Sub

Private Sub txtFind_Change()
    Static sLastValue
    bChanged = True
    'if the textbox is empty, disable the find button.
    UpdateButtons
    'update combo on main form
    frmMainForm.cboFind.Text = txtFind.Text
    'revert to Find if the value has changed
    If txtFind.Text <> sLastValue Then
        sLastValue = txtFind
        SetCaption
    End If
End Sub
Private Sub UpdateButtons()
    'if find is empty, disable everything...
    If txtFind.Text = Empty Then
        cmdFind.Enabled = False
        'enable replace if we are only showing find part of dialog...
        cmdReplace.Enabled = Not (bShowReplace) 'Or (chkOutput.Value = 0)
        cmdReplaceAll.Enabled = False
    Else
        'enable everything
        cmdFind.Enabled = True
        cmdReplace.Enabled = (bShowReplace = False) Or (chkOutput.Value <> 1)
        cmdReplaceAll.Enabled = True
    End If
End Sub

'changes dialog for displaying Replace...
Private Sub SetReplace()
    'we are showing replace...
    bShowReplace = True
    'move results window to below find window if it hasn't been moved yet
    If IsLoaded("frmFindResults") Then
        If frmFindResults.Top = Top + Height Then frmFindResults.Top = Top + 3285 'new height'Height
    End If
    'change form height
    Height = 3285
    'options pos
    fraOptions.Top = 1080
    'holder pos
    picScreen(0).Height = 2295
    'cancel button pos
    cmdCancel.Top = 2505
    'bottom stuff...
    lblHelp.Top = 2550
    Line1(2).Y1 = 2430
    Line1(2).Y2 = 2430
    Line1(1).Y1 = 2415
    Line1(1).Y2 = 2415
    'Show items associated with Replace box...
    ShowItems True
    'Change the replace caption
    cmdReplace.Caption = LoadResString(1013)
End Sub
'adjusts dialog for displaying Find...
Private Sub SetFind()
    'we are not showing replace dialog
    bShowReplace = False
    'move results window if it hasn't been moved...
    If IsLoaded("frmFindResults") Then
        If frmFindResults.Top = Top + Height Then frmFindResults.Top = Top + 2740 'new height'Height
    End If
    'set the various positions...
    Height = 2740
    fraOptions.Top = 520
    picScreen(0).Height = 1735
    cmdCancel.Top = 1955
    lblHelp.Top = 2000
    Line1(2).Y1 = 1855
    Line1(2).Y2 = 1855
    Line1(1).Y1 = 1845
    Line1(1).Y2 = 1845
    'hide items associated with replace box...
    ShowItems False
    'change replace caption
    cmdReplace.Caption = LoadResString(1013) & "..."
End Sub
Private Sub ShowItems(bShow As Boolean)
    'hides/shows items associated with Replace box
    txtReplace.Visible = bShow
    lblReplace.Visible = bShow
    picExpandReplace.Visible = bShow
    cmdReplaceAll.Visible = bShow
    cmdMoreReplace.Visible = bShow
    imgUp.Visible = bShow
    'ensure we update the buttons
    txtFind_Change
End Sub
'loads the find/replace history...
Private Sub LoadFindReplaceHistory(bNewItem As Boolean)
    LoadItems mnuFindHistory, True
    LoadItems mnuReplaceHistory, False
    'there is a new item... select it
    If bNewItem Then frmMainForm.cboFind.ListIndex = 0
End Sub
'loads the items from a history list
Private Sub LoadItems(mnuMenu As Variant, bFind As Boolean)
On Error Resume Next
    Dim cItems As Collection
    Dim i As Long
    Dim sText As String
    With frmMainForm
        'set to correct history array
        If bFind Then
            Set cItems = .cFindHistory.Items
        Else
            Set cItems = .cReplaceHistory.Items
        End If
        
        If mnuMenu.Count <> 1 Then
            'unload the previous items...
            '(except first item, which is <none> )
            For i = mnuMenu.Count - 1 To 1 Step -1
                Unload mnuMenu(i)
            Next
        End If
        If bFind Then
            'save find text
            sText = .cboFind.Text
            'if we are loading find history, clear the main form
            'combo box too
            .cboFind.Clear
        End If
        'loop through the history items
        For i = 1 To cItems.Count
            'load the menu item
            Load mnuMenu(i)
            'enable it
            mnuMenu(i).Enabled = True
            'add ... if necessary... we only want first 30 chars
            If Len(cItems(i)) > 30 Then
                mnuMenu(i).Caption = Left$(cItems(i), 27) & "..."
            Else
                mnuMenu(i).Caption = Left$(cItems(i), 30)
            End If
            'set the menu tag to full value
            mnuMenu(i).Tag = cItems(i)
            'add the item to the find combo too
            If bFind Then .cboFind.AddItem cItems(i)
        Next
        'show the <none> if there are no items...
        mnuMenu(0).Visible = (mnuMenu.Count = 1)
        If bFind Then .cboFind.Text = sText
    End With
End Sub
'returns the current word in the active box...
Private Function GetCurrentWord() As String
On Error GoTo ErrHandler
    Dim sLine As String
    Dim lCurPos As Long
    Dim lLineStart As Long
    Dim lLastSpace As Long
    Dim lNextSpace As Long
    With ActiveDoc
        lCurPos = .SelStart
        'get the start of the line
        lLineStart = .LineIndex(.CurrentLine)
        'get the cursor pos relative to the line start...
        lCurPos = lCurPos - lLineStart
        'get the current line text
        sLine = .LineText()
    End With
    'find the last & next space
    lLastSpace = InStrRev(Left$(sLine, lCurPos), " ") + 1
    If lLastSpace > lCurPos Then lLastSpace = 1
    'get the next space
    lNextSpace = InStr(1, Right$(sLine, Len(sLine) - lCurPos), " ") + lCurPos
    If lNextSpace = lCurPos Then lNextSpace = Len(sLine) + 1
    'get the text in between
    GetCurrentWord = Mid$(sLine, lLastSpace, (lNextSpace - lLastSpace))
    Exit Function
ErrHandler:
    cDialog.ErrHandler Err, Error, "Find.GetCurrentWord"
End Function
