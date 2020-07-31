VERSION 5.00
Object = "{6109FEA1-4DC7-493D-8651-9BE2E014D0D2}#3.0#0"; "DevPadEditor.ocx"
Begin VB.Form frmDocument 
   Caption         =   "Document"
   ClientHeight    =   3165
   ClientLeft      =   -165
   ClientTop       =   720
   ClientWidth     =   5985
   DrawStyle       =   1  'Dash
   DrawWidth       =   10
   Icon            =   "frmDocument.frx":0000
   MDIChild        =   -1  'True
   ScaleHeight     =   3165
   ScaleWidth      =   5985
   Begin DevPadEditor.Editor txtText 
      Height          =   3105
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   5477
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmDocument"
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
Implements DevPadAddInTlb.IDevPadDocument
Private m_vSaveOption As SaveOptions
Private m_lDocID As Long
Private m_sLoadOnActivate As String
Private m_lLoadSelStart As Long
Private m_lLoadSelLength As Long
Private m_bIgnore As Boolean
Private Sub Form_Activate()
    If WindowState = vbMinimized Or bClosingMultiple Then Exit Sub
   ' cGlobalEditor.SaveSyntaxInfo ("SQL")
    If m_sLoadOnActivate <> "" And m_bIgnore = False Then
        IDevPadDocument_LoadFile m_sLoadOnActivate
        txtText.SelStart = m_lLoadSelStart
        txtText.SelLength = m_lLoadSelLength
        m_sLoadOnActivate = ""
    End If
    On Error Resume Next
    'ensure the textbox has the focus
    txtText.SetFocus
    'update status text
    txtText_SelChange
    'update ShowLines toolbar button state
    'doesn't work... gets the VB Accelerator control is a twist...
    frmMainForm.tbrMain(0).ButtonPressed("ShowLines") = txtText.ShowLines
    frmMainForm.tbrMain(0).ButtonChecked("ShowLines") = txtText.ShowLines
End Sub

Private Sub Form_Load()
    'assign global editor, to retreive syntax information
    txtText.AssignGlobalEditor cGlobalEditor
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'abort any colourizing
    txtText.CancelColour = True
End Sub
Private Sub Form_Resize()
    'resize the usercontrol
    If WindowState = vbMinimized Then Exit Sub
    txtText.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
Private Sub Form_Terminate()
    'update the lists of windows...
    'and enable/disable the toolbars
    frmMainForm.UpdateWindowList
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrHandler
    'reset Cancel flag
    frmMainForm.bCancelClose = False
    'if we are only closing this one, then set status bar
    If bClosingMultiple = False Then SetStatusBar "Closing File...."
    'save the window state
    SaveSetting REG_KEY, "Settings", "DocumentState", WindowState
    Select Case m_vSaveOption
    Case vbwNotSet 'the save option has not been set
        ' Want to save changes?
        If txtText.Modified = True Then
            txtText.CancelColour = True
            Select Case cDialog.ShowYesNo(LoadResString(1159) & Caption & "?", True)
            Case Yes
                'yes, save the file
                'if the file save was cancelled
                'then return - 1
                If IDevPadDocument_Save() = False Then GoTo CancelUnload
            Case Cancelled
                GoTo CancelUnload
            End Select
        End If
    Case vbwSave 'user has already said to save the file...
        'if the file save was cancelled
        'then return - 1
        If IDevPadDocument_Save() = False Then GoTo CancelUnload
    Case vbwDiscard
        'discard changes...
    End Select
    cDocuments.Remove m_lDocID
    ' Hide the document
    Visible = False
    ' reset status bar
    If Not bClosingMultiple Then SetStatusBar
    Exit Sub
ErrHandler:
    'error handling...
    cDialog.ErrHandler Err, Error, "Document.Unload"
    Exit Sub
CancelUnload:
    'abort save...
    Cancel = -1
    frmMainForm.bCancelClose = True
End Sub



Private Sub txtText_Error(ByVal Err As Long, ByVal Source As String, ByVal Error As String)
    cDialog.ErrHandler Err, Error, Source, "EditorError"
End Sub

Private Sub txtText_FlagsChanged()
    frmMainForm.UpdateFlags
End Sub

Private Sub txtText_LoadFile(ByVal sFileName As String)
    'update the form's caption...
    Caption = GetCaption(sFileName)
    '... and save it to the tag
    Tag = Caption
End Sub

Private Sub txtText_SelChange()
    Dim i As Long
    Dim lCurCol As Long
    Dim lCurLine As Long
    Dim lPos As Long
    Dim lLineCount As Long
    Dim bEnabled As Boolean
    Dim sStatusText As String
    On Error GoTo ErrHandler
    With txtText
        'Build up the status bar text...
        'Display selection...
        If .SelLength >= 1 Then sStatusText = "Selected " & .SelLength & " chars  "
        'Cursor position
        sStatusText = sStatusText & "Pos " & .SelStart
        'Line/Col info
        sStatusText = sStatusText & "  Ln " & .CurrentLine & "/" & .LineCount & "  Col " & .CurrentColumn
        SetStatusBar sStatusText, "CharNum"
        'Update menu/toolbar items
        bEnabled = .CanPaste
        EnableTB "Paste", bEnabled
        EnableItem "EditPaste", bEnabled
        ' Can Cut/Copy?
        bEnabled = .CanCopy
        EnableItem "EditCut", bEnabled
        EnableItem "EditCopy", bEnabled
        EnableItem "EditAppend", bEnabled
        EnableTB "Cut", bEnabled
        EnableTB "Copy", bEnabled
        EnableTB "Append", bEnabled
        
        EnableTB "Undo", .CanUndo
        EnableTB "Redo", .CanRedo
        
        EnableTB "LastLine", .CanGoBack, 2
        EnableTB "NextLine", .CanGoForward, 2
        EnableItem "EditLastLine", .CanGoBack
        EnableItem "EditNextLine", .CanGoForward
        
        'Update the undo/redo tooltiptext...
        frmMainForm.tbrMain(0).ButtonToolTip("Redo") = "Redo " & .RedoText
        frmMainForm.tbrMain(0).ButtonToolTip("Undo") = "Undo " & .UndoText
    End With
Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Document.SelChange"
End Sub



Private Sub txtText_ShowPopup(X As Single, Y As Single)
Dim lPos As Long
    With frmMainForm.ctlPopMenu
        'build it... awful!
        BuildPopupMenu 1
        'set selection to current mouse pos
        lPos = txtText.CharFromPos(X / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY)
        If lPos <> 0 And txtText.SelLength = 0 Then txtText.SelStart = lPos
        'set the popup menu items
        .Enabled("EditPopCut") = txtText.CanCopy
        .Enabled("EditPopCopy") = txtText.CanCopy
        .Enabled("EditPopPaste") = txtText.CanPaste
        .Enabled("EditPopAppend") = txtText.CanCopy
        .Enabled("EditPopUndo") = txtText.CanUndo
        .Enabled("EditPopRedo") = txtText.CanRedo
        If txtText.CursorFile(False) = "" Then
            'no linked file... remove the menu item
            .RemoveItem ("EditPopOpenLinked")
        Else
            'there is a linked file... remove the open item
            .RemoveItem ("EditPopOpen")
        End If
        DoEvents
        'show the popup
        ShowPopup "mnuDocPopup", txtText
        'delete popup
        DeletePopupMenu 1
    End With
End Sub
'*** IDevPadDocument Implementation ***
'ugly... but for the sake of Late binding and
'access for add-in's, it's worth it :-)



Private Sub IDevPadDocument_Close()
    'Close the document
    cDocuments.Remove m_lDocID
    Unload Me
End Sub
Private Property Let IDevPadDocument_DocID(ByVal RHS As Long)
    If m_lDocID = 0 Then m_lDocID = RHS
End Property
Private Property Get IDevPadDocument_DocID() As Long
    IDevPadDocument_DocID = m_lDocID
End Property
Private Property Let IDevPadDocument_DocumentCaption(ByVal RHS As String)
    Caption = RHS
End Property
Private Property Get IDevPadDocument_DocumentCaption() As String
    IDevPadDocument_DocumentCaption = Caption
End Property
Private Property Let IDevPadDocument_Saved(ByVal RHS As Boolean)
    txtText.Saved = RHS
End Property
Private Sub IDevPadDocument_Show()
    m_bIgnore = True
    Show
    DoEvents
    m_bIgnore = False
End Sub
Private Property Let IDevPadDocument_WindowState(ByVal RHS As Long)
    WindowState = RHS
End Property
Private Property Get IDevPadDocument_WindowState() As Long
    IDevPadDocument_WindowState = WindowState
End Property
Private Sub IDevPadDocument_ChangeMode(ByVal File As String)
    txtText.ChangeMode File
End Sub
Private Property Get IDevPadDocument_CharacterCount() As Long
    IDevPadDocument_CharacterCount = txtText.CharacterCount
End Property
Private Property Get IDevPadDocument_CurrentLine() As Long
    IDevPadDocument_CurrentLine = txtText.CurrentLine
End Property
Private Function IDevPadDocument_CursorFile(ByVal AbsolutePath As Boolean) As String
    IDevPadDocument_CursorFile = txtText.CursorFile(AbsolutePath)
End Function
Private Property Get IDevPadDocument_FileMode() As String
    IDevPadDocument_FileMode = txtText.Mode
End Property
Private Property Get IDevPadDocument_FileName() As String
    If m_sLoadOnActivate <> "" Then
        IDevPadDocument_FileName = m_sLoadOnActivate
    Else
        IDevPadDocument_FileName = txtText.FileName
    End If
End Property
Private Property Get IDevPadDocument_FileTitle() As String
    IDevPadDocument_FileTitle = GetCaption(txtText.FileName)
End Property
Private Function IDevPadDocument_SaveDocument(ByVal FileName As String, Optional ByVal bIgnoreSave As Boolean = False) As Boolean
    IDevPadDocument_SaveDocument = txtText.SaveFile(FileName, bIgnoreSave)
End Function
Private Function IDevPadDocument_GetIndent(ByVal Line As String) As String
    IDevPadDocument_GetIndent = txtText.GetIndent(Line)
End Function
Private Property Get IDevPadDocument_IndentString() As String
    IDevPadDocument_IndentString = txtText.IndentString
End Property
Private Property Get IDevPadDocument_IsRTF() As Boolean
    IDevPadDocument_IsRTF = txtText.IsRTF
End Property
Private Property Get IDevPadDocument_LineIndex(ByVal Line As Long) As Long
    IDevPadDocument_LineIndex = txtText.LineIndex(Line)
End Property
Private Property Get IDevPadDocument_LineText(Optional ByVal Line As Long = -1&) As String
    IDevPadDocument_LineText = txtText.LineText(Line)
End Property

Private Property Let IDevPadDocument_LoadingFile(ByVal RHS As Boolean)
    txtText.LoadingFile = RHS
End Property
Private Sub IDevPadDocument_LoadTemplate(ByVal FileName As String)
    txtText.LoadTemplate FileName
End Sub
Private Property Let IDevPadDocument_Modified(ByVal RHS As Boolean)
    txtText.Modified = RHS
End Property
Private Property Get IDevPadDocument_Modified() As Boolean
    IDevPadDocument_Modified = txtText.Modified
End Property
Private Function IDevPadDocument_SaveAs() As Boolean
    Load frmSaveAs
    IDevPadDocument_SaveAs = frmSaveAs.Init(True, Me, , "Save " & Caption & " As...")
End Function
Private Property Get IDevPadDocument_Saved() As Boolean
    IDevPadDocument_Saved = IIf(m_sLoadOnActivate = "", txtText.Saved, True)
End Property
Private Property Let IDevPadDocument_SelLength(ByVal RHS As Long)
    txtText.SelLength = RHS
End Property
Private Property Get IDevPadDocument_SelLength() As Long
    IDevPadDocument_SelLength = txtText.SelLength
End Property
Private Property Let IDevPadDocument_SelRTF(ByVal RHS As String)
End Property
Private Property Get IDevPadDocument_SelRTF() As String
End Property

Private Property Let IDevPadDocument_SelStart(ByVal RHS As Long)
    txtText.SelStart = RHS
End Property
Private Property Get IDevPadDocument_SelStart() As Long
    IDevPadDocument_SelStart = txtText.SelStart
End Property
Private Property Let IDevPadDocument_SelText(ByVal RHS As String)
    txtText.SelText = RHS
End Property
Private Property Get IDevPadDocument_SelText() As String
    IDevPadDocument_SelText = txtText.SelText
End Property
Private Sub IDevPadDocument_SetFont(ByVal FontName As String, ByVal FontSize As Integer)
    txtText.SetFont FontName, FontSize
End Sub
Private Property Let IDevPadDocument_ShowLines(ByVal RHS As Boolean)
    txtText.ShowLines = RHS
End Property
Private Property Get IDevPadDocument_ShowLines() As Boolean
    IDevPadDocument_ShowLines = txtText.ShowLines
End Property
Private Property Let IDevPadDocument_ViewMode(RHS As DevPadAddInTlb.DocumentViewModes)
    txtText.ViewMode = RHS
End Property

Private Property Get IDevPadDocument_ViewMode() As DevPadAddInTlb.DocumentViewModes
    IDevPadDocument_ViewMode = txtText.ViewMode
End Property

Private Function IDevPadDocument_Find(ByVal Find As String, Optional ByVal Start As Long = -1&, Optional ByVal EndLimit As Long = -1&, Optional ByVal Options As Long = -1&) As Long
    IDevPadDocument_Find = txtText.Find(Find, Start, EndLimit, Options)
End Function
Private Function IDevPadDocument_Save() As Boolean
    On Error GoTo ErrHandler
    Dim bResult As Boolean
    With txtText
        If .Saved = True Then
            ' we have a file to save to
            IDevPadDocument_Save = .SaveFile(.FileName)
        Else
            ' we need to find a file to save to
            IDevPadDocument_Save = IDevPadDocument_SaveAs() ' get result
        End If
        .Modified = Not IDevPadDocument_Save ' we have saved, so the doc has not changed
                                ' opposite of this function result
        Exit Function
    End With
ErrHandler:
    cDialog.ErrHandler vbObjectError + 1002, "Error Saving File. Error " & Err & ": " & Error, "Document.SaveFile"
End Function
Private Sub IDevPadDocument_SetFocus()
    SetFocus
    txtText_SelChange
    Form_Activate
End Sub
Private Property Get IDevPadDocument_hWnd() As Long
    IDevPadDocument_hWnd = txtText.hWnd
End Property

Private Property Get IDevPadDocument_DocumenthWnd() As Long
    IDevPadDocument_DocumenthWnd = hWnd
End Property
Private Function IDevPadDocument_LineFromChar(ByVal CharPos As Long) As Long
    IDevPadDocument_LineFromChar = txtText.LineFromChar(CharPos)
End Function
Private Property Get IDevPadDocument_Text() As String
    IDevPadDocument_Text = txtText.Text
End Property
Private Property Let IDevPadDocument_Text(ByVal RHS As String)
    txtText.Text = RHS
End Property

Private Property Get IDevPadDocument_RichText() As String
    IDevPadDocument_RichText = txtText.TextRTF
End Property
Private Property Let IDevPadDocument_Redraw(ByVal RHS As Boolean)
    txtText.Redraw = RHS
End Property

Private Property Let IDevPadDocument_SaveOption(RHS As DevPadAddInTlb.SaveOptions)
    m_vSaveOption = RHS
End Property

Private Property Get IDevPadDocument_LineCount() As Long
    IDevPadDocument_LineCount = txtText.LineCount
End Property

Private Sub IDevPadDocument_Append()
    txtText.Append
End Sub

Private Sub IDevPadDocument_ClearFlags()
    txtText.ClearFlags
End Sub

Private Sub IDevPadDocument_CommentBlock()
    txtText.CommentBlock
End Sub

Private Sub IDevPadDocument_Copy()
    txtText.Copy
End Sub

Private Sub IDevPadDocument_Cut()
    txtText.Cut
End Sub


Private Sub IDevPadDocument_Indent()
    txtText.Indent
End Sub

Private Sub IDevPadDocument_InsertTag()
Dim sTag As String
    '"Enter the tag to insert"
    sTag = cDialog.InputBox(LoadResString(1280), "Quick Tag", "<>", 1)
    If sTag <> "" Then txtText.InsertTag sTag
End Sub

Private Function IDevPadDocument_LastFlag() As Boolean
    IDevPadDocument_LastFlag = pDoFlag(txtText.LastFlag())
End Function

Private Function IDevPadDocument_PreviousFlag(Optional ByVal LoopDoc As Boolean = False) As Boolean
    IDevPadDocument_PreviousFlag = pDoFlag(txtText.PreviousFlag(LoopDoc))
End Function

Private Function IDevPadDocument_NextFlag(Optional ByVal LoopDoc As Boolean = False) As Boolean
    IDevPadDocument_NextFlag = pDoFlag(txtText.NextFlag(LoopDoc))
End Function

Private Function IDevPadDocument_FirstFlag() As Boolean
    IDevPadDocument_FirstFlag = pDoFlag(txtText.FirstFlag)
End Function
Private Function pDoFlag(lLine As Long) As Boolean
    If lLine <> 0 Then
        pDoFlag = True
        
        'activate doc
        If ActiveDoc.DocID <> m_lDocID Then
            'erase these vals...
            m_lLoadSelStart = 0
            m_lLoadSelLength = 0
            SetFocus
            Form_Activate
        End If
        'only change selpos if we aren't on that line
        If txtText.CurrentLine <> lLine Then txtText.SelStart = txtText.LineIndex(lLine)
    End If
End Function

Private Property Get IDevPadDocument_FlagCount() As Long
    IDevPadDocument_FlagCount = txtText.FlagCount
End Property

Private Sub IDevPadDocument_GetFlags(Flags() As Long)
    txtText.GetFlags Flags()
End Sub

Private Sub IDevPadDocument_NextLine()
    txtText.Forward
End Sub

Private Sub IDevPadDocument_Outdent()
    txtText.Outdent
End Sub

Private Sub IDevPadDocument_Paste()
    txtText.Paste
End Sub


Private Sub IDevPadDocument_PreviousLine()
    txtText.Back
End Sub

Private Sub IDevPadDocument_Redo()
    txtText.Redo
End Sub

Private Sub IDevPadDocument_RunCommand(ByVal Command As String)
    'just in case we want extensions, without changing
    'IDevPadDocument definition
    Select Case Command
    Case Else
        Err.Raise vbObjectError + 100, "DevPad.Document:RunCommand", "Invalid Command: " & Command
    End Select
End Sub

Private Sub IDevPadDocument_SelectAll()
    txtText.SelectAll
End Sub

Private Sub IDevPadDocument_ToggleFlag()
    txtText.AddFlag txtText.CurrentLine
End Sub

Private Sub IDevPadDocument_UncommentBlock()
    txtText.UncommentBlock
End Sub

Private Sub IDevPadDocument_Undo()
    txtText.Undo
End Sub

Private Sub IDevPadDocument_AddFlag(ByVal Line As Long)
    txtText.AddFlag Line, True
End Sub

Private Function IDevPadDocument_LoadFile(ByVal FileName As String, Optional ByVal ForceText As Boolean = False) As Boolean
Dim bRTF As Boolean
    'is this an rtf file?
    If GetExtension(FileName) = "rtf" Then bRTF = True
    If (bRTF) Then
        'if it is, hide the lines etc...
        txtText.ShowLines = False
        'goto standard view mode
        txtText.ViewMode = ercWordWrap
    End If
    'give way...
    DoEvents
    'load the file
    If txtText.LoadFile(FileName, bRTF, ForceText) Then
        IDevPadDocument_LoadFile = True
        'update file MRU with opened file...
        frmMainForm.UpdateFileMenu FileName, 1
        'update the window lists...
        If Not bClosingMultiple Then frmMainForm.UpdateWindowList
    End If
End Function

Private Sub IDevPadDocument_LoadOnActivate(ByVal FileName As String, ByVal SelStart As Long, ByVal SelLength As Long)
    m_sLoadOnActivate = FileName
    m_lLoadSelStart = SelStart
    m_lLoadSelLength = SelLength
End Sub

Private Sub IDevPadDocument_ChangeCase(ByVal ToUpperCase As Boolean)
    If ToUpperCase Then
        txtText.SelText = UCase$(txtText.SelText)
    Else
        txtText.SelText = LCase$(txtText.SelText)
    End If
End Sub

Private Sub IDevPadDocument_DeleteLine()
Dim lIndex As Long
    With txtText
        .Redraw = False
        lIndex = .LineIndex(.CurrentLine)
        .SetSelection lIndex, lIndex + Len(.LineText) + 2
        .Delete
        .Redraw = True
    End With
End Sub

Private Sub IDevPadDocument_StringEncode(ByVal ToString As Boolean)
    txtText.SelText = txtText.StringEncode(txtText.SelText, ToString)
End Sub
