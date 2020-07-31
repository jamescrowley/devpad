VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDocuments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IDevPadDocuments
Private m_lDocCount As Long
Private m_lIDSeed As Long
Private m_cDocs() As DevPadAddInTlb.IDevPadDocument

Private Property Get IDevPadDocuments_ActiveDoc() As DevPadAddInTlb.IDevPadDocument
    Set IDevPadDocuments_ActiveDoc = ActiveDoc
End Property

Private Property Get IDevPadDocuments_Count() As Long
    'return the number of loaded documents
    IDevPadDocuments_Count = m_lDocCount
End Property

Private Property Get IDevPadDocuments_DocumentIndex(ByVal ID As Long) As Long
    'return the document by it's ID
    Dim i As Long
    For i = 1 To m_lDocCount
        If m_cDocs(i).DocID = ID Then
            IDevPadDocuments_DocumentIndex = i
            Exit Property
        End If
    Next
End Property

Private Property Get IDevPadDocuments_Item(ByVal Item As Long) As DevPadAddInTlb.IDevPadDocument
    Set IDevPadDocuments_Item = m_cDocs(Item)
End Property

Private Property Get IDevPadDocuments_ItemByID(ByVal ID As Long) As DevPadAddInTlb.IDevPadDocument
    'return the document by it's ID
    Dim i As Long
    For i = 1 To m_lDocCount
        If m_cDocs(i).DocID = ID Then
            Set IDevPadDocuments_ItemByID = m_cDocs(i)
            Exit Property
        End If
    Next
End Property

Private Property Get IDevPadDocuments_ItemByPath(ByVal FileName As String) As DevPadAddInTlb.IDevPadDocument
    'return the document by it's path
    Dim i As Long
    For i = 1 To m_lDocCount
        If UCase(FileName) = UCase(m_cDocs(i).FileName) Then
            Set IDevPadDocuments_ItemByPath = m_cDocs(i)
            Exit Property
        End If
    Next
End Property

Private Function IDevPadDocuments_LoadFile(ByVal FileName As String, Optional ByVal ForceText As Boolean = False) As DevPadAddInTlb.IDevPadDocument
    Dim frmForm   As DevPadAddInTlb.IDevPadDocument
    Dim i         As Long
    Dim bExisting As Boolean
    Dim bResult   As Boolean
    On Error GoTo ErrHandler
    Set frmForm = IDevPadDocuments_ItemByPath(FileName)
    If Not frmForm Is Nothing Then
        'activate it
        SendMessage frmMainForm.GetMDIClienthWnd, WM_MDIACTIVATE, frmForm.DocumenthWnd, 0
    Else
        If m_lDocCount = 0 Then
            'load a new document
            Set frmForm = IDevPadDocuments_New(False)
        Else
            'see if it is modified or saved...
            If ActiveDoc.Modified Or ActiveDoc.Saved Then
                'load a new document
                Set frmForm = IDevPadDocuments_New(False)
            Else
                'current document ok to use...
                Set frmForm = ActiveDoc
                'remember that we haven't actually loaded this form...
                bExisting = True
            End If
        End If
        With frmForm
            'we are loading...
            .LoadingFile = True
            'set the status bar
            If Not bClosingMultiple Then SetStatusBar "Opening file...."
            .Show
            On Error Resume Next
            bResult = .LoadFile(FileName, ForceText)
            If bResult = False Or Err <> 0 Then
                'Err.Raise vbObjectError + 101, "DevPad.Documents:LoadFile", Error
                cDialog.ErrHandler Err, Error, "DevPad.Documents:LoadFile", FileName
                'close the form
                If bExisting = False Then frmForm.Close
            End If
            On Error GoTo ErrHandler
            .ClearFlags
        End With
        'reset status bar
        If Not bClosingMultiple Then SetStatusBar
    End If
    'return the form
    Set IDevPadDocuments_LoadFile = frmForm
    Exit Function
ErrHandler:
    Err.Raise vbObjectError + 101, "DevPad.Documents:LoadFile", Error
End Function

Private Property Get IDevPadDocuments_NextIndex(ByVal Index As Long) As Long
    'return the next doc
    If Index >= m_lDocCount Then
        Index = 1
    Else
        Index = Index + 1
    End If
    IDevPadDocuments_NextIndex = Index
End Property

Private Property Get IDevPadDocuments_PreviousIndex(ByVal Index As Long) As Long
    'return the previous doc
    If Index = 1 Then
        Index = m_lDocCount
    Else
        Index = Index - 1
    End If
    IDevPadDocuments_PreviousIndex = Index
End Property

Private Sub IDevPadDocuments_Remove(ByVal Item As Long)
Dim i As Long
Dim lIndex As Long
    lIndex = IDevPadDocuments_DocumentIndex(Item)
    If lIndex <> 0 Then
        Set m_cDocs(lIndex) = Nothing
        'remove the reference to this document...
        'should only be called by the Document's Unload event
        If m_lDocCount > 1 Then
            For i = lIndex To m_lDocCount - 1
                Set m_cDocs(i) = m_cDocs(i + 1)
            Next i
            m_lDocCount = m_lDocCount - 1
            ReDim Preserve m_cDocs(1 To m_lDocCount) As IDevPadDocument
        Else
            m_lDocCount = 0
            Erase m_cDocs
        End If
    End If
End Sub

Private Function IDevPadDocuments_New(Optional ByVal Show As Boolean = True, Optional ByVal Template As String = "", Optional ByVal ForceText As Boolean = False) As DevPadAddInTlb.IDevPadDocument
    On Error GoTo ErrHandler

    'increment document count
    m_lDocCount = m_lDocCount + 1
    
    ReDim Preserve m_cDocs(1 To m_lDocCount)
    'create a new instance of frmDocument

    Set m_cDocs(m_lDocCount) = New frmDocument
    'load it
    Load m_cDocs(m_lDocCount)
    With m_cDocs(m_lDocCount)
        m_lIDSeed = m_lIDSeed + 1
        .DocID = m_lIDSeed
        If Show = False Then .LoadingFile = True
        DoEvents
        'if we want plain-text mode, set the mode...
        If ForceText Then .ChangeMode "text.stx"
        'if we don't want the doc shown yet,
        'set loading tag to true, so usercontrol doesn't
        'waste unnecessary energy!
        'If bShow = False Then .Loading = True
        'set the form's caption
        If Show Then .DocumentCaption = LoadResString(1031) & " " & m_lDocCount 'Document 1 etc...
        'window pos might have changed... get it
        If Not ActiveDoc Is Nothing Then SaveSetting REG_KEY, "Settings", "DocumentState", ActiveDoc.WindowState
        'set the window position
        .WindowState = GetSetting(REG_KEY, "Settings", "DocumentState", vbMaximized)
        'document has not been saved
        .Saved = False
        'set the font to the default one
        .SetFont vDefault.sFont, vDefault.nFontSize
        ' load a template, if specified
        If Template <> "" Then .LoadTemplate Template
        'not modified
        .Modified = False
        'show lines if shown last time
        .ShowLines = GetSetting(REG_KEY, "Settings", "ShowLines", "True")

        If Show Then
            'update the window list, and toolbars
            .Show
            frmMainForm.UpdateWindowList
        End If
    End With
    Set IDevPadDocuments_New = m_cDocs(m_lDocCount)
Exit Function
ErrHandler:
    cDialog.ErrHandler Err, Error, "Documents.New"
End Function

