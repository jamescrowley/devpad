VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private cHTMLHelp        As clsHTMLHelp
Implements IDevPadDialog


Private Function IDevPadDialog_BrowseForFolder(Optional ByVal InitialDir As String = "") As String
    Dim s As String
    Dim cDlg As clsBrowseForFolder
  '  If InitialDir = "" Then InitialDir = CurDir$
    Set cDlg = New clsBrowseForFolder
    With cDlg
        .hWndOwner = frmMainForm.hWnd
        .InitialDir = InitialDir
        .FileSystemOnly = True
        .StatusText = True
        .EditBox = True
        .UseNewUI = True
        IDevPadDialog_BrowseForFolder = .BrowseForFolder
    End With
End Function


Private Sub IDevPadDialog_ErrHandler(ByVal ErrorNum As Long, ByVal ErrorText As String, Optional ByVal Source As String = "", Optional ByVal DebugInfo As String = "", Optional ByVal ParentForm As Object = Nothing)
        Dim sErrMsg As String
    Dim iFileNum As Integer
    
    Load frmMessage
    With frmMessage
        On Error Resume Next
        If Dir(App.Path & "\Errorlog.txt") = Empty Then
            ' create file
            iFileNum = FreeFile
            Open App.Path & "\Errorlog.txt" For Output As iFileNum
            Close #iFileNum
        End If
        ' Write to log
        sErrMsg = "Err " & ErrorNum & ": " & ErrorText & vbCrLf
        If DebugInfo <> "" Then sErrMsg = sErrMsg & DebugInfo & vbCrLf
        sErrMsg = sErrMsg & "Last DLL Error: " & Err.LastDllError & vbCrLf
        sErrMsg = sErrMsg & "Error Source  : " & Err.Source & vbCrLf
        sErrMsg = sErrMsg & "Procedure     : " & Source & vbCrLf
        sErrMsg = sErrMsg & "Build         : " & App.Major & "." & App.Minor & "." & App.Revision
        ' set the details text
        .txtDetails = sErrMsg
        '.SetForm False ' ok, exit buttons
        .InitForm False, False, False
        .picDetails.Visible = True
        
        WriteToFile App.Path & "\errorlog.txt", vbCrLf & sErrMsg, True
        .Caption = "Error"

        .lblMsg.Caption = ErrorText
        If ParentForm Is Nothing Then Set ParentForm = frmMainForm
        .Show vbModal, ParentForm
        Unload frmMessage
    End With
End Sub

Private Property Get IDevPadDialog_FileName() As String
    IDevPadDialog_FileName = CmDlg.FileName
End Property

Private Property Get IDevPadDialog_FileTitle() As String
    IDevPadDialog_FileTitle = CmDlg.FileTitle
End Property

Private Property Get IDevPadDialog_FilterIndex() As Long
    IDevPadDialog_FilterIndex = CmDlg.FilterIndex
End Property

Private Function IDevPadDialog_InputBox(ByVal Prompt As String, ByVal DialogTitle As String, Optional ByVal DefaultValue As String = "", Optional ByVal SelStart As Long = -1&) As String
    Load frmMessage
    With frmMessage
        .txtValue.Text = DefaultValue
        If SelStart = -1 Then
            .txtValue.SelLength = Len(DefaultValue)
        Else
            .txtValue.SelStart = SelStart
        End If
        .lblMsg = Prompt
        .Caption = DialogTitle
        .InitForm False, True, True
        .Show vbModal, frmMainForm
        If .Result <> Cancelled And .txtValue.Text <> DefaultValue Then IDevPadDialog_InputBox = .txtValue
    End With
    Unload frmMessage
End Function

Private Sub IDevPadDialog_ShowHelp(ByVal Contents As Boolean)
    Dim bResult As Boolean
    InitHelp
    cHTMLHelp.hWnd = frmMainForm.hWnd
    If (Contents) Then
        bResult = cHTMLHelp.HTMLHelpContents()
    Else
        bResult = cHTMLHelp.HTMLHelpIndex()
    End If
    If bResult = False Then IDevPadDialog_ErrHandler vbObjectError + 1016, LoadResString(1278), "Main.ShowHelp"
End Sub

Private Sub IDevPadDialog_ShowHelpTopic(ByVal Topic As Long, ByVal hWnd As Long)
    InitHelp
    If hWnd <> 0 Then cHTMLHelp.hWnd = hWnd
    If cHTMLHelp.HTMLShowTopicByID(Topic) = False Then
        cDialog.ErrHandler vbObjectError + 1016, LoadResString(1278), "Main.ShowHelp"
    End If
End Sub

Private Sub IDevPadDialog_ShowMe(DisplayForm As Object, ByVal Modal As Boolean)
    On Error Resume Next
   ' DisplayForm.Show IIf(Modal, vbModal, vbModeless)
End Sub

Private Function IDevPadDialog_ShowOpenSaveDialog(ByVal Save As Boolean, ByVal Title As String, ByVal Filter As String, Optional ByVal FileName As String = "", Optional ByVal Owner As Object = Nothing, Optional ByVal InitFilterIndex As Long = -1&) As Boolean
    If Owner Is Nothing Then Set Owner = frmMainForm
    Load frmSaveAs
    IDevPadDialog_ShowOpenSaveDialog = frmSaveAs.Init(False, , Filter, Title, FileName, Save, Owner, InitFilterIndex)
End Function




Private Sub InitHelp()
    If cHTMLHelp Is Nothing Then
        Set cHTMLHelp = New clsHTMLHelp
        With cHTMLHelp
            .HelpPath = App.Path & "\devpadhelp.chm"
            .hWnd = frmMainForm.hWnd
            .Window = "MainWindow"
        End With
    End If
End Sub

Private Function WriteToFile(sFile As String, sText As String, Optional bAppend As Boolean = False)
    Dim iFileNum As Integer
    iFileNum = FreeFile
    If bAppend Then
        Open sFile For Append As iFileNum
    Else
        Open sFile For Output As iFileNum
    End If
    Print #iFileNum, sText
    Close #iFileNum
End Function

Private Sub IDevPadDialog_ShowWarning(ByVal Prompt As String, Optional ByVal Source As String = "", Optional ByVal DialogTitle As String = "Warning", Optional ByVal ParentForm As Object = Nothing)
    Load frmMessage
    With frmMessage
        .InitForm False, False, False
        .picDetails.Visible = True
        .Caption = DialogTitle
        .lblMsg.Caption = Prompt
        .txtDetails = "Source: " & Source
        If ParentForm Is Nothing Then Set ParentForm = frmMainForm
        .Show vbModal, ParentForm
    End With
    Unload frmMessage
End Sub

Private Function IDevPadDialog_ShowYesNo(ByVal Question As String, ByVal ShowCancel As Boolean, Optional ByVal ParentForm As Object = Nothing) As DevPadAddInTlb.ShowYesNoResult
    Dim szErrMsg As String
    Dim intFileNum As Integer
    'Load frmMessage
    With frmMessage
        ' set the details text
        .ResetResult
        .lblMsg = Question
        .Caption = "Developers Pad"
        .InitForm True, ShowCancel, False
        .Show vbModal, frmMainForm
        IDevPadDialog_ShowYesNo = .Result
    End With
End Function
