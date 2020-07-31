VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWorkspace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_sFileName As String
Private m_sFiles()  As String
Private m_lFileCount As Long
Private m_bChanged As Boolean
Public Property Get FileName() As String
    FileName = m_sFileName
End Property
Public Property Let Changed(bNew As Boolean)
    m_bChanged = bNew
End Property
Public Function CloseWorkspace(Optional bNoPrompt As Boolean = False) As Boolean
Dim i       As Long
Dim oForm   As IDevPadDocument
Dim vResult As ShowYesNoResult
    If m_sFileName = "" Then
        'ok
        CloseWorkspace = True
        Exit Function
    End If
    If m_bChanged Then
        'display the prompt
        vResult = cDialog.ShowYesNo("Save changes to the current Workspace?", True)
        Select Case vResult
        Case Cancelled
            Exit Function
        Case No
            CloseWorkspace = True
        Case Yes
            'save the workspace
            CloseWorkspace = Save()
        End Select
    Else
        CloseWorkspace = True
    End If
    If CloseWorkspace Then
        '"Close Workspace files too?"
        If bNoPrompt = False Then
            vResult = cDialog.ShowYesNo(LoadResString(1300), True)
            If vResult = Cancelled Then Exit Function
        Else
            vResult = No
        End If
        'saved...
        If vResult = Yes Then
            'close the workspace files too...
            For i = 1 To m_lFileCount
                Set oForm = cDocuments.ItemByPath(m_sFiles(i))
                If Not oForm Is Nothing Then oForm.Close
            Next
        End If
        'erase file info
        m_sFileName = ""
        Erase m_sFiles()
        m_lFileCount = 0
    End If
End Function
Public Function Save() As Boolean
    If m_sFileName = "" Then
        Save = SaveAs()
        'if cancelled, abort
        If Save = False Then Exit Function
    End If
    'save it...
    pSaveWorkspace m_sFileName
    Save = True
End Function
Public Function SaveAs() As Boolean
Dim sFileName As String
    'get a new filename...
    sFileName = pGetFileName
    If sFileName <> "" Then
        m_sFileName = sFileName
        'save the workspace using the new filename
        SaveAs = Save()
    End If
End Function
Private Function pGetFileName() As String
    If cDialog.ShowOpenSaveDialog(True, "Save Workspace As...", "Workspaces (*.dpw)|*.dpw") Then
        pGetFileName = cDialog.FileName
    End If
End Function
Private Sub pSaveWorkspace(sFile As String)
Dim i           As Long
Dim lIndex      As Long
Dim nFileNum    As Integer
Dim sOut        As String
Dim sFlags      As String
Dim lFlags()    As Long
    On Error Resume Next
    'erase the existing file, if there is one
    If Dir$(sFile) <> "" Then Kill sFile
    Err.Clear
    
    On Error GoTo ErrHandler
    nFileNum = FreeFile
    'open the file...
    Open sFile For Output Lock Read Write As #nFileNum
    sOut = "<DevelopersPadWorkspace>" & vbCrLf & _
           "  <Settings" & vbCrLf & _
           "    Version = ""1.0""" & vbCrLf & _
           "    Project = """ & frmProject.ProjectPath & """" & vbCrLf & _
           "  />" & vbCrLf & _
           "    <Toolbars>" & vbCrLf & _
           "    </Toolbars>" & vbCrLf & _
           "  </Settings>"
           
    Print #nFileNum, sOut
    Print #nFileNum, "  <Files>"
    sOut = ""
    Erase m_sFiles
    If cDocuments.Count <> 0 Then
        For i = 1 To cDocuments.Count
            If cDocuments.Item(i).Saved = True Then
                'update array
                m_lFileCount = m_lFileCount + 1
                ReDim Preserve m_sFiles(1 To m_lFileCount)
                m_sFiles(m_lFileCount) = cDocuments.Item(i).FileName
                
                sFlags = ""
                If cDocuments.Item(i).FlagCount > 0 Then
                    'get the flags
                    cDocuments.Item(i).GetFlags lFlags()
                    'build up flag string
                    For lIndex = 1 To UBound(lFlags)
                        sFlags = sFlags & CStr(lFlags(lIndex)) & ","
                    Next
                    If Right$(sFlags, 1) = "," Then sFlags = Left$(sFlags, Len(sFlags) - 1)
                End If
                sOut = sOut & "    <File" & vbCrLf & _
                         "      Path = """ & cDocuments.Item(i).FileName & """" & vbCrLf & _
                         "      SelStart = """ & cDocuments.Item(i).SelStart & """" & vbCrLf & _
                         "      SelLength = """ & cDocuments.Item(i).SelLength & """" & vbCrLf & _
                         "      Flags = """ & sFlags & """" & vbCrLf & _
                         "      Active = """ & Abs(cDocuments.Item(i).DocID = ActiveDoc.DocID) & """" & vbCrLf & _
                         "    />" & vbCrLf
            End If
        Next i
    End If
    Print #nFileNum, sOut
    Print #nFileNum, "  </Files>"
    Print #nFileNum, "</DevelopersPadWorkspace>"
    'close
    Close #nFileNum
    'we haven't changed...
    m_bChanged = False
    Exit Sub
ErrHandler:
    Err.Raise Err, "DevPad.WorkSpace:Save", Error
    Close #nFileNum
End Sub
Public Function Load(sFile As String) As Boolean
Dim lParentIndex As Long
Dim lNodeIndex   As Long
Dim lItemCount   As Long
Dim lFlags()     As Long
Dim sFlags()     As String
Dim i            As Long
Dim lItem        As Long
Dim lActiveItem  As Long
Dim cXML         As clsXML
Dim oForm        As IDevPadDocument
Dim sPath        As String
    On Error GoTo ErrHandler
    'try to close the current workspace...
    If sFile <> m_sFileName Then
        If CloseWorkspace() = False Then Exit Function
    End If
    m_sFileName = sFile
    SetStatusBar "Loading Workspace..."
    'parse the XML file...
    Set cXML = New clsXML
    With cXML
        cXML.ParseXMLFile sFile

        'get the files node
        lParentIndex = .Node("Files").Index
        lItemCount = .NodeChildCount(lParentIndex, "File")
        'resize the file array
        Erase m_sFiles
        If lItemCount <> 0 Then
            ReDim m_sFiles(1 To lItemCount)
            m_lFileCount = lItemCount
            
            For i = 1 To lItemCount Step 1
                lNodeIndex = .NodeChildIndex(lParentIndex, "File", i)
                sPath = .NodeAttribute(lNodeIndex, "Path")
                'save the path
                m_sFiles(i) = sPath
                Set oForm = cDocuments.ItemByPath(sPath)
                If oForm Is Nothing Then
                    'document is not open...create a new one
                    Set oForm = cDocuments.New(False)
                    'set loadonactivate info
                    oForm.LoadOnActivate sPath, .NodeAttribute(lNodeIndex, "SelStart"), .NodeAttribute(lNodeIndex, "SelLength")
                    'set the document's caption
                    oForm.DocumentCaption = GetCaption(sPath)
                    If .NodeAttribute(lNodeIndex, "Active") = "1" Then lActiveItem = oForm.DocID
                    'add the flags...
                    sFlags() = Split(.NodeAttribute(lNodeIndex, "Flags"), ",")
                    For lItem = 0 To UBound(sFlags)
                        oForm.AddFlag sFlags(lItem)
                    Next
                    DoEvents
                    'show the form
                    oForm.Show
                End If
            Next i
            If lActiveItem = 0 Then
                'activate it (this ensures LoadOnActivate triggers for last item)
                oForm.SetFocus
            Else
                cDocuments.ItemByID(lActiveItem).SetFocus
            End If
        End If
    End With
    'add to history
    frmMainForm.UpdateFileMenu sFile, 3
    'we haven't changed...
    m_bChanged = False
    'succeeded
    Load = True
    'reset status bar
    SetStatusBar
    Exit Function
ErrHandler:
    Err.Raise Err, "DevPad.WorkSpace:Load", Error
End Function


