VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProjectVBGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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
'
Implements IDevPadProjectEx
Private cVBProject As clsProjectVB

Private Function IDevPadProjectEx_LoadProject(ByVal FileName As String) As Boolean
    AddGroupItems FileName, frmProject.NewFolder(GetCaption(FileName))
End Function
Private Function AddGroupItems(ByVal sProjectFile As String, ByVal sParent As String) As Boolean
On Error GoTo ErrHandler
    Dim filenum As Integer
    Dim sNextLine As String
    Dim sCurrentFile As String
    Dim lEqualPos As Long
    Dim sPath As String
    Dim sVBProjectText As String
    Dim sTitle As String
    filenum = FreeFile

    Open sProjectFile For Input As filenum
    If cVBProject Is Nothing Then Set cVBProject = New clsProjectVB
    Do While Not EOF(filenum)
        Line Input #filenum, sNextLine
        lEqualPos = InStr(1, sNextLine, "=")
        If lEqualPos > 1 Then
            sTitle = LCase$(Left$(sNextLine, lEqualPos - 1))
            Select Case sTitle
            Case "project", "startupproject"
                cVBProject.ParseFileName sProjectFile, sNextLine, sPath, sCurrentFile
                cVBProject.AddProjectItems sPath, frmProject.NewFolder(GetCaption(sPath), , sParent), (sTitle = "startupproject")
                frmProject.tvProject.ItemExpanded(sParent) = True
            Case Else
                If sVBProjectText = "" Then
                    sVBProjectText = sNextLine
                Else
                    sVBProjectText = sVBProjectText & vbCrLf & sNextLine
                End If
            End Select
        End If
    Loop

    Close filenum
    'frmProject.FolderTag(frmProject.tvProject.ItemIndex(sParent)) = sProjectFile & vbCrLf & sVBProjectText

    AddGroupItems = True
    Exit Function
ErrHandler:
    Close filenum
    cDialog.ErrHandler Err, Error, "Project.AddVBGroupItems"
End Function
Private Property Get IDevPadProjectEx_ProjectType() As DevPadAppTlb.ProjectTypes
    IDevPadProjectEx_ProjectType = tVBGroup
End Property

Private Function IDevPadProjectEx_SaveProject(ByVal FileName As String) As Boolean
    'NOTE: no longer supported
    'if you think support should remain, email james@vbweb.co.uk!
    cDialog.ErrHandler vbObjectError + 2000, LoadResString(1287), "ProjectVBGroup.SaveProject"
'Dim sFile As String
'    Dim sLine As String
'    Dim sProjectGroup As String
'    Dim sVBProjectPath As String
'    Dim sVBGroupPath As String
'    Dim lHandle As Long
'    Dim lFirstLineEnd As Long
'    Dim VBProjects As Collection
'    Dim VBProject As VB_Project_Info
'    Dim frmNewDoc As IDevPadDocument
'    Dim lParent As Long
'    Dim sExtra As String
'    Dim i As Long
'    Set VBProjects = New Collection
'    'output under construction... output to a new doc rather than file
'    Set frmNewDoc = cDocuments.New(, , True) 'frmMainForm.LoadNewDoc(, , True)
'    On Error Resume Next
'    With frmProject.tvProject
'        'display warning
'        cDialog.ErrHandler 0, "This feature is under construction, and its result is outputted to an unsaved document instead. If you find any problems with the code it produces, please contact me", "VBProjects.Save"
'        frmNewDoc.SelText = vbCrLf & FileName
'        sFile = "VBGROUP 5.0" & vbCrLf
'        'loop through all the items in the TreeView...
'        For i = 1 To .Count
'            lHandle = .ItemIndexByPos(i)
'            If .ItemText(lHandle) Like "*.vbp" Then
'                Dim sTag As String
'                sTag = frmProject.FolderTag(lHandle) 'frmProject.Folder(.ItemData(.Selected)).sPath  '.nType
'                'Set VBProject = New clsVBProjectInfo
'                If sTag <> "" Then
'                    lFirstLineEnd = InStr(1, sTag, vbCrLf)
'                    'get the project path (secondchar to end of first line)
'                    sVBProjectPath = Left$(sTag, lFirstLineEnd - 1)
'                    sVBProjectPath = Right$(sVBProjectPath, Len(sVBProjectPath) - 1)
'                    VBProject.ProjectText = Mid$(sTag, lFirstLineEnd + 2, Len(sTag) - lFirstLineEnd)
'                    VBProject.Path = sVBProjectPath
'                    VBProject.lProjectParent = lHandle
'                    'If bIsVBProjectGroup Then
'                        ' get the relative path
'                        sLine = cVBProject.WriteFileName(GetFolder(sVBGroupPath), sVBProjectPath)
'                        ' get the project type (first char)
'                        Select Case Left$(sTag, 1) 'TAG
'                        Case "1" ' startup project
'                            sLine = "StartupProject=" & sLine
'                        Case "0"
'                            sLine = "Project=" & sLine
'                        End Select
'                        sFile = sFile & vbCrLf & sLine
'                    'End If
'                Else
'                    VBProject.ProjectText = ""
'                    VBProject.Path = FileName
'                    VBProject.lProjectParent = lHandle
'                End If
'                'add VBProject to list...
'                VBProjects.Add VBProject
'                VBProject.Path = ""
'                VBProject.lProjectParent = 0
'                VBProject.ProjectText = ""
'            ElseIf .ItemText(lHandle) Like "*.vbg" Then
'
'                'frmNewDoc.SelText = vbcrlf & nodNode.Tag
'                sTag = frmProject.FolderTag(lHandle) 'frmProject.ReadItemTag(lHandle)
'                If sTag = "" Then
'                    sVBGroupPath = "<Not Set>"
'                Else
'                    sVBGroupPath = Left$(sTag, InStr(1, sTag, vbCrLf) - 1)
'                End If
'
'                frmNewDoc.SelText = vbCrLf & "===Output to " & sVBGroupPath & " ==="
'            End If
'        Next
'        frmNewDoc.SelText = vbCrLf & sFile
'        cVBProject.SaveVBProject VBProjects
'
'        SetStatusBar
'    End With
End Function
