VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProjectVB"
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
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Implements IDevPadProjectEx

Friend Function ParseFileName(ByVal sFile As String, ByVal sLine As String, ByRef sFullPath As String, ByRef sTitle As String)
    Dim lSemiColonPos As Long
    lSemiColonPos = InStr(1, sLine, ";")
    If lSemiColonPos <> 0 Then
        sTitle = Right$(sLine, Len(sLine) - lSemiColonPos - 1)
    Else
        sTitle = Right$(sLine, Len(sLine) - InStr(1, sLine, "="))
    End If
    sFullPath = GetFolder(sFile) ' folder
    Do While Left$(sTitle, 3) = "..\"
        ' up another folder
        sFullPath = Left$(sFullPath, InStrRev(sFullPath, "\") - 1)
        sTitle = Right$(sTitle, Len(sTitle) - 3)
    Loop
    sFullPath = sFullPath & "\" & sTitle '& ";" & sExtra
    sTitle = GetCaption(sTitle)
End Function
'Friend Function WriteFileName(ByVal sFolder As String, ByVal sFile As String) As String
'    Dim sFullPath As String
'    Dim sFilePath As String
'    Dim sNewPath As String
'    Dim sPath As String
'
'    sFullPath = sFolder
'    sFilePath = GetFolder(sFile)
'    Do
'        ' go through the actual location of the project
'        ' each time the current directory does not match
'        ' the actual file's path, substitute a ..\
'
'        ' For Example
'        ' Correct VB Project line
'        ' ..\..\Downloads\VB\vbexplorer\WinDefs.bas
'        ' Actual project path
'        ' C:\Documents and Settings\James\My Documents\Visual Studio Projects\DevPad_V1\vbwExplorer.vbp
'        ' Actual file path
'        ' C:\Documents and Settings\James\My Documents\Downloads\VB\vbexplorer\WinDefs.ba
'        If Not sFilePath Like sFullPath & "*" Then
'            sNewPath = "..\" & sNewPath
'             sFullPath = Left$(sFullPath, InStrRev(sFullPath, "\") - 1)
'             DoEvents
'        Else
'            ' then, get the remaining half of the path
'            ' ie \Downloads\VB\vbexplorer\, and all the ..\'s
'            sPath = Replace(sFilePath, sFullPath, "")
'            If sPath = "" Then
'
'            Else
'                sPath = Right$(sPath, Len(sPath) - 1) & "\"
'            End If
'            ' finally, remove the first \ from sPath, and add the filename
'            sNewPath = sNewPath & sPath & GetCaption(sFile)
'            Exit Do
'        End If
'    Loop
'    WriteFileName = sNewPath
'End Function

Friend Function AddProjectItems(sFileName As String, sParent As String, Optional bStartupProject As Boolean = False) As Boolean
    On Error GoTo ErrHandler
    Dim nFileNum As Integer
    Dim sLine As String
    Dim sFileTitle As String
    Dim lEqualPos As Long
    Dim sItemFileName As String
    'Dim nodNode As Node
    Dim sText As String
    Dim sParentNode As String
    Dim sVBProjectText As String
    Dim lItem As Long
    nFileNum = FreeFile
    
    Open sFileName For Input As nFileNum

    SetStatusBar "Parsing " & GetCaption(sFileName) & "..."
    Do While Not EOF(nFileNum)
        Line Input #nFileNum, sLine
        lEqualPos = InStr(1, sLine, "=")
        lItem = 0
        If lEqualPos > 1 Then
            Select Case LCase$(Left$(sLine, lEqualPos - 1))
            Case "class", "form", "module", "usercontrol", "propertypage", "designer"
                Select Case LCase$(Left$(sLine, lEqualPos - 1))
                Case "class"
                    sText = "Class Modules"
                Case Else
                    sText = Left$(sLine, lEqualPos - 1) & "s"
                End Select
                sParentNode = frmProject.NewFolder(sText, True, sParent)
                ParseFileName sFileName, sLine, sItemFileName, sFileTitle
                'sParent.Selected = True
                lItem = frmProject.AddItem(sParentNode, sFileTitle, sItemFileName, vbwFile)
                'AddItemTag lItem, "fred" & sText
            End Select
        End If
        If lItem = 0 Then
            If sVBProjectText = "" Then
                sVBProjectText = sLine
            Else
                sVBProjectText = sVBProjectText & vbCrLf & sLine
            End If
        End If
    Loop
    Close nFileNum
    SetStatusBar
    'tvProject.
    frmProject.tvProject.ItemExpanded(sParent) = True
    '*** Don't forget to put this
    'frmProject.FolderTag(frmProject.tvProject.ItemIndex(sParent)) = IIf(bStartUpProject, "1", "0") & sProjectFile & vbCrLf & sVBProjectText
    'frmProject.FolderTag(frmProject.tvProject.ItemIndex(sParent)) = IIf(bStartupProject, "1", "0") & sFileName & vbCrLf & sVBProjectText
    AddProjectItems = True
    Exit Function
ErrHandler:
    Close nFileNum
    cDialog.ErrHandler Err, Error, "clsProjectVB.LoadProject"
End Function
Private Function IDevPadProjectEx_LoadProject(ByVal FileName As String) As Boolean
    AddProjectItems FileName, frmProject.NewFolder(GetCaption(FileName))
End Function

Private Property Get IDevPadProjectEx_ProjectType() As DevPadAppTlb.ProjectTypes
    IDevPadProjectEx_ProjectType = tVB
End Property

Private Function IDevPadProjectEx_SaveProject(ByVal FileName As String) As Boolean
    'NOTE: no longer supported
    'if you think support should remain, email james@vbweb.co.uk!
    cDialog.ErrHandler vbObjectError + 2000, LoadResString(1287), "ProjectVB.SaveProject"
'Dim sFile As String
'
'
'
'    Dim sLine As String
'
'    Dim sProjectGroup As String
'    Dim sVBProjectPath As String
'    Dim sVBGroupPath As String
'
'
'    Dim lHandle As Long
'    Dim lFirstLineEnd As Long
'    Dim VBProjects As Collection
'    Dim VBProject As VB_Project_Info
'    Dim frmNewDoc As IDevPadDocument
'
'    Dim lParent As Long
'
'    Dim sExtra As String
'    Set VBProjects = New Collection
'
'
'    Dim i As Long
'    'output under construction... output to a new doc rather than file
'    'Set frmNewDoc = frmMainForm.LoadNewDoc(, , True)
'    On Error GoTo ErrHandler
'    With frmProject.tvProject
'        'display warning
'        cDialog.ErrHandler 0, "This feature is under construction, and its result is outputted to an unsaved document instead. If you find any problems with the code it produces, please contact me", "VBProjects.Save"
'        frmNewDoc.SelText = vbCrLf & FileName
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
'                Else
'                    VBProject.ProjectText = ""
'                    VBProject.Path = FileName
'                    VBProject.lProjectParent = lHandle
'                End If
'                'add VBProject to list...
'                VBProjects.Add VBProject
'                'we are done...
'                Exit For
'            End If
'        Next
'        SaveVBProject VBProjects
'
'        SetStatusBar
'    End With
'    Exit Function
'ErrHandler:
'    cDialog.ErrHandler Err, Error, "ProjectVB.SaveProject"
End Function
Friend Function SaveVBProject(VBProjects As Collection)
'    Dim frmNewDoc As IDevPadDocument
'    Dim sVBProjectText As String
'    Dim sVBProjectPath As String
'    Dim sFolder As String
'    Dim i As Long
'    Dim lRefPos As Long
'    Dim lLen As Long
'    Dim sHeader As String
'    Dim sFooter As String
'    Dim lProject As Long
'    Dim lHandle As Long
'    Dim sFile As String
'    Dim sExtension As String
'    Dim sCaption As String
'    Dim lItem As Long
'
'    With frmProject.tvProject
'        For i = 1 To VBProjects.Count
'            Set frmNewDoc = cDocuments.New(, , True) '.LoadNewDoc(, , True)
'            sVBProjectText = VBProjects(i).ProjectText
'            sVBProjectPath = VBProjects(i).Path
'            sFolder = GetFolder(sVBProjectPath)
'            sFolder = ShortPath(sFolder)
'            SetStatusBar "Saving " & GetCaption(VBProjects(i).Path) & "..."
'            ' get the part we insert the files into
'            lRefPos = InStrRev(sVBProjectText, vbCrLf & "Object=")
'            If lRefPos = 0 Then
'                ' get the ref pos
'                lRefPos = InStrRev(sVBProjectText, vbCrLf & "Reference=")
'                If lRefPos = 0 Then
'                    ' get the typepos
'                    lRefPos = InStrRev(sVBProjectText, "Type=")
'                End If
'            End If
'            ' get the length until the end of the line
'            lLen = InStr(lRefPos, sVBProjectText, vbCrLf) - 1
'            If lLen <> -1 Then
'                sHeader = Left$(sVBProjectText, lLen)
'                sFooter = Right$(sVBProjectText, Len(sVBProjectText) - Len(sHeader) - 2)
'            End If
'            frmNewDoc.SelText = vbCrLf & "====Output for " & sVBProjectPath & "===="
'            frmNewDoc.SelText = vbCrLf & sHeader
'
'            lProject = .ItemChild(VBProjects(i).lProjectParent)
'            'Do
'                '.ItemChild(
'            lHandle = lProject
'            Do
'                If .ItemKey(lHandle) <> "Project" And .ItemImage(lHandle) <> IndexForKey("FolderClosed") Then
'                    ' Fill next item in file array
'                    sFile = frmProject.ItemFile(lHandle) 'frmProject.FolderTag(lHandle) 'frmProject.ReadItemTag(lHandle)
'
'                    'we have to make sure both paths are in short format
'                    'whilst keeping the filetitle in longformat!
'                    sExtension = LCase$(GetExtension(sFile))
'                    sCaption = GetCaption(sFile)
'                    sFile = ShortPath(GetFolder(sFile)) & "\" & sCaption 'ShortPath(sFile)
'                    sFile = WriteFileName(sFolder, sFile)
'                  '  If Left$(sFile, 3) = "..\" Then Stop
'                    'Debug.Print sFile
'
'                    Select Case sExtension
'                    Case "bas"
'                        sFile = "Module=" & Left$(sCaption, Len(sCaption) - Len(sExtension) - 1) & ";" & sFile
'                    Case "frm"
'                        sFile = "Form=" & sFile
'                    Case "ctl"
'                        sFile = "UserControl=" & sFile
'                    Case "cls"
'                        sFile = "Class=" & Left$(sCaption, Len(sCaption) - Len(sExtension) - 1) & ";" & sFile
'                    Case "dob"
'                        sFile = "UserDocument=" & sFile
'                    Case "dsr"
'                        sFile = "Designer=" & sFile
'                    Case "pag"
'                        sFile = "PropertyPage=" & sFile
'                    End Select
'                    On Error GoTo 0
'                    frmNewDoc.SelText = vbCrLf & sFile
'                End If
'                lItem = lHandle
'                lHandle = .ItemChild(lItem) 'get its child
'                If lHandle = 0 Then
'                    lHandle = .ItemNextSibling(lItem) 'next item along
'                    If lHandle = 0 Then
'                        lHandle = .ItemNextSibling(lProject)
'                        lProject = lHandle
'                    End If
'                End If
'            Loop While lHandle <> 0
'            frmNewDoc.SelText = vbCrLf & sFooter
'        Next
'    End With
End Function
Private Function ShortPath(sLongPath As String) As String
    Dim sBuffer As String * 255
    Dim lRet As Long
    lRet = GetShortPathName(sLongPath, sBuffer, Len(sBuffer))
    ShortPath = Left$(sBuffer, lRet)
End Function
