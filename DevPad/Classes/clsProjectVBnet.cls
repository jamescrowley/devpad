VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProjectVBnet"
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
Private cXML As clsXML
Private Function IDevPadProjectEx_LoadProject(ByVal FileName As String) As Boolean
    AddProjectItems FileName, "Project"
End Function
Private Function AddProjectItems(FileName As String, sParent As String) As Boolean
Dim lItemCount As Long
Dim lParentIndex As Long
Dim lNodeIndex As Long
Dim sItemFileName As String
Dim sCurrentSection As String
Dim sFileTitle As String
Dim sFilePath As String
Dim i As Long
    On Error GoTo ErrHandler
    
    If cXML Is Nothing Then Set cXML = New clsXML
    cXML.ParseXMLFile FileName
    With cXML
        lParentIndex = .Node("Files").Index
        lParentIndex = .NodeChildIndex(lParentIndex, "Include")
        lItemCount = .NodeChildCount(lParentIndex, "File")
        If lItemCount > 0 Then
            For i = 1 To lItemCount
                lNodeIndex = .NodeChildIndex(lParentIndex, "File", i)
                sItemFileName = .NodeAttribute(lNodeIndex, "RelPath")
                sCurrentSection = sParent
                'see if there is a folder
                If GetFolder(sItemFileName) <> "" Then
                    sFilePath = sItemFileName
                    Do While InStr(1, sFilePath, "\") <> 0
                        sFileTitle = Left$(sFilePath, InStr(1, sFilePath, "\") - 1)
                        sCurrentSection = frmProject.NewFolder(sFileTitle, True, sCurrentSection)
                        sFilePath = Right$(sFilePath, Len(sFilePath) - InStr(1, sFilePath, "\"))
                    Loop
                End If
                sFileTitle = GetCaption(sItemFileName)
                sItemFileName = GetFolder(FileName) & "\" & sItemFileName
                frmProject.AddItem sCurrentSection, sFileTitle, sItemFileName, vbwFile
            Next
        End If
    End With

'
'    Open FileName For Input As nFileNum
'
'    SetStatusBar "Parsing " & GetCaption(FileName) & "..."
'    Do While Not EOF(nFileNum)
'        Line Input #nFileNum, sLine
'        If InStr(1, sLine, "<Files>") <> 0 Then
'            iStatus = 1 'in files section...
'        ElseIf InStr(1, sLine, "<Include>") <> 0 And iStatus = 1 Then
'            iStatus = 2
'        ElseIf iStatus = 2 And InStr(1, sLine, "<File") <> 0 Then
'            Do
'                Line Input #nFileNum, sLine
'            Loop While InStr(1, sLine, "RelPath=") Or EOF(nFileNum)
'            If EOF(nFileNum) = False Then
'                sLine = Trim$(sLine)
'                lEqualPos = InStr(1, sLine, "=")
'                sItemFileName = Trim$(Right$(sLine, Len(sLine) - lEqualPos))
'                sItemFileName = StripQuotes(sItemFileName)
'                sCurrentSection = sParent
'                'see if there is a folder
'                If GetFolder(sItemFileName) <> "" Then
'                    sFilePath = sItemFileName
'                    Do While InStr(1, sFilePath, "\") <> 0  ' = "..\"
'                        ' up another folder
'                        'sitemfilename=
'                        sFileTitle = Left$(sFilePath, InStr(1, sFilePath, "\") - 1)
'                        sCurrentSection = frmProject.NewFolder(sFileTitle, True, sCurrentSection)
'                        sFilePath = Right$(sFilePath, Len(sFilePath) - InStr(1, sFilePath, "\"))
'                        'sFullPath = Left$(sFullPath, InStrRev(sFullPath, "\") - 1)
'                        'sTitle = Right$(sTitle, Len(sTitle) - 3)
'
'                    Loop
'                End If
'                sFileTitle = GetCaption(sItemFileName)
'                sItemFileName = GetFolder(FileName) & "\" & sItemFileName
'
'                lItem = frmProject.AddItem(sCurrentSection, sFileTitle, sItemFileName, vbwFile)
'
'            End If
'        End If
'    Loop
'    Close nFileNum
    SetStatusBar
    'tvProject.
    frmProject.tvProject.ItemExpanded(sParent) = True

    AddProjectItems = True
    Exit Function
ErrHandler:
   ' Close nFileNum
    cDialog.ErrHandler Err, Error, "clsProjectVB.LoadProject"
End Function
Private Property Get IDevPadProjectEx_ProjectType() As DevPadAppTlb.ProjectTypes
    IDevPadProjectEx_ProjectType = tVBnet
End Property

Private Function IDevPadProjectEx_SaveProject(ByVal FileName As String) As Boolean
    'save not supported
    '"Developers Pad does not currently support saving projects in this format"
    cDialog.ErrHandler vbObjectError + 2000, LoadResString(1287), "ProjectVC.SaveProject"
End Function



