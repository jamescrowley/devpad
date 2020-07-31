VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProjectVC"
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

Friend Function ParseFileName(ByVal sFile As String, ByVal sLine As String, ByRef sFullPath As String, ByRef sTitle As String)
    sTitle = Right$(sLine, Len(sLine) - InStr(1, sLine, "="))
    sTitle = StripChar("""", sTitle)
    sFullPath = GetFolder(sFile) ' folder
    If Left$(sTitle, 2) = ".\" Then
        sTitle = Right$(sTitle, Len(sTitle) - 2)
    Else
        Do While Left$(sTitle, 3) = "..\"
            ' up another folder
            sFullPath = Left$(sFullPath, InStrRev(sFullPath, "\") - 1)
            sTitle = Right$(sTitle, Len(sTitle) - 3)
        Loop
    End If
    sFullPath = sFullPath & "\" & sTitle '& ";" & sExtra
    sTitle = GetCaption(sTitle)
End Function

Private Function IDevPadProjectEx_LoadProject(ByVal FileName As String) As Boolean
    AddProjectItems FileName, "Project"
End Function
Private Function AddProjectItems(FileName As String, sParent As String, Optional bStartupProject As Boolean = False) As Boolean
    On Error GoTo ErrHandler
    Dim nFileNum As Integer
    Dim sLine As String
    Dim sFileTitle As String
    Dim lEqualPos As Long
    Dim sItemFileName As String
    'Dim nodNode As Node
    Dim sText As String
    Dim lPos As Long
    Dim sParentNode As String
    Dim sVBProjectText As String
    Dim lItem As Long
    Dim sCurrentSection As String
    nFileNum = FreeFile
    'sCurrentSection = sParent
    Open FileName For Input As nFileNum

    SetStatusBar "Parsing " & GetCaption(FileName) & "..."
    Do While Not EOF(nFileNum)
        Line Input #nFileNum, sLine
        If Left$(sLine, 1) = "#" Then
            If Left$(sLine, 13) = "# Begin Group" Then
                lEqualPos = InStr(1, sLine, """")
                lPos = InStr(lEqualPos + 1, sLine, """")
                If lEqualPos <> 0 And lPos <> 0 Then
                    sText = Mid$(sLine, lEqualPos + 1, lPos - lEqualPos - 1)
                    sCurrentSection = frmProject.NewFolder(sText, True, sCurrentSection)
                    SetStatusBar "Parsing " & GetCaption(FileName) & " - " & sCurrentSection
                 '   If sText = "Text Files" Then Stop
                End If
            ElseIf Left$(sLine, 11) = "# End Group" Then
                If InStr(1, sCurrentSection, "\") Then
                    'up a level...
                    sCurrentSection = Left$(sCurrentSection, InStrRev(sCurrentSection, "\") - 1)
                Else
                    sCurrentSection = sParent
                End If
            ElseIf Left$(sLine, 11) = "# Begin Target" Then
                'start of files...
                sCurrentSection = sParent
            End If
        ElseIf sCurrentSection <> "" Then
            lEqualPos = InStr(1, sLine, "=")
            lItem = 0
            If lEqualPos > 1 Then
                Select Case LCase$(Left$(sLine, lEqualPos - 1))
                Case "source"
                    ParseFileName FileName, sLine, sItemFileName, sFileTitle
                    'sParent.Selected = True
                    lItem = frmProject.AddItem(sCurrentSection, sFileTitle, sItemFileName, vbwFile)
                    'AddItemTag lItem, "fred" & sText
                End Select
            End If
        End If
    Loop
    Close nFileNum
    SetStatusBar
    'tvProject.
    frmProject.tvProject.ItemExpanded(sParent) = True

    AddProjectItems = True
    Exit Function
ErrHandler:
    Close nFileNum
    cDialog.ErrHandler Err, Error, "clsProjectVB.LoadProject"
End Function
Private Property Get IDevPadProjectEx_ProjectType() As DevPadAppTlb.ProjectTypes
    IDevPadProjectEx_ProjectType = tVC
End Property

Private Function IDevPadProjectEx_SaveProject(ByVal FileName As String) As Boolean
    'save not supported
    '"Developers Pad does not currently support saving projects in this format"
    cDialog.ErrHandler vbObjectError + 2000, LoadResString(1287), "ProjectVC.SaveProject"
End Function


