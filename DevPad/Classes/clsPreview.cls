VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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
Private m_frmPreview As frmPreview
Private m_sTempFile As String
Private m_lCount As Long

Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Sub PreviewDocument(frmDocument As IDevPadDocument)

    Dim sTempPath As String

    Dim sServerPath As String
    Dim lPos As Long
    Dim sFileName As String
    Dim sLocalPath As String
    Dim vLocalPaths As Variant
    Dim vServers As Variant
    Dim i As Long
    Dim sOldTempFile As String
    Dim sTempFileName As String
    sOldTempFile = m_sTempFile
    m_sTempFile = ""
    On Error Resume Next
    If sOldTempFile <> "" Then Kill sOldTempFile
    On Error GoTo ErrHandler
    If m_frmPreview Is Nothing Then
        Set m_frmPreview = New frmPreview
        Load m_frmPreview
    End If

    On Error GoTo ErrHandler
    If frmDocument.CharacterCount = 0 Then
        m_frmPreview.NavigateTo App.Path & "\_files\empty_error.htm", "Error"
    Else
        sFileName = frmDocument.FileName
        'if the file is one of the server-only files...
        If InStr(1, GetSetting(REG_KEY, "Settings", "ServerFiles", ".asp .cfm .cfml .ihtml .js .jsp .php"), "." & GetExtension(sFileName)) And sFileName <> "" Then
            'split the sings using ;
            vLocalPaths = Split(GetSetting(REG_KEY, "Settings", "ServerLocalPath", "C:\Inetpub\wwwroot"), ";")
            vServers = Split(GetSetting(REG_KEY, "Settings", "Server", "http://localhost"), ";")
            For i = 0 To UBound(vLocalPaths)
                If Trim$(LCase$(Left$(sFileName, Len(vLocalPaths(i))))) = Trim$(LCase$(vLocalPaths(i))) Then
                    'we are using a server-only file, located in the server path
                    'remove server path
                    sServerPath = Trim$(Right$(sFileName, Len(sFileName) - Len(vLocalPaths(i))))
                    'sServerPath = Replace(sFileName, GetSetting(REG_KEY, "Settings", "ServerLocalPath", ""), "")
                    sLocalPath = GetFolder(Trim$(vLocalPaths(i)) & sServerPath)
                    If frmDocument.Modified Then
                        'if it is modified, save a temporary file...
                        sTempFileName = pGetTempFileName(sLocalPath, GetExtension(sFileName))
                        m_sTempFile = sLocalPath & "\" & sTempFileName
                        SetStatusBar "Saving temporary file..."
                        frmDocument.SaveDocument m_sTempFile, True
                        SetStatusBar
                    Else
                        m_sTempFile = "-1"
                        sTempFileName = GetCaption(sFileName)
                    End If
                    sServerPath = GetFolder(vServers(i) & sServerPath)
                    sServerPath = Replace(sServerPath, "\", "/")
                    sServerPath = sServerPath & "/" & sTempFileName
                    
                    m_frmPreview.brwWebBrowser.Offline = False
                    m_frmPreview.NavigateTo sServerPath, frmDocument.DocumentCaption
                    
                    Exit For
                End If
            Next
            If m_sTempFile = "" Then
                m_frmPreview.brwWebBrowser.Offline = False
                m_frmPreview.NavigateTo App.Path & "\_files\serverside_error.htm", "Error"
                m_sTempFile = "-1"
            End If
        End If
        If m_sTempFile = "" Then
            ' get the path of the saved file
            ' we do this so that if it is a HTML file
            ' and has references to other files in the same
            ' directory, they will not be broken when we
            ' preview the temporarily saved file
            sTempPath = GetFolder(sFileName)
            If sTempPath = "" Then sTempPath = CurDir$
            m_sTempFile = sTempPath & "\" & pGetTempFileName(sTempPath, GetExtension(sFileName))
            If m_sTempFile <> "" Then
                ' save the file
                SetStatusBar "Saving temporary file..."
                frmDocument.SaveDocument m_sTempFile, True
                SetStatusBar
            End If
            m_frmPreview.NavigateTo m_sTempFile, frmDocument.DocumentCaption
        End If
    End If
    If m_sTempFile = "-1" Then m_sTempFile = ""
    ' show the preview
    If m_frmPreview.Visible = False Then
        m_frmPreview.Show
    Else
        ActivateWindow m_frmPreview.hWnd
    End If
    
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "clsPreview.PreviewDocument"
End Sub
Private Function pGetTempFileName(ByVal sTempPath As String, ByVal sExtension As String) As String
    Dim lResult     As Long
    Dim sTempFile   As String
    ' otherwise, use the current directory
    If sExtension = "" Then sExtension = "tmp"
    Do
        sTempFile = App.hInstance & m_lCount & "." & sExtension
        m_lCount = m_lCount + 1
    Loop While (Dir$(sTempPath & "\" & sTempFile) <> "")
    pGetTempFileName = sTempFile
'    ' delete the last temporary file we saved
'    On Error Resume Next
'    sTempFile = Space(MAX_PATH)
'    ' get a temporary filename in that directory
'    lResult = GetTempFileName(sTempPath, "XXX", 0, sTempFile)
'    If lResult = 0 Then
'        cDialog.ErrHandler Err, "Unable to save temporary file " & m_sTempFile & ". " & Error, "clsPreview.PreviewDocument"
'    Else
'        sTempFile = StripChar(Chr$(0), Trim$(sTempFile))
'        pGetTempFileName = Right$(sTempFile, Len(sTempFile) - Len(sTempPath) - 1)
'    End If
End Function
Private Sub Class_Terminate()
    ' delete the temporary file
    On Error Resume Next
    If m_sTempFile <> "" Then Kill m_sTempFile
End Sub
