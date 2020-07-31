VERSION 5.00
Begin VB.Form frmFindFiles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find in Files"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFindFiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboFolder 
      Height          =   315
      Left            =   1050
      TabIndex        =   15
      Text            =   "Combo1"
      Top             =   1170
      Width           =   3480
   End
   Begin VB.ComboBox cboFilter 
      Height          =   315
      Left            =   1050
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   810
      Width           =   3480
   End
   Begin VB.Frame fraOptions 
      Caption         =   "1047"
      Height          =   915
      Left            =   45
      TabIndex        =   5
      Top             =   1590
      Width           =   3855
      Begin VB.CheckBox chkSubfolders 
         Caption         =   "1295"
         Height          =   252
         Left            =   105
         TabIndex        =   18
         ToolTipText     =   "Case Sensitivity"
         Top             =   225
         Width           =   2895
      End
      Begin VB.CheckBox chkMatchCase 
         Caption         =   "1092"
         Height          =   252
         Left            =   105
         TabIndex        =   6
         ToolTipText     =   "Case Sensitivity"
         Top             =   525
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "1000"
      Height          =   345
      Left            =   4935
      TabIndex        =   2
      Top             =   2715
      Width           =   1200
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "1011"
      Default         =   -1  'True
      Height          =   345
      Left            =   4935
      TabIndex        =   1
      Top             =   60
      Width           =   1200
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "1013"
      Height          =   345
      Left            =   4935
      TabIndex        =   0
      Top             =   480
      Width           =   1200
   End
   Begin DevPad.vbwFlatButton cmdMoreFind 
      Height          =   285
      Left            =   4560
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   75
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   503
   End
   Begin DevPad.vbwFlatButton cmdMoreReplace 
      Height          =   285
      Left            =   4560
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   450
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   503
   End
   Begin VB.ComboBox cboReplace 
      Height          =   315
      Left            =   1050
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   435
      Width           =   3480
   End
   Begin VB.ComboBox cboFind 
      Height          =   315
      Left            =   1050
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   60
      Width           =   3480
   End
   Begin DevPad.vbwFlatButton cmdFolder 
      Height          =   285
      Left            =   4560
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1185
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   503
   End
   Begin DevPad.vbwFlatButton vbwFlatButton2 
      Height          =   285
      Left            =   4560
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   825
      Width           =   150
      _ExtentX        =   265
      _ExtentY        =   503
   End
   Begin VB.Label lblLabel 
      Caption         =   "1181"
      Height          =   255
      Index           =   4
      Left            =   60
      TabIndex        =   13
      Top             =   825
      Width           =   735
   End
   Begin VB.Label lblLabel 
      Caption         =   "1216"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   12
      Top             =   1215
      Width           =   735
   End
   Begin VB.Label lblLabel 
      Caption         =   "1089"
      Height          =   255
      Index           =   3
      Left            =   60
      TabIndex        =   11
      Top             =   105
      Width           =   615
   End
   Begin VB.Label lblLabel 
      Caption         =   "1090"
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   10
      Top             =   480
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   60
      X2              =   6150
      Y1              =   2610
      Y2              =   2610
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      Caption         =   "1240"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   120
      MouseIcon       =   "frmFindFiles.frx":27A2
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2745
      Width           =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   2
      X1              =   60
      X2              =   6135
      Y1              =   2625
      Y2              =   2625
   End
End
Attribute VB_Name = "frmFindFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_cFindHistory          As clsHistory
Private m_cReplaceHistory       As clsHistory
Private m_cPathHistory          As clsHistory
Private m_sFolders()            As String
Private m_lFolderCount          As Long
Private m_sRootPath             As String
Private m_sExtensions           As String
Private m_bCancel               As Boolean
Private Sub cboFind_Change()
    cmdFind.Enabled = (cboFind.Text <> "")
End Sub

Private Sub cmdCancel_Click()
    If cmdCancel.Caption = LoadResString(1002) Then 'cancel
        m_bCancel = True
    Else
        'hide the form
        Hide
    End If
End Sub

Private Sub cmdFind_Click()
    'do the find
    pDoFind False
End Sub

Private Sub cmdReplace_Click()
    'do the find & replace
    pDoFind True
End Sub

Private Sub cmdFolder_Click()
    Dim sPath As String
    sPath = cDialog.BrowseForFolder(cboFolder.Text)
    If sPath <> "" Then cboFolder.Text = sPath
End Sub

Private Sub Form_Load()
    MsgBox "This feature is under construction!", vbCritical
    'load the resource strings
    LoadResStrings Controls
    'initialize the find history...
    Set m_cFindHistory = New clsHistory
    m_cFindHistory.RegSection = "FindInFiles"
    m_cFindHistory.RegKey = "Find"
    'load the items...
    pLoadHistory m_cFindHistory, cboFind
    
    cboFind.Text = ""
    cboReplace.Text = ""
    cboFilter.Text = cGlobalEditor.SupportedFiles
    cboFolder.Text = CurDir$
    cboFind_Change
End Sub
Private Sub pLoadHistory(cHistory As clsHistory, oCombo As ComboBox)
    Dim i As Long
    oCombo.Clear
    For i = 1 To cHistory.Items.Count
        oCombo.AddItem cHistory.Items(i)
    Next
End Sub
Private Sub pDoFind(bReplace As Boolean)
    Dim lpFindFileData  As WIN32_FIND_DATA
    Dim hFile           As Long
    Dim lRet            As Long
    Dim sTemp           As String
    Dim i               As Long
    Dim sExtension      As String
    'disable everything...
    pDisableForm (False)
    'erase any existing data
    Erase m_sFolders
    m_lFolderCount = 0
    m_sRootPath = cboFolder.Text
    'add a trailing \ if needed
    If Right$(m_sRootPath, 1) <> "\" Then m_sRootPath = m_sRootPath & "\"
    m_sExtensions = LCase$(cboFilter.Text) & ";"
    'if it contains *.*, then clear it... we'll search everything
    If InStr(1, m_sExtensions, "*.*;") Then m_sExtensions = ""
    
    SetStatusBar "Indexing Folders..."
    'walk the folders
    pWalkFolders m_sRootPath, (chkSubfolders.Value = 1)
    'create a new result set
    frmFindResults.NewResults cboFind.Text, cboReplace.Text, bReplace, 1
    'loop through each of the folders looking for files
    'using the specified filter
    For i = 1 To m_lFolderCount Step 1
        ' get a file handle
        hFile = FindFirstFile(m_sFolders(i) & "*.*", lpFindFileData)
        If hFile <> -1 Then
            Do
                'we only want files....
                If (lpFindFileData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> vbDirectory Then
                    'get the filename...
                    sTemp = StripTerminator(lpFindFileData.cFileName)
                    'get it's extension
                    sExtension = GetExtension(sTemp)
                    If InStr(1, m_sExtensions, "*." & sExtension & ";") Or m_sExtensions = "" Then
                        'search the file...
                        SetStatusBar "Checking " & m_sFolders(i) & sTemp
                        pSearchFile (m_sFolders(i) & sTemp)
                        DoEvents
                    End If
                End If
                'go to next item
                lRet = FindNextFile(hFile, lpFindFileData)
            Loop Until lRet = ERROR_NO_MORE_FILES Or lRet = 0 Or m_bCancel = True
        End If
        If m_bCancel Then Exit For
    Next
    ' close the file handle
    lRet = FindClose(hFile)
    ' complete resultset
    frmFindResults.Complete
    SetStatusBar "Complete"
    
    'enable everything...
    pDisableForm True
End Sub
Private Sub pWalkFolders(ByVal sStartPath As String, ByVal bSubFolders As Boolean)
    Dim lpFindFileData  As WIN32_FIND_DATA
    Dim hFile           As Long
    Dim lRet            As Long
    Dim sTemp           As String
    
    ' add start path to folder list
    m_lFolderCount = m_lFolderCount + 1
    ReDim Preserve m_sFolders(1 To m_lFolderCount)
    m_sFolders(m_lFolderCount) = sStartPath
    If bSubFolders Then
        ' get a file handle
        hFile = FindFirstFile(sStartPath & "*.*", lpFindFileData)
        If hFile <> -1 Then
            Do
                'we only want directories....
                If (lpFindFileData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = vbDirectory Then
                    'Strip off null chars and format the sing
                    sTemp = StripTerminator(lpFindFileData.cFileName)
                    ' make sure it is not a reference
                    If sTemp <> "." And sTemp <> ".." Then
                        'walk this folder too...
                        pWalkFolders sStartPath & sTemp & "\", True
                        DoEvents
                    End If
                End If
                'go to next item
                lRet = FindNextFile(hFile, lpFindFileData)
            Loop Until lRet = ERROR_NO_MORE_FILES Or lRet = 0 Or m_bCancel
        End If
        ' close the file handle
        lRet = FindClose(hFile)
    End If
End Sub
Private Sub pSearchFile(ByVal sFile As String)
Dim iFileNum    As Integer
Dim sBuf        As String
Dim lLen        As Long
Dim lPos        As Long
Dim sCaption    As String
Dim vSearchMode As VbCompareMethod
On Error GoTo ErrHandler
    'get filenum
    iFileNum = FreeFile
    'open the file
    Open sFile For Binary Access Read Lock Write As #iFileNum
    'read the file into the buffer
    lLen = LOF(iFileNum)
    sBuf = Space$(lLen)
    Get #iFileNum, , sBuf
    'close the file
    Close #iFileNum
    sCaption = Replace(sFile, m_sRootPath, "")
    If chkMatchCase Then
        vSearchMode = vbBinaryCompare
    Else
        vSearchMode = vbTextCompare
    End If
    Do
        lPos = InStr(lPos + 1, sBuf, cboFind.Text, vSearchMode)
        If lPos <> 0 Then
            'item found
            frmFindResults.AddItem lPos - 1, sCaption, 0, -1, sFile
            lPos = lPos + Len(cboFind.Text)
        End If
    Loop Until lPos = 0

    Exit Sub
ErrHandler:
    Debug.Print "File Skipped: " & sFile
End Sub
'enable/disable all the items in the window...
'used when doing a long find/replace action
Private Sub pDisableForm(bEnable As Boolean)
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
    'reset flag
    m_bCancel = False
End Sub
