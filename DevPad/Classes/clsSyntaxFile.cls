VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSyntaxFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_vSyntaxFiles As Collection
Private m_lProjectFilter As Long
Private m_sFileFilter As String
Private m_sSupportedFiles As String
Private m_lFilterCount As Long
Public Property Get ProjectFilter() As Long
    ProjectFilter = m_lProjectFilter
End Property
Public Property Get FileFilter() As String
    FileFilter = m_sFileFilter
End Property
Public Property Get FilterCount() As Long
    FilterCount = m_lFilterCount
End Property

Public Property Get SyntaxFiles() As Collection
    Set SyntaxFiles = m_vSyntaxFiles
End Property

Private Sub ListLanguages()
    If IsLoaded("frmSplash") Then frmSplash.SetStatus "Loading Syntax Files..."
    Dim lLargestLine As Long
    Dim sName As String
    Dim i As Long

    Dim lpFindFileData As WIN32_FIND_DATA, lFileHdl As Long
    Dim sTemp As String, lRet As Long

    On Error Resume Next
    ' get a file handle
    Dim sSrc As String
    'If vFilters Is Nothing Then Set vFilters = New Collection
    'reset global variables controlling filters
    'ClearStack vFilters
    m_sFileFilter = ""
    ' lFilterCount = 0
    
    sSrc = App.Path & "\_syntax\"
    lFileHdl = FindFirstFile(sSrc & "*.stx", lpFindFileData)
    If lFileHdl <> -1 Then
        Do Until lRet = ERROR_NO_MORE_FILES
            DoEvents
            ' if it is a file
            sTemp = StripTerminator(lpFindFileData.cFileName) ', vbProperCase)
            If sTemp <> "." And sTemp <> ".." Then
                GetSyntaxInfo (sTemp)
                'cboCombo.AddItem left$(sTemp, Len(sTemp) - 4)
            End If
            ' based on the file handle iterate through all files and dirs
            lRet = FindNextFile(lFileHdl, lpFindFileData)
            If lRet = 0 Then Exit Do
        Loop
    End If
    ' close the file handle
    lRet = FindClose(lFileHdl)
    'build the filter sing
    'm_sFileFilter = "All Supported Files|" & m_sSupportedFiles & ";" & PROJECT_EXTENSIONS & m_sFileFilter
    m_sFileFilter = "All Files|*.*" & m_sFileFilter
    m_lProjectFilter = m_vSyntaxFiles.Count + 2 'lFilterCount + 1
    m_sFileFilter = m_sFileFilter & "|Project Files|" & PROJECT_EXTENSIONS
    m_sFileFilter = m_sFileFilter & "|All Files (Force Text)|*.*"

    'SendMessage vbwParent.cboLanguage.hwnd, CB_SETDROPPEDWIDTH, (lLargestLine / Screen.TwipsPerPixelX) + 30, 0&
    'ListLanguages = lLargestLine 'CLng(lLargestLine / Screen.TwipsPerPixelX)
    'If IsLoaded("frmSplash") Then frmSplash.SetStatus "Loading..."
End Sub
Public Sub GetSyntaxFiles(ByRef Items() As String)
Dim i As Long
Dim sName As String
    ReDim Items(1 To m_vSyntaxFiles.Count)
    For i = 1 To m_vSyntaxFiles.Count
        Items(i) = m_vSyntaxFiles(i).sName
    Next
End Sub
Public Property Get SyntaxItems() As Collection
    Set SyntaxItems = m_vSyntaxFiles
End Property
    
Private Sub GetSyntaxInfo(ByVal sFile As String)
Dim iFileNum As Integer
    Dim sLine As String
    Dim iSection As Integer
    Dim sKey As String
    Dim sValue As String
    Dim sName As String
    Dim sFilter As String
    Dim sFilterName As String
    Dim cSyntax As Syntax_Item
    
    iFileNum = FreeFile
    'sFile =
    On Error Resume Next
    Open App.Path & "\_syntax\" & sFile For Input As iFileNum
    If Err Then Exit Sub
    Do While Not EOF(iFileNum)
        Line Input #iFileNum, sLine
        'Ignore comments
        If Left$(sLine, 1) <> ";" And sLine <> "" Then
            If Left$(sLine, 1) = "[" Then
                'section
                Select Case LCase$(Mid$(sLine, 2, Len(sLine) - 2))
                Case "main"
                    iSection = 1
                Case Else
                    Exit Do
                End Select
            ElseIf iSection <> 0 Then
                sKey = Trim$(LCase$(Left$(sLine, InStr(1, sLine, "=") - 1)))
                sValue = Trim$(Right$(sLine, Len(sLine) - InStr(1, sLine, "=")))
                'main
                Select Case sKey
                Case "name"
                    sName = sValue
                Case "filter"
                    sFilter = sValue
                    'm_sFileFilter = m_sFileFilter & "|" & sValue
                    '*** Filter parameter must come last
                    Exit Do
                Case "filtername"
                    sFilterName = sValue 'FileFilter = m_sFileFilter & "|" & sValue
                End Select
            End If
        End If
    Loop
    
    'Set cSyntax = New clsSyntax
    cSyntax.sFile = sFile 'filename of syntax file
    cSyntax.sName = sName 'name of language
    cSyntax.sFilter = sFilter
    cSyntax.bLoaded = False
    m_vSyntaxFiles.Add cSyntax, sName
    'add name
  '  If InStr(1, sFilter, "vbp") Then Stop
    m_sFileFilter = m_sFileFilter & "|" & sFilterName & "|" & sFilter
    
    If Left$(sFilter, 1) <> ";" And m_sSupportedFiles <> "" Then sFilter = ";" & sFilter
    m_sSupportedFiles = m_sSupportedFiles & sFilter
    
    'm_sFileFilter = m_sFileFilter & sFilter
    Close iFileNum
End Sub

Private Sub Class_Initialize()
    Set m_vSyntaxFiles = New Collection
    ListLanguages
End Sub
Public Function LoadSyntaxFile(ByVal sFile As String, ByRef vSyntaxInfo As Syntax_Info) As Long
    Dim i As Long
    Dim lSyntaxItem As Long
    For i = 1 To m_vSyntaxFiles.Count
        If m_vSyntaxFiles(i).sFile = sFile Then
            If m_vSyntaxFiles(i).bLoaded = True Then
                vSyntaxInfo = m_vSyntaxFiles(i).SyntaxInfo
                LoadSyntaxFile = i
                Exit Function
            Else
                lSyntaxItem = i
                LoadSyntaxFile = i
                Exit For
            End If
        End If
    Next i
    With vSyntaxInfo
        'Exit Sub
        .vClr_Text = 0
        .vClr_Keyword = 0
        .vClr_Comment = 0
        .vClr_Operator = 0
        .vClr_HTMLTag = 0
        .vClr_HTMLExTag = 0
        .vClr_HTMLScript = 0
        .vClr_HTMLComment = 0
        .lSecondKeywordStart = 0
        .sSeps = ""
        .bColourByLine = True
        .sInterestList = ""
        .sStrings = ""
        .sKeywords = ""
        .sOperators = ""
        .sProcedures = ""
        .bTabIndent = False
        .sIndent = Chr(vbKeyTab) '""
        .sHTMLIndent = "  "
        .bAutoIndent = False
        .sFalseQuote = ""
        .bCheckFalseQuote = False
        .bDelIndent = True
        .sMultiCommentEnd = ""
        .sMultiCommentStart = ""
        .bMultiComment = False
        .sSingleComment = ""
        .bHTMLExtension = False
        .bHTML = False
        .bRTF = False
        .bAutoCase = False
        .vCaseSensitive = 0
        .sHTMLExtensionStart = ""
        .sHTMLExtensionEnd = ""
        
    '    If .AltStringStarts Is Nothing Then Set .AltStringStarts = New Collection
    '    If .AltStringEnds Is Nothing Then Set .AltStringEnds = New Collection
        'ClearStack .AltStringStarts
        'ClearStack .AltStringEnds
    '    Erase .AltStringEnds
        Dim iFileNum As Integer
        Dim sLine As String
        Dim iSection As Integer
        Dim sKey As String
        Dim sValue As String
        iFileNum = FreeFile
    
        On Error Resume Next
        Open App.Path & "\_syntax\" & sFile For Input As iFileNum
        If Err Then Exit Function
        Do While Not EOF(iFileNum)
            Line Input #iFileNum, sLine
            'Ignore comments
            If Left$(sLine, 1) <> ";" And sLine <> "" Then
                If Left$(sLine, 1) = "[" And Right$(sLine, 1) = "]" Then
                    'section
                    Select Case LCase$(Mid$(sLine, 2, Len(sLine) - 2))
                    Case "main"
                        iSection = 1
                    Case "options"
                        iSection = 2
                    Case "keywordoptions"
                        iSection = 3
                    Case "procedurestarts"
                        iSection = 4
'                    Case "procedureends"
'                        iSection = 5
                    Case "keywords"
                        iSection = 6
                    Case "keywords2"
                        iSection = 7
'                    Case "operators"
'                        iSection = 9
                    Case Else
                        iSection = 0
                    End Select
                ElseIf iSection <> 0 Then
                    If iSection <= 3 Then
                        sKey = Trim$(LCase$(Left$(sLine, InStr(1, sLine, "=") - 1)))
                        sValue = Trim$(Right$(sLine, Len(sLine) - InStr(1, sLine, "=")))
                    End If
        
                    Select Case iSection
                    Case 1 'main
                    Case 2 'options
                        Select Case sKey
                        Case "filecode"
                            .bCode = -sValue
                        Case "filetype"
                            If sValue = 3 Then 'html
                                .bHTML = True
                            ElseIf sValue = 2 Then
                                .bRTF = True
                            End If
                        Case "ignorecase" 'casesensitive
                            .vCaseSensitive = IIf(sValue = "1", 1, 0)
                        Case "autocase" 'autocasecorrect
                            .bAutoCase = -sValue
                        
                        Case "syntaxstart"
                            .sHTMLExtensionStart = sValue
                        Case "syntaxend"
                            .sHTMLExtensionEnd = sValue
                        Case "singlecomment" '"singlelinecomment"
                            .sSingleComment = sValue
                        Case "singlecommentesc"
                            .sSingleCommentEsc = sValue
'                        Case "enablemultilinecomment"
'                            .bMultiComment = -sValue
                        Case "commentstart"
                            .sMultiCommentStart = sValue
                        Case "commentend"
                            .sMultiCommentEnd = sValue
                        Case "stringesc"
                            .sFalseQuote = sValue
                        Case "indentauto"
                            .bAutoIndent = -sValue
                        Case "indentdel"
                            .bDelIndent = -sValue
                        Case "indentvalue"
                            If sValue Like "{*}" Then sValue = Space$(Mid$(sValue, 2, Len(sValue) - 2))
                            .sIndent = sValue
                        Case "indentvaluehtml"
                            .sHTMLIndent = sValue
                        Case "indentwithtab"
                            .bTabIndent = -sValue
                        Case "string", "stringalt"
                            '.sStrings = .sStrings & "*" & sValue
                            .sStrings = .sStrings & sValue
                            .sInterestList = .sInterestList & sValue
                        Case "colourbyline"
                            .bColourByLine = -sValue
                        Case "operators"
                            .sOperators = sValue
                            .sSeps = .sSeps & sValue
                            .sInterestList = .sInterestList & sValue
                        End Select
                    Case 3 'keyword options
                        Select Case sKey
                        Case "wordseperators"
                            .sSeps = .sSeps & ConvertSection(sValue, False, .sInterestList)
                        Case "commentcolour"
                            .vClr_Comment = sValue
                        Case "keywordcolour"
                            .vClr_Keyword = sValue
                        Case "keywordcolour2"
                            .vClr_Keyword2 = sValue
                        Case "textcolour"
                            .vClr_Text = sValue
                        Case "htmltagcolour"
                            .vClr_HTMLTag = sValue
                        Case "htmlextagcolour"
                            .vClr_HTMLExTag = sValue
                        Case "htmlcommentcolour"
                            .vClr_HTMLComment = sValue
                        Case "htmlscriptcolour"
                            .vClr_HTMLScript = sValue
                        Case "operatorcolour"
                            .vClr_Operator = sValue
                        End Select
                    Case 4
                        .sProcedures = .sProcedures & "*" & sLine
                    Case 6 'keywords
                        .sKeywords = .sKeywords & "*" & sLine
                    Case 7
                        If .lSecondKeywordStart = 0 Then
                            .lSecondKeywordStart = Len(.sKeywords)
                        End If
                        .sKeywords = .sKeywords & "*" & sLine
    '                Case 7, 8 'altstringstarts
    '                    sLine = Replace(sLine, "\\n", "\n")
    '                    sLine = Replace(sLine, "\n", vbCrLf)
    '                    If iSection = 7 Then
    '                        .AltStringStarts.Add sLine '"~"
    '                    Else
    '                        .AltStringEnds.Add sLine '"~"
    '                    End If
'                    Case 9 'operators
'                        If sLine = "\;" Then sLine = ";"
                        
                        
                    End Select
                End If
            End If
        Loop
        
        Close iFileNum
        'ensure sings are surrounded by *'s!
        '.sSeps = CorrectString(.sSeps)
        .sKeywords = CorrectString(.sKeywords)
        '.sStrings = CorrectString(.sStrings)
        '.sOperators = CorrectString(.sOperators)
        .sProcedures = CorrectString(.sProcedures)
        
        .bMultiComment = (.sMultiCommentEnd <> "" And .sMultiCommentStart <> "")
        .bCheckFalseQuote = (.sFalseQuote <> "")
        .sInterestList = .sInterestList & .sSingleComment & IIf(.bMultiComment, .sMultiCommentStart & .sMultiCommentEnd, "") & IIf(.bCheckFalseQuote, .sFalseQuote, "")
        .bHTMLExtension = (.sHTMLExtensionEnd <> "" And .sHTMLExtensionStart <> "")
        If .bHTMLExtension Then .bHTML = True
        
        If .lSecondKeywordStart = 0 Then .lSecondKeywordStart = Len(.sKeywords)
        If .vClr_Keyword = 0 Then .sKeywords = ""
        
        If Len(.sIndent) < 2 And Len(.sHTMLIndent) < 2 Then .bDelIndent = False
    '    .sCurIndent = .sIndent
        m_vSyntaxFiles(lSyntaxItem).bLoaded = True
    End With

End Function
Private Function ConvertSection(sText As String, bCrLf As Boolean, ByRef sList As String) As String
Dim lPos As Long
Dim i As Long
Dim lNextPos As Long
Dim sVal As String
    If bCrLf Then
        ConvertSection = Replace(sText, vbCrLf, "*")
    Else
        lPos = 1
        For i = 1 To Len(sText)
            lNextPos = InStr(lPos + 1, sText, ",")
            If lNextPos = 0 Then
                lNextPos = Len(sText) + 1 ' Exit For
                i = Len(sText)
            End If
            sVal = Mid$(sText, lPos, (lNextPos - lPos))
            If Left$(sVal, 1) = "\" Then
                If Len(sVal) > 2 Then
                    If Left$(sVal, 2) <> "\\" Then
                        sVal = Chr(Right$(sVal, Len(sVal) - 1))
                    End If
                Else
                    sVal = Chr(Right$(sVal, Len(sVal) - 1))
                End If
            End If
            
            'ConvertSection = ConvertSection & "*" & sVal
            ConvertSection = ConvertSection & sVal
            sList = sList & sVal
           ' Debug.Print Chr(sVal)
            lPos = lNextPos + 1 'InStr(i, sText, ",")
           ' If lPos = Len(sText) + 2 Then Exit For
        Next
    End If
  '  CorrectString (ConvertSection)
    
End Function
Private Function CorrectString(ByVal sText As String) As String
    If Left$(sText, 1) <> "*" Then sText = "*" & sText
    If Right$(sText, 1) <> "*" Then sText = sText & "*"
    CorrectString = sText
End Function
