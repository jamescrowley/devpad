VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGlobalEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*** API Constants ***
Private Const MAX_PATH = 260
Private Const ERROR_NO_MORE_FILES = 18&
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
'*** API Types ***
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime  As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh   As Long
    nFileSizeLow    As Long
    dwReserved0     As Long
    dwReserved1     As Long
    cFileName       As String * MAX_PATH
    cAlternate      As String * 14
End Type

'*** Types ***
Public Type TemplateInfo
    sSyntax As String
    sDescription As String
    lSelLen As Long
    lSelStart As Long
End Type

Public Type Syntax_Info
    sStrings                   As String
    sStringEncoded             As String
    sSeps                      As String
    sInterestList              As String
    lSecondKeywordStart        As Long
    sSingleComment             As String
    sSingleCommentAlt          As String
    sSingleCommentEsc          As String
    bMultiComment              As Boolean
    sMultiCommentStart         As String
    sMultiCommentEnd           As String
    sFalseQuote                As String
    bCheckFalseQuote           As Boolean
    bDelIndent                 As Boolean
    sKeywords                  As String
    sProcedures                As String
    sOperators                 As String
    bAutoCase                  As Boolean
    bAutoIndent                As Boolean
    sAutoIndentChar            As String
    sAutoOutdentChar           As String
    bTabIndent                 As Boolean
    sIndent                    As String
    sHTMLIndent                As String
    bHTML                      As Boolean
    bHTMLExtension             As Boolean
    sHTMLExtensionStart        As String
    sHTMLExtensionEnd          As String
    bColourByLine              As Boolean
    bRTF                       As Boolean
    bCode                      As Boolean
    vCaseSensitive             As VbCompareMethod
    vClr_Comment               As OLE_COLOR
    vClr_Keyword               As OLE_COLOR
    vClr_Keyword2              As OLE_COLOR
    vClr_Text                  As OLE_COLOR
    vClr_Operator              As OLE_COLOR
    vClr_HTMLExTag             As OLE_COLOR
    vClr_HTMLComment           As OLE_COLOR
    vClr_HTMLScript            As OLE_COLOR
    vClr_HTMLTag               As OLE_COLOR
End Type

Public Type Syntax_Item
    sFilter     As String
    sFilterName As String
    sName       As String
    sFile       As String
    bLoaded     As Boolean
    vSyntaxInfo As Syntax_Info
End Type

'*** ENUM ***

Public Enum FindConstants
    rtfMatchCase = 4
    rtfWholeWord = 2
    rtfNoHighlight = 8
End Enum
'*** API ***
'File Operations
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

'*** Member Variables ***

Private m_tSyntaxFiles() As Syntax_Item
Private m_lSyntaxFileCount As Long

Private m_lProjectFilter As Long
Private m_sFileFilter As String
Private m_sSupportedFiles As String
Private m_lFilterCount As Long
Private m_sSyntaxPath As String
Private m_sTemplatePath As String

'*** Public Properties ***
Public Property Get FileFilter() As String
    FileFilter = m_sFileFilter
End Property
Public Property Get SupportedFiles() As String
    SupportedFiles = m_sSupportedFiles
End Property
Public Property Get SyntaxPath() As String
    SyntaxPath = m_sSyntaxPath
End Property
Public Property Let SyntaxPath(sNew As String)
    m_sSyntaxPath = sNew
End Property
Public Property Get TemplatePath() As String
    TemplatePath = m_sTemplatePath
End Property
Public Property Let TemplatePath(sNew As String)
    m_sTemplatePath = sNew
End Property
Public Property Get FilterCount() As Long
    FilterCount = m_lFilterCount
End Property
'Public Property Get SyntaxFiles() As Collection
'    Set SyntaxFiles = m_vSyntaxFiles
'End Property

Public Sub ListLanguages()
    Dim lLargestLine As Long
    Dim sName As String
    Dim i As Long

    Dim lpFindFileData As WIN32_FIND_DATA, lFileHdl As Long
    Dim sTemp As String, lRet As Long

    On Error Resume Next
    ' get a file handle
    lFileHdl = FindFirstFile(m_sSyntaxPath & "*.stx", lpFindFileData)
    If lFileHdl <> -1 Then
        Do Until lRet = ERROR_NO_MORE_FILES
            DoEvents
            ' if it is a file
            sTemp = StripTerminator(lpFindFileData.cFileName)
            If sTemp <> "." And sTemp <> ".." Then
                GetSyntaxInfo (sTemp)
            End If
            ' based on the file handle iterate through all files and dirs
            lRet = FindNextFile(lFileHdl, lpFindFileData)
            If lRet = 0 Then Exit Do
        Loop
    End If
    ' close the file handle
    lRet = FindClose(lFileHdl)
End Sub
Public Function NewLanguage(sName As String) As Long
Dim cSyntax As Syntax_Item
    'fill the basic details...
    cSyntax.sName = sName
    cSyntax.sFile = LCase$(sName) & ".stx"
    cSyntax.vSyntaxInfo.bCode = True
    cSyntax.vSyntaxInfo.sSeps = "{Return}{Space}{Tab}"
    'create a new language...
    m_lSyntaxFileCount = m_lSyntaxFileCount + 1
    ReDim Preserve m_tSyntaxFiles(1 To m_lSyntaxFileCount)
    LSet m_tSyntaxFiles(m_lSyntaxFileCount) = cSyntax
    'save it...
    SaveSyntaxInfo (m_lSyntaxFileCount)
    NewLanguage = m_lSyntaxFileCount
End Function
Public Sub GetSyntaxFiles(ByRef Items() As String)
Dim i As Long
Dim sName As String
    If m_lSyntaxFileCount <> 0 Then
        ReDim Items(1 To m_lSyntaxFileCount)
        For i = 1 To m_lSyntaxFileCount
            Items(i) = m_tSyntaxFiles(i).sName
        Next
    End If
End Sub
Public Property Get SyntaxFilesCount() As Long
    SyntaxFilesCount = m_lSyntaxFileCount
End Property
Public Property Get SyntaxFile(Item As Variant) As Syntax_Item
Dim lIndex As Long
    lIndex = SyntaxFileIndex(Item)
    
    SyntaxFile = m_tSyntaxFiles(lIndex)
End Property
Public Property Let SyntaxFile(Item As Variant, New_Syntax As Syntax_Item)
Dim lIndex As Long
    lIndex = SyntaxFileIndex(Item)
    m_tSyntaxFiles(lIndex) = New_Syntax
End Property

Public Property Get SyntaxFileIndex(Item As Variant) As Long
Dim i As Long
    If IsNumeric(Item) Then
        SyntaxFileIndex = CLng(Item)
    Else
        For i = 1 To m_lSyntaxFileCount
            If m_tSyntaxFiles(i).sName = CStr(Item) Then
                SyntaxFileIndex = i
            End If
        Next
    End If
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
    Open m_sSyntaxPath & sFile For Input As iFileNum
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
                    '*** Filter parameter must come last
                    Exit Do
                Case "version"
                    If sValue <> "1.0" Then
                        'Close iFileNum
                        'Exit Sub
                    End If
                    'm_sFileFilter = m_sFileFilter & "|" & sValue
                    
                Case "filtername"
                    sFilterName = sValue 'FileFilter = m_sFileFilter & "|" & sValue
                End Select
            End If
        End If
    Loop
    Close iFileNum
    With cSyntax
        .sFile = sFile 'filename of syntax file
        .sName = sName 'name of language
        .sFilter = sFilter
        .sFilterName = sFilterName
        .bLoaded = False
    End With
    
    'add to array
    m_lSyntaxFileCount = m_lSyntaxFileCount + 1
    ReDim Preserve m_tSyntaxFiles(1 To m_lSyntaxFileCount)
    LSet m_tSyntaxFiles(m_lSyntaxFileCount) = cSyntax
    
    'add name
  '  If InStr(1, sFilter, "vbp") Then Stop
    If sFilter <> "" And sFilterName <> "" Then
        m_sFileFilter = m_sFileFilter & "|" & sFilterName & "|" & sFilter
        m_lFilterCount = m_lFilterCount + 1
        If Left$(sFilter, 1) <> ";" And m_sSupportedFiles <> "" Then sFilter = ";" & sFilter
        m_sSupportedFiles = m_sSupportedFiles & sFilter
    End If

End Sub
Public Sub LoadSyntaxInfo(Item As Variant)
Dim lIndex As Long
    lIndex = SyntaxFileIndex(Item)
    LoadSyntaxFile m_tSyntaxFiles(lIndex).sFile, lIndex
End Sub
Public Sub SaveSyntaxInfo(Item As Variant)
Dim lIndex      As Long
Dim iFileNum    As Integer
Dim lVal        As Long
Dim sString     As String
Dim i           As Long
Dim sTemp       As String
    'get the syntax file's index
    lIndex = SyntaxFileIndex(Item)
    'get a free file#
    iFileNum = FreeFile
    'open the file...
    On Error Resume Next
    Open m_sSyntaxPath & m_tSyntaxFiles(lIndex).sFile For Output As iFileNum
    If Err Then Exit Sub
    With m_tSyntaxFiles(lIndex)
        'add the header
        Print #iFileNum, ";Developers Pad Syntax File"
        'main section
        Print #iFileNum, "[Main]"
        Print #iFileNum, "Version = 1.0"
        Print #iFileNum, "Name = " & .sName
        Print #iFileNum, "FilterName = " & .sFilterName
        Print #iFileNum, "Filter = " & .sFilter
        Print #iFileNum, ""
        Print #iFileNum, "[Options]"
        With .vSyntaxInfo
            Print #iFileNum, "FileCode = " & Abs(.bCode)
            'output correct filetype...
            If .bHTML And .bHTMLExtension = False Then
                lVal = 1
            ElseIf .bRTF Then
                lVal = 2
            Else
                lVal = 0
            End If
            Print #iFileNum, "FileType = " & CStr(lVal)
            Print #iFileNum, "IgnoreCase = " & Abs(.vCaseSensitive = vbTextCompare)
            Print #iFileNum, "AutoCase = " & Abs(.bAutoCase)
            Print #iFileNum, "IndentAuto = " & Abs(.bAutoIndent)
            Print #iFileNum, "IndentDel = " & Abs(.bDelIndent)
            Print #iFileNum, "IndentValue = " & ParseHiddenChars(.sIndent)
            Print #iFileNum, "IndentValueHTML = " & ParseHiddenChars(.sHTMLIndent)
            Print #iFileNum, "IndentWithTab = " & Abs(.bTabIndent)
            Print #iFileNum, "IndentAfter = " & .sAutoIndentChar
            Print #iFileNum, "OutdentAfter = " & .sAutoOutdentChar
            Print #iFileNum, "ColourByLine = " & Abs(.bColourByLine)
            Print #iFileNum, "Operators = " & .sOperators
            Print #iFileNum, "String = " & .sStrings
            Print #iFileNum, "StringEsc = " & .sFalseQuote
            Print #iFileNum, "StringEncoded = " & .sStringEncoded
            Print #iFileNum, "SingleComment = " & .sSingleComment
            Print #iFileNum, "SingleCommentAlt = " & .sSingleCommentAlt
            Print #iFileNum, "SingleCommentEsc = " & .sSingleCommentEsc
            Print #iFileNum, "CommentStart = " & .sMultiCommentStart
            Print #iFileNum, "CommentEnd = " & .sMultiCommentEnd
            Print #iFileNum, "SyntaxStart = " & .sHTMLExtensionStart
            Print #iFileNum, "SyntaxEnd = " & .sHTMLExtensionEnd
            Print #iFileNum, ""
            Print #iFileNum, "[KeywordOptions]"
            sTemp = .sSeps
            sTemp = Right$(sTemp, Len(sTemp) - Len(.sOperators))
            'sString = .sSeps
            For i = 1 To Len(sTemp)
                If Asc(Mid$(sTemp, i, 1)) <= 32 Or Mid$(sTemp, i, 1) = "," Then
                    sString = sString & "\" & Asc(Mid$(sTemp, i, 1)) & ","
                Else
                    sString = sString & Mid$(sTemp, i, 1) & ","
                End If
            Next
            If Right$(sString, 1) = "," Then sString = Left$(sString, Len(sString) - 1)
            Print #iFileNum, "WordSeperators = " & sString
            Print #iFileNum, "CommentColour = &H" & Hex$(.vClr_Comment)
            Print #iFileNum, "KeywordColour = &H" & Hex$(.vClr_Keyword)
            Print #iFileNum, "KeywordAltColour = &H" & Hex$(.vClr_Keyword2)
            Print #iFileNum, "TextColour = &H" & Hex$(.vClr_Text)
            Print #iFileNum, "OperatorColour = &H" & Hex$(.vClr_Operator)
            Print #iFileNum, "HTMLCommentColour = &H" & Hex$(.vClr_HTMLComment)
            Print #iFileNum, "HTMLTagColour = &H" & Hex$(.vClr_HTMLTag)
            Print #iFileNum, "HTMLExTagColour = &H" & Hex$(.vClr_HTMLExTag)
            Print #iFileNum, "HTMLScriptColour = &H" & Hex$(.vClr_HTMLScript)
            Print #iFileNum, ""
            Print #iFileNum, "[ProcedureStarts]"
            Print #iFileNum, Replace(.sProcedures, "*", vbCrLf)
            Print #iFileNum, ""
            Print #iFileNum, "[Keywords]"
            Print #iFileNum, Replace(Left$(.sKeywords, .lSecondKeywordStart), "*", vbCrLf)
            If .lSecondKeywordStart <> Len(.sKeywords) Then
                Print #iFileNum, ""
                Print #iFileNum, "[KeywordsAlt]"
                Print #iFileNum, Replace(Right$(.sKeywords, Len(.sKeywords) - .lSecondKeywordStart), "*", vbCrLf)
            End If
        End With
    End With
    Close #iFileNum
End Sub

Friend Function LoadSyntaxFile(ByVal sFile As String, ByRef lIndex As Long) As Syntax_Info
    Dim i As Long
    Dim lSyntaxItem As Long
    Dim iFileNum As Integer
    Dim sLine As String
    Dim iSection As Integer
    Dim sKey As String
    Dim sValue As String
    
    For i = 1 To m_lSyntaxFileCount
        If m_tSyntaxFiles(i).sFile = sFile Then
            If m_tSyntaxFiles(i).bLoaded = True Then
                LoadSyntaxFile = m_tSyntaxFiles(i).vSyntaxInfo
                lIndex = i
                Exit Function
            Else
                lSyntaxItem = i
                lIndex = i
                Exit For
            End If
        End If
    Next i
    If lSyntaxItem = 0 Then Exit Function
    With LoadSyntaxFile
        'initialize variables...
'        .vClr_Text = 0
'        .vClr_Keyword = 0
'        .vClr_Comment = 0
'        .vClr_Operator = 0
'        .vClr_HTMLTag = 0
'        .vClr_HTMLExTag = 0
'        .vClr_HTMLScript = 0
'        .vClr_HTMLComment = 0
'        .lSecondKeywordStart = 0
'        .sSeps = ""
        .bColourByLine = True
'        .sInterestList = ""
'        .sStrings = ""
'        .sKeywords = ""
'        .sOperators = ""
'        .sProcedures = ""
'        .bTabIndent = False
        .sIndent = Chr(vbKeyTab) '""
'        .sHTMLIndent = ""
'        .bAutoIndent = False
'        .sFalseQuote = ""
'        .bCheckFalseQuote = False
        .bDelIndent = True
'        .sMultiCommentEnd = ""
'        .sMultiCommentStart = ""
'        .bMultiComment = False
'        .sSingleComment = ""
'        .bHTMLExtension = False
'        .bHTML = False
'        .bRTF = False
'        .bAutoCase = False
'        .vCaseSensitive = 0
'        .sHTMLExtensionStart = ""
'        .sHTMLExtensionEnd = ""
        
        
        iFileNum = FreeFile
    
        On Error Resume Next
        Open m_sSyntaxPath & sFile For Input As iFileNum
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
                    Case "keywordsalt"
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
                            If sValue = 1 Then 'html
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
                            .sSingleComment = LCase$(sValue) 'not case sensitive...
                        Case "singlecommentesc"
                            .sSingleCommentEsc = sValue
                        Case "singlecommentalt"
                            .sSingleCommentAlt = LCase$(sValue)
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
                            .sIndent = GetHiddenChars(sValue)
                        Case "indentvaluehtml"
                            .sHTMLIndent = GetHiddenChars(sValue)
                        Case "indentwithtab"
                            .bTabIndent = -sValue
                        Case "indentafter"
                            .sAutoIndentChar = sValue
                        Case "outdentafter"
                            .sAutoOutdentChar = sValue
                        Case "string"
                            .sStrings = .sStrings & sValue
                            .sInterestList = .sInterestList & sValue
                        Case "stringencoded"
                            .sStringEncoded = sValue
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
                        Case "keywordaltcolour"
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
        'don't allow autocase if we require the case to be the same anyway
        If .vCaseSensitive = vbBinaryCompare Then .bAutoCase = False

        .bMultiComment = (.sMultiCommentEnd <> "" And .sMultiCommentStart <> "")
        .bCheckFalseQuote = (.sFalseQuote <> "")
        .sInterestList = .sInterestList & .sSingleComment & .sSingleCommentAlt & IIf(.bMultiComment, .sMultiCommentStart & .sMultiCommentEnd, "") & IIf(.bCheckFalseQuote, .sFalseQuote, "")
        .bHTMLExtension = (.sHTMLExtensionEnd <> "" And .sHTMLExtensionStart <> "")
        If .bHTMLExtension Then .bHTML = True
        
        If .lSecondKeywordStart = 0 Then .lSecondKeywordStart = Len(.sKeywords)
        If .vClr_Keyword = 0 Then .sKeywords = ""
        
        If Len(.sIndent) < 2 And Len(.sHTMLIndent) < 2 Then .bDelIndent = False
    '    .sCurIndent = .sIndent
        m_tSyntaxFiles(lSyntaxItem).vSyntaxInfo = LoadSyntaxFile
        m_tSyntaxFiles(lSyntaxItem).bLoaded = True
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
Private Function StripTerminator(ByVal sString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(sString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(sString, intZeroPos - 1)
    Else
        StripTerminator = sString
    End If
End Function

Public Function GetTemplateInfo(ByVal sTemplate As String) As TemplateInfo
    Dim nFileNum As Integer
    Dim sLine As String
    Dim sKey As String
    Dim sValue As String
    Dim bFound As Boolean
    Dim bInSection As Boolean
    Dim lPos As Long
    On Error GoTo ErrHandler
    sTemplate = LCase$(sTemplate)
    nFileNum = FreeFile
    Open m_sTemplatePath & "index.ins" For Input As nFileNum
    Do While EOF(nFileNum) = False
        Line Input #nFileNum, sLine
        If Left$(sLine, 1) <> ";" And sLine <> "" Then
            Select Case Left$(sLine, 1)
            Case "["
                bInSection = (LCase$(sLine) = "[" & sTemplate & "]")
            Case Else
                If bInSection Then
                    lPos = InStr(1, sLine, "=")
                    sKey = LCase$(Trim$(Left$(sLine, lPos - 1)))
                    sValue = Trim$(Right$(sLine, Len(sLine) - lPos))
                    Select Case sKey
                    Case "synxfile"
                        GetTemplateInfo.sSyntax = sValue
                        bFound = True
                    Case "description"
                        GetTemplateInfo.sDescription = sValue
                    Case "selstart"
                        GetTemplateInfo.lSelStart = CLng(sValue)
                    Case "sellen"
                        GetTemplateInfo.lSelLen = CLng(sValue)
                    End Select
                End If
            End Select
        End If
    Loop
    If bFound = False Then
        'not found
        GetTemplateInfo.sSyntax = "text.stx"
    End If
ErrHandler:
    If Err Then Err.Raise Err, "DevPadEditor.Editor:GetTemplateInfo", Error
    Close nFileNum
    Exit Function
End Function
Public Function ParseHiddenChars(sString As String) As String
    If sString = Space$(Len(sString)) And sString <> "" Then
        'all spaces
        ParseHiddenChars = "{" & Len(sString) & "}"
    Else
        ParseHiddenChars = sString
    End If
    ParseHiddenChars = Replace(ParseHiddenChars, Chr$(vbKeyTab), "{Tab}")
    ParseHiddenChars = Replace(ParseHiddenChars, vbCrLf, "{Return}")
    ParseHiddenChars = Replace(ParseHiddenChars, " ", "{Space}")
End Function
Public Function GetHiddenChars(sString As String) As String
    GetHiddenChars = sString
    'remove {tab} to tab char
    GetHiddenChars = Replace(GetHiddenChars, "{Tab}", Chr$(vbKeyTab), , , vbTextCompare)
    GetHiddenChars = Replace(GetHiddenChars, "{Return}", vbCrLf, , , vbTextCompare)
    GetHiddenChars = Replace(GetHiddenChars, "{Space}", " ", , , vbTextCompare)
    If GetHiddenChars Like "{*}" Then
        If IsNumeric(Mid$(GetHiddenChars, 2, Len(GetHiddenChars) - 2)) Then
            'change to spaces...
            GetHiddenChars = Space$(Mid$(GetHiddenChars, 2, Len(GetHiddenChars) - 2))
        End If
    End If
End Function

