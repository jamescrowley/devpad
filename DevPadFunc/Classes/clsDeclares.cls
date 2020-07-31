VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDeclares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type NodeProperty
    Name            As String
    Value           As String
End Type
Public Type Node
    Index           As Long
    ParentNode      As Long
    Name            As String
    Value           As String
    HasChildren     As Boolean
    Attributes()    As NodeProperty
    AttributeCount  As Long
End Type
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'Public Type Undo_Item
'    sAddText     As String
'    sDelText     As String
'    lStart       As Long
'    lAddTextLen  As Long
'    lDelTextLen  As Long
'    bAddTextPlain As Boolean
'    bDelTextPlain As Boolean
''    bAddTextRTF  As Boolean
''    bDelTextRTF  As Boolean
'    ModifyType   As ModifyTypes
'End Type
'
'Public Enum ModifyTypes
'    AddText = 0
'    DeleteText = 1
'    ReplaceText = 2
'    CutText = 3
'    PasteText = 4
'    IndentText = 5
'    OutdentText = 6
'    CommentText = 7
'    UnCommentText = 8
'End Enum
Public Type VB_Project_Info
    Path As String
    ProjectText As String
    lProjectParent As Long
End Type
'Public Type Syntax_Info
'    sStrings                   As String
'    sSeps                      As String
'    sInterestList              As String
'    lSecondKeywordStart        As Long
'    sSingleComment             As String
'    sSingleCommentEsc          As String
'    bMultiComment              As Boolean
'    sMultiCommentStart         As String
'    sMultiCommentEnd           As String
'    sFalseQuote                As String
'    bCheckFalseQuote           As Boolean
'    bDelIndent                 As Boolean
'    sKeywords                  As String
'    sProcedures                As String
'    sOperators                 As String
'    bAutoCase                  As Boolean
'    bAutoIndent                As Boolean
'    bTabIndent                 As Boolean
'    sIndent                    As String
'    sHTMLIndent                As String
'    bHTML                      As Boolean
'    bHTMLExtension             As Boolean
'    sHTMLExtensionStart        As String
'    sHTMLExtensionEnd          As String
'    bColourByLine              As Boolean
'    bRTF                       As Boolean
'    bCode                      As Boolean
'    vCaseSensitive             As VbCompareMethod
'    vClr_Comment               As OLE_COLOR
'    vClr_Keyword               As OLE_COLOR
'    vClr_Keyword2              As OLE_COLOR
'    vClr_Text                  As OLE_COLOR
'    vClr_Operator              As OLE_COLOR
'    vClr_HTMLExTag             As OLE_COLOR
'    vClr_HTMLComment           As OLE_COLOR
'    vClr_HTMLScript            As OLE_COLOR
'    vClr_HTMLTag               As OLE_COLOR
'End Type
'Public Type Syntax_Item
'    sFilter As String 'stores filter sing
'    sName As String 'stores name of language
'    sFile As String 'stores name of syntax file
'    bLoaded As Boolean
'    vSyntaxInfo As Syntax_Info
'End Type


Public Enum ShowCommonDialogFlags
    SCD_OPENFLAGS = 4096 'OFN_FILEMUSTEXIST
    SCD_saveFLAGS = 2 + 2048 'OFN_OVERWRITEPROMPT + OFN_PATHMUSTEXIST
End Enum

Public Enum ShowCommonDialogConstants
'    SCD_OPEN = 0
'    scd_save = 1
    SCD_PRINT = 2
    SCD_COLOR = 3
End Enum
'Public Enum ShowYesNoResult
'    Yes = 1
'    YesToAll = 2
'    No = 3
'    NoToAll = 4
'    None = 5
'    Cancel = -1
'End Enum
'Public Enum ModeConstants
'    vbwVB = 0
'    vbwQBasic = 1
'    vbwHTML = 2
'    vbwJava = 3
'    vbwJavaScript = 4
'    vbwVBScript = 5
'    vbwPerl = 6
'    vbwC = 7
'    vbwText = 8
'    vbwrtf = 9
'    vbwNone = 10
'End Enum
Public Type pCMDLG_VALUES
    FileName As String
    FileTitle As String
    FilterIndex As Long
End Type
