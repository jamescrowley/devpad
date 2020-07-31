VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHTMLHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
' Visual Basic code for implementing HTML Help 1.1

'*****
' Declare the following two constants
' as PUBLIC
Private Const HH_HELP_CONTEXT = &HF            ' display mapped numeric
Private Const HH_TP_HELP_WM_HELP = &H11       ' text popup help, same as
                                             ' WinHelp HELP_WM_HELP

Private Const HH_DISPLAY_TOPIC = &H0
Private Const HH_HELP_FINDER = &H0            ' WinHelp equivalent
Private Const HH_DISPLAY_TOC = &H1            ' WinHelp equivalent
Private Const HH_DISPLAY_INDEX = &H2         ' WinHelp equivalent
Private Const HH_DISPLAY_SEARCH = &H3        ' not currently implemented
Private Const HH_SET_WIN_TYPE = &H4
Private Const HH_GET_WIN_TYPE = &H5
Private Const HH_GET_WIN_HANDLE = &H6
Private Const HH_ENUM_INFO_TYPE = &H7        ' Get Info type name, call
                                             ' repeatedly to enumerate,
                                             ' -1 at end
Private Const HH_SET_INFO_TYPE = &H8         ' Add Info type to filter.
Private Const HH_SYNC = &H9
Private Const HH_ADD_NAV_UI = &HA             ' not currently implemented
Private Const HH_ADD_BUTTON = &HB             ' not currently implemented
Private Const HH_GETBROWSER_APP = &HC        ' not currently implemented
Private Const HH_KEYWORD_LOOKUP = &HD
Private Const HH_DISPLAY_TEXT_POPUP = &HE    ' display sing resource id
                                             ' or text in a popup window
                                             ' value in dwData
Private Const HH_TP_HELP_CONTEXTMENU = &H10  ' text popup help, same as
                                             ' WinHelp HELP_CONTEXTMENU
Private Const HH_CLOSE_ALL = &H12             ' close all windows opened
                                             ' directly or indirectly by
                                             ' the caller
Private Const HH_ALINK_LOOKUP = &H13         ' ALink version of
                                             ' HH_KEYWORD_LOOKUP
Private Const HH_GET_LAST_ERROR = &H14       ' not currently implemented
Private Const HH_ENUM_CATEGORY = &H15        ' Get category name, call
                                             ' repeatedly to enumerate,
                                             ' -1 at end
Private Const HH_ENUM_CATEGORY_IT = &H16     ' Get category info type
                                             ' members, call repeatedly to
                                             ' enumerate, -1 at end
Private Const HH_RESET_IT_FILTER = &H17      ' Clear the info type filter
                                             ' of all info types.
Private Const HH_SET_INCLUSIVE_FILTER = &H18 ' set inclusive filtering
                                             ' method for untyped topics
                                             ' to be included in display
Private Const HH_SET_EXCLUSIVE_FILTER = &H19  ' set method for untyped
                                             ' topics to be excluded from
                                             ' the display
Private Const HH_SET_GUID = &H1A              ' For Microsoft Installer --
                                             ' dwData is a pointer to the
                                             ' GUID sing
Private Const HH_INTERNAL = &HFF              ' Used internally.

' Button IDs

Private Const IDTB_EXPAND = 200
Private Const IDTB_CONTRACT = 201
Private Const IDTB_STOP = 202
Private Const IDTB_REFRESH = 203
Private Const IDTB_BACK = 204
Private Const IDTB_HOME = 205
Private Const IDTB_SYNC = 206
Private Const IDTB_PRINT = 207
Private Const IDTB_OPTIONS = 208
Private Const IDTB_FORWARD = 209
Private Const IDTB_NOTES = 210                ' not implemented
Private Const IDTB_BROWSE_FWD = 211
Private Const IDTB_BROWSE_BACK = 212
Private Const IDTB_CONTENTS = 213             ' not implemented
Private Const IDTB_INDEX = 214                ' not implemented
Private Const IDTB_SEARCH = 215               ' not implemented
Private Const IDTB_HISTORY = 216              ' not implemented
Private Const IDTB_BOOKMARKS = 217            ' not implemented
Private Const IDTB_JUMP1 = 218
Private Const IDTB_JUMP2 = 219
Private Const IDTB_CUSTOMIZE = 221
Private Const IDTB_ZOOM = 222
Private Const IDTB_TOC_NEXT = 223
Private Const IDTB_TOC_PREV = 224


Private Type tagHHN_NOTIFY
  hdr As Variant
  pszUrl As String                            ' Multi-byte, null-terminated sing
End Type

Private Type tagHH_POPUP
  cbStruct As Integer                         ' sizeof this sucture
  hinst As Variant                            ' instance handle for sing resource
  idString As Variant                         ' sing resource id, or text id if pszFile
                                             ' is specified in HtmlHelp call
  pszText As String                           ' used if idString is zero
  pt As Integer                               ' top center of popup window
  clrForeground As ColorConstants             ' use -1 for default
  clrBackground As ColorConstants             ' use -1 for default
  rcMargins As RECT                           ' amount of space between edges of window and
                                             ' text, -1 for each member to ignore
  pszFont As String                           ' facename, point size, char set, BOLD ITALIC
                                             ' UNDERLINE
End Type

Private Type tagHH_AKLINK
  cbStruct As Integer                         ' sizeof this sucture
  fReserved As Boolean                        ' must be FALSE (really!)
  pszKeywords As String                       ' semi-colon separated keywords
  pszUrl As String                            ' URL to jump to if no keywords found (may be
                                             ' NULL)
  pszMsgText As String                        ' Message text to display in MessageBox if
                                             ' pszUrl
                                             ' is NULL and no keyword match
  pszMsgTitle As String                       ' Message text to display in MessageBox if
                                             ' pszUrl is NULL and no keyword match
  pszWindow As String                         ' Window to display URL in
  fIndexOnFail As Boolean                     ' Displays index if keyword lookup fails.
End Type

Private Enum NavigationTypes
  HHWIN_NAVTYPE_TOC
  HHWIN_NAVTYPE_INDEX
  HHWIN_NAVTYPE_SEARCH
  HHWIN_NAVTYPE_BOOKMARKS
  HHWIN_NAVTYPE_HISTORY ' not implemented
End Enum

Private Enum IT
  IT_INCLUSIVE
  IT_EXCLUSIVE
  IT_HIDDEN
End Enum

Private Type tagHH_ENUM_IT
  cbStruct As Integer                         ' size of this sucture
  iType As Integer                            ' the type of the information type i.e.
                                             ' Inclusive, Exclusive, or Hidden
  pszCatName As String                        ' Set to the name of the Category to
                                             ' enumerate the info types in a category;
                                             ' else NULL
  pszITName As String                         ' volitile pointer to the name of the
                                             ' infotype. Allocated by call. Caller
                                             ' responsible for freeing
  pszITDescription As String                  ' volitile pointer to the description of the
                                             ' infotype.
End Type

Private Type tagHH_ENUM_CAT
  cbStruct As Integer                         ' size of this sucture
  pszCatName As String                        ' volitile pointer to the category name
  pszCatDescription As String                 ' volitile pointer to the category
                                             ' description
End Type

Private Type tagHH_SET_INFOTYPE
  cbStruct As Integer                         ' the size of this sucture
  pszCatName As String                        ' the name of the category, if any, the
                                             ' InfoType is a member of.
  pszInfoTypeName As String                   ' the name of the info type to add to the
                                             ' filter
End Type

Private Enum NavTabs
  HHWIN_NAVTAB_TOP
  HHWIN_NAVTAB_LEFT
  HHWIN_NAVTAB_BOTTOM
End Enum

Private Const HH_MAX_TABS = 19 ' maximum number of tabs
Private Enum Tabs
  HH_TAB_CONTENTS
  HH_TAB_INDEX
  HH_TAB_SEARCH
  HH_TAB_BOOKMARKS
  HH_TAB_HISTORY
End Enum

' HH_DISPLAY_SEARCH Command Related Structures and Constants

Private Const HH_FTS_DEFAULT_PROXIMITY = (-1)

Private Type tagHH_FTS_QUERY
  cbStruct As Integer                         ' Sizeof sucture in bytes.
  fUniCodeStrings As Boolean                  ' TRUE if all sings are unicode.
  pszSearchQuery As String                    ' String containing the search query.
  iProximity As Long                          ' Word proximity.
  fStemmedSearch As Boolean                   ' TRUE for StemmedSearch only.
  fTitleOnly As Boolean                       ' TRUE for Title search only.
  fExecute As Boolean                         ' TRUE to initiate the search.
  pszWindow As String                         ' Window to display in
End Type

' HH_WINTYPE Structure

Private Const SW_MAXIMIZE = 3
Private Const SW_MINIMIZE = 6
Private Const SW_NORMAL = 1
Private Const SW_SHOW = 5

Private Type HH_WINTYPE
  cbStruct As Integer                         ' IN: size of this sucture including all
                                             ' Information Types
  fUniCodeStrings As Boolean                  ' IN/OUT: TRUE if all sings are in UNICODE
  pszType As String                           ' IN/OUT: Name of a type of window
  fsValidMembers As Variant                   ' IN: Bit flag of valid members
                                             ' (HHWIN_PARAM_)
  fsWinProperties As Variant                  ' IN/OUT: Properties/attributes of the window
                                             ' (HHWIN_)
  pszCaption As String                        ' IN/OUT: Window title
  dwStyles As Variant                         ' IN/OUT: Window styles
  dwExStyles As Variant                       ' IN/OUT: Extended Window styles
  rcWindowPos As RECT                         ' IN: Starting position, OUT: current
                                             ' position
  nShowState As Integer                       ' IN: show state (e.g., SW_SHOW)
  hwndHelp As Variant                         ' OUT: window handle
  hwndCaller As Variant                       ' OUT: who called this window
                                             ' The following members are only valid if
                                             ' HHWIN_PROP_TRI_PANE is set
  hwndToolBar As Variant                      ' OUT: toolbar window in tri-pane window
  hwndNavigation As Variant                   ' OUT: navigation window in tri-pane window
  hwndHTML As Variant                         ' OUT: window displaying HTML in tri-pane
                                             ' window
  iNavWidth As Integer                        ' IN/OUT: width of navigation window
  rcHTML As RECT                              ' OUT: HTML window coordinates
  pszToc As String                            ' IN: Location of the table of contents file
  pszIndex As String                           ' IN: Location of the index file
  pszFile As String                           ' IN: Default location of the html file
  pszHome As String                           ' IN/OUT: html file to display when Home
                                             ' button is clicked
  fsToolBarFlags As Variant                   ' IN: flags controling the appearance of the
                                             ' toolbar
  fNotExpanded As Boolean                     ' IN: TRUE/FALSE to contract or expand, OUT:
                                             ' current state
  curNavType As Integer                       ' IN/OUT: UI to display in the navigational
                                             ' pane
  tabpos As Integer                           ' IN/OUT: HHWIN_NAVTAB_TOP, HHWIN_NAVTAB_LEFT,
                                             ' or HHWIN_NAVTAB_BOTTOM
  idNotify As Integer                         ' IN: ID to use for WM_NOTIFY messages
  tabOrder(HH_MAX_TABS + 1) As Byte           ' IN/OUT: tab order: Contents, Index,
                                             ' Search, History, Favorites, Reserved 1-5,
                                             ' Custom tabs
  cHistory As Integer                         ' IN/OUT: number of history items to keep
                                             ' (default is 30)
  pszJump1 As String                          ' Text for HHWIN_BUTTON_JUMP1
  pszJump2 As String                          ' Text for HHWIN_BUTTON_JUMP2
  pszUrlJump1 As String                       ' URL for HHWIN_BUTTON_JUMP1
  pszUrlJump2 As String                       ' URL for HHWIN_BUTTON_JUMP2
  rcMinSize As RECT                           ' Minimum size for window (ignored in version
                                             ' 1 of the Workshop)
  cbInfoTypes As Integer                      ' size of paInfoTypes;
End Type

Private Enum Actions
  HHACT_TAB_CONTENTS
  HHACT_TAB_INDEX
  HHACT_TAB_SEARCH
  HHACT_TAB_HISTORY
  HHACT_TAB_FAVORITES
  HHACT_EXPAND
  HHACT_CONTRACT
  HHACT_BACK
  HHACT_FORWARD
  HHACT_STOP
  HHACT_REFRESH
  HHACT_HOME
  HHACT_SYNC
  HHACT_OPTIONS
  HHACT_PRINT
  HHACT_HIGHLIGHT
  HHACT_CUSTOMIZE
  HHACT_JUMP1
  HHACT_JUMP2
  HHACT_ZOOM
  HHACT_TOC_NEXT
  HHACT_TOC_PREV
  HHACT_NOTES
  HHACT_LAST_ENUM
End Enum

Private Type tagHHNTRACK
  hdr As Variant
  pszCurUrl As String                         ' Multi-byte, null-terminated sing
  idAction As Integer                         ' HHACT_ value
  phhWinType As HH_WINTYPE                    ' Current window type sucture
End Type

Private Type HH_IDPAIR
  dwControlId As Long
  dwTopicId As Long
End Type

Private m_sWindow As String
Private m_lhWnd As Long
Private m_sHelpPath As String
Private Declare Function htmlhelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Any) As Long

Public Property Get Window() As String
    Window = m_sWindow
End Property
Public Property Let Window(sNewWindow As String)
    m_sWindow = sNewWindow
End Property
Public Property Get HelpPath() As String
    HelpPath = m_sHelpPath
End Property
Public Property Let HelpPath(sNewHelpPath As String)
    m_sHelpPath = sNewHelpPath
End Property
Public Property Get hwnd() As Long
    hwnd = m_lhWnd
End Property
Public Property Let hwnd(lNewhWnd As Long)
    m_lhWnd = lNewhWnd
End Property
Public Function HTMLHelpContents() As Boolean

  ' Force the Help window to display
  ' the Contents file (*.hhc) in the left pane
  If htmlhelp(m_lhWnd, m_sHelpPath & ">" & m_sWindow, HH_DISPLAY_TOC, 0&) <> 0 Then
    HTMLHelpContents = True
  End If
End Function

Public Function HTMLHelpIndex() As Boolean
    ' Force the Help window to display the Index file
    ' (*.hhk) in the left pane
    If htmlhelp(m_lhWnd, m_sHelpPath & ">" & m_sWindow, HH_DISPLAY_INDEX, 0&) <> 0 Then
        HTMLHelpIndex = True
    End If
End Function
Public Function HTMLShowTopic(sTopic As String) As Boolean
    ' Force the Help window to load a specific topic.
    ' The Help window will synchronize the
    ' Contents display automatically
    If htmlhelp(m_lhWnd, m_sHelpPath & ">" & m_sWindow, HH_DISPLAY_TOPIC, sTopic) <> 0 Then
        HTMLShowTopic = True
    End If
End Function
Public Function HTMLShowTopicByID(lTopicID As Long) As Boolean
    ' Force the Help window to load a specific topic.
    ' The Help window will synchronize the
    ' Contents display automatically
    If htmlhelp(m_lhWnd, m_sHelpPath & ">" & m_sWindow, HH_HELP_CONTEXT, lTopicID) <> 0 Then
        HTMLShowTopicByID = True
    End If
End Function
