VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' ============================================
' Original Author: vbAccelerator.com

Option Explicit

' Registry Specific Access Rights
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_ALL_ACCESS = &H3F

' Open/Create Options
Private Const REG_OPTION_NON_VOLATILE = 0&
Private Const REG_OPTION_VOLATILE = &H1

' Key creation/open disposition
Private Const REG_CREATED_NEW_KEY = &H1
Private Const REG_OPENED_EXISTING_KEY = &H2

' masks for the predefined standard access types
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL = &HFFFF

' Define severity codes
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_ACCESS_DENIED = 5
Private Const ERROR_INVALID_DATA = 13&
Private Const ERROR_MORE_DATA = 234 '   dderror
Private Const ERROR_NO_MORE_ITEMS = 259


' Structures Needed For Registry Prototypes
Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type

' Registry Function Prototypes
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" _
  (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
  ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function RegSetValueExStr Lib "advapi32" Alias "RegSetValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
   ByVal dwType As Long, ByVal szData As String, ByVal cbData As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Declare Function RegQueryValueExStr Lib "advapi32" Alias "RegQueryValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
   ByRef lpType As Long, ByVal szData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32" Alias "RegQueryValueExA" _
  (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
   ByRef lpType As Long, szData As Long, ByRef lpcbData As Long) As Long
   
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" _
  (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
   ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
   lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, _
   lpdwDisposition As Long) As Long

Private Declare Function RegEnumValueStr Lib "advapi32.dll" Alias "RegEnumValueA" _
  (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
   lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, _
   ByVal lpData As String, lpcbData As Long) As Long

'  Other declares:
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long

Public Enum ERegistryClassConstants
    HKEY_classes_root = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum

Public Enum ERegistryValueTypes
' Predefined Value Types
    REG_SZ = (1)                           ' Unicode nul terminated sing
    REG_EXPAND_SZ = (2)                    ' Unicode nul terminated sing w/enviornment var
    REG_MULTI_SZ = (7)                     ' Multiple Unicode sings
End Enum

Private m_hClassKey As Long
Private m_sSectionKey As String
Private m_sValueKey As String
Private m_vValue As Variant
Private m_sSetValue As String
Private m_vDefault As Variant
Private m_eValueType As ERegistryValueTypes

Public Property Get Value() As Variant
Dim vValue As Variant
Dim cData As Long, sData As String, ordType As Long, e As Long
Dim hKey As Long

    e = RegOpenKeyEx(m_hClassKey, m_sSectionKey, 0, KEY_QUERY_VALUE, hKey)
    ' ApiRaiseIf e

    e = RegQueryValueExLong(hKey, m_sValueKey, 0&, ordType, 0&, cData)
    If e And e <> ERROR_MORE_DATA Then
        Value = m_vDefault
        Exit Property
    End If
    
    m_eValueType = ordType
    Select Case ordType
    Case REG_SZ, REG_MULTI_SZ '  Same thing to Visual Basic
        sData = String$(cData - 1, 0)
        e = RegQueryValueExStr(hKey, m_sValueKey, 0&, _
                               ordType, sData, cData)
        vValue = sData
        
    Case REG_EXPAND_SZ
        sData = String$(cData - 1, 0)
        e = RegQueryValueExStr(hKey, m_sValueKey, 0&, _
                               ordType, sData, cData)
        vValue = ExpandEnvStr(sData)

        
    End Select
    Value = vValue
    
End Property
Public Property Let Value( _
        ByVal vValue As Variant _
    )
Dim ordType As Long
Dim c As Long
Dim hKey As Long
Dim e As Long
Dim lCreate As Long
Dim tSA As SECURITY_ATTRIBUTES

    ' Open or Create the key
    e = RegCreateKeyEx(m_hClassKey, m_sSectionKey, 0, "", REG_OPTION_NON_VOLATILE, _
                 KEY_ALL_ACCESS, tSA, hKey, lCreate)
    
    If e Then
        cDialog.ErrHandler 26001, "Failed to set registry value Key: " & m_hClassKey & ", Section: " & m_sSectionKey & ", Key: " & m_sValueKey & " to value: " & m_vValue, App.EXEName & ".clsRegistry"
    Else

        Select Case m_eValueType
        Case REG_SZ, REG_EXPAND_SZ
            Dim s As String, iPos As Long
            s = vValue
            ordType = REG_SZ
            '  Assume anything with two non-adjacent percents is expanded sing
            iPos = InStr(s, "%")
            If iPos Then
                If InStr(iPos + 2, s, "%") Then ordType = REG_EXPAND_SZ
            End If
            c = Len(s) + 1
            e = RegSetValueExStr(hKey, m_sValueKey, 0&, ordType, s, c)
            
        '  User should convert to a compatible type before calling
        Case Else
            e = ERROR_INVALID_DATA
            
        End Select
        
        If Not e Then
            m_vValue = vValue
        Else
            Err.Raise vbObjectError + 1048 + 26001, App.EXEName & ".clsRegistry", "Failed to set registry value Key: ' " & m_hClassKey & "' ,Section: ' " & m_sSectionKey & "' ,Key: ' " & m_sValueKey & "'  to value: ' " & m_vValue & "' "
        End If
        
        ' Close the key
        RegCloseKey hKey
    
    End If
    
End Property

Public Sub CreateEXEAssociation( _
        ByVal sExePath As String, _
        ByVal sClassName As String, _
        ByVal sClassDescription As String, _
        ByVal sAssociation As String, _
        Optional ByVal sOpenMenuText As String = "&Open", _
        Optional ByVal bSupportPrint As Boolean = False, _
        Optional ByVal sPrintMenuText As String = "&Print", _
        Optional ByVal bSupportNew As Boolean = False, _
        Optional ByVal sNewMenuText As String = "&New", _
        Optional ByVal bSupportInstall As Boolean = False, _
        Optional ByVal sInstallMenuText As String = "", _
        Optional ByVal lDefaultIconIndex As Long = -1 _
    )
   '  Check if path is wrapped in quotes:
   
   
   
   
   sExePath = Trim$(sExePath)
   If (Left$(sExePath, 1) <> """") Then
      sExePath = """" & sExePath
   End If
   If (Right$(sExePath, 1) <> """") Then
      sExePath = sExePath & """"
   End If

   '  Create the .File to Class association:
   ClassKey = HKEY_classes_root
   SectionKey = "." & sAssociation
   ValueType = REG_SZ
   ValueKey = ""
   
   If Value = sClassName And App.Path = GetSetting(REG_KEY, "Settings", "AppPath", "") Then
       ' Exit Sub ' already exists
   Else
        If GetSetting(REG_KEY, "Settings", "FilesAssociated", 0) = 1 And GetSetting(REG_KEY, "Settings", "PromptToSetExtension", 1) = 1 Then
            Select Case cDialog.ShowYesNo("The Developers Pad Project files (*.dpp) are not currently associated with Developers Pad. Do you want Developers Pad to associate these files?", False)
            Case No
                SaveSetting REG_KEY, "Settings", "SetExtension", 0
                Exit Sub
            Case Yes
                 SaveSetting REG_KEY, "Settings", "SetExtension", 1
            End Select
        End If
    End If
    SaveSetting REG_KEY, "Settings", "AppPath", App.Path
    
   
   Value = sClassName
   
   '  Create the Class shell open command:
   SectionKey = sClassName
   Value = sClassDescription
   
   SectionKey = sClassName & "\shell\open"
   If (sOpenMenuText = "") Then sOpenMenuText = "&Open"
   ValueKey = ""
   Value = sOpenMenuText
   SectionKey = sClassName & "\shell\open\command"
   ValueKey = ""
   Value = sExePath & " ""%1"""
   
'   If (bSupportPrint) Then
'      SectionKey = sClassName & "\shell\print"
'      If (sPrintMenuText = "") Then sPrintMenuText = "&Print"
'      ValueKey = ""
'      Value = sPrintMenuText
'      SectionKey = sClassName & "\shell\print\command"
'      ValueKey = ""
'      Value = sExePath & " /p ""%1"""
'   End If
'
'   If (bSupportInstall) Then
'      If (sInstallMenuText = "") Then
'         sInstallMenuText = "&Install " & sAssociation
'      End If
'      SectionKey = sClassName & "\shell\add"
'      ValueKey = ""
'      Value = sInstallMenuText
'      SectionKey = sClassName & "\shell\add\command"
'      ValueKey = ""
'      Value = sExePath & " /a ""%1"""
'   End If
'
'   If (bSupportNew) Then
'      SectionKey = sClassName & "\shell\new"
'      ValueKey = ""
'      If (sNewMenuText = "") Then sNewMenuText = "&New"
'      Value = sNewMenuText
'      SectionKey = sClassName & "\shell\new\command"
'      ValueKey = ""
'      Value = sExePath & " /n ""%1"""
'   End If
   
   If lDefaultIconIndex > -1 Then
      SectionKey = sClassName & "\DefaultIcon"
      ValueKey = ""
      Value = sExePath & "," & CStr(lDefaultIconIndex)
   End If
    SaveSetting REG_KEY, "Settings", "FilesAssociated", 0
End Sub

Public Property Get ValueType() As ERegistryValueTypes
    ValueType = m_eValueType
End Property
Public Property Let ValueType(ByVal eValueType As ERegistryValueTypes)
    m_eValueType = eValueType
End Property
Public Property Get ClassKey() As ERegistryClassConstants
    ClassKey = m_hClassKey
End Property
Public Property Let ClassKey( _
        ByVal eKey As ERegistryClassConstants _
    )
    m_hClassKey = eKey
End Property
Public Property Get SectionKey() As String
    SectionKey = m_sSectionKey
End Property
Public Property Let SectionKey( _
        ByVal sSectionKey As String _
    )
    m_sSectionKey = sSectionKey
End Property
Public Property Get ValueKey() As String
    ValueKey = m_sValueKey
End Property
Public Property Let ValueKey( _
        ByVal sValueKey As String _
    )
    m_sValueKey = sValueKey
End Property
Public Property Get Default() As Variant
    Default = m_vDefault
End Property
Public Property Let Default( _
        ByVal vDefault As Variant _
    )
    m_vDefault = vDefault
End Property

Private Function ExpandEnvStr(sData As String) As String
    Dim c As Long, s As String
    '  Get the length
    s = "" '  Needed to get around Windows 95 limitation
    c = ExpandEnvironmentStrings(sData, s, c)
    '  Expand the sing
    s = String$(c - 1, 0)
    c = ExpandEnvironmentStrings(sData, s, c)
    ExpandEnvStr = s
End Function

