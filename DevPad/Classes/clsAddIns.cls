VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAddIns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private m_cAddInInfo() As ADDININFO
Private m_lAddInCount  As Long
Private m_cAddInApp    As DevPadAddInTlb.IDevPadApp
Private m_bChanged     As Boolean
Public Property Get Info(lAddInIndex As Long) As ADDININFO
    Info = m_cAddInInfo(lAddInIndex)
End Property
Public Sub UpdateInfo(lAddInIndex As Long, NewInfo As ADDININFO)
    LSet m_cAddInInfo(lAddInIndex) = NewInfo
    m_bChanged = True
End Sub
'Public Property Set Info(lAddInIndex As Long, NewAddInInfo As ADDININFO)
'    m_cAddInInfo(lAddInIndex) = NewAddInInfo
'End Property

'Public Property Let Info(NewAddInInfo As ADDININFO)
'    Info = m_cAddInInfo(lAddInIndex)
'End Property
'Public Property Get Info() As ADDININFO
'    Info = m_cAddInInfo(lAddInIndex)
'End Property
'Public Property Set Info(NewAddInInfo As ADDININFO)
    'LSet m_cAddInInfo(lAddInIndex) = NewAddInInfo
'End Property

Public Property Get InfoCount() As Long
    InfoCount = m_lAddInCount
End Property

Public Function Add(ByVal lAddInIndex As Long) As Boolean
Dim o As DevPadAddInTlb.IDevPadTools
On Error GoTo ErrHandler
    If m_cAddInApp Is Nothing Then Set m_cAddInApp = New clsAddInApp
    With m_cAddInInfo(lAddInIndex)
        If .Loaded Then
            '.oTool.NewInstance
        Else
            'create an instance of the object
            Set .Tool = CreateObject(.ClassName)
            'we have loaded the add-in
            .Loaded = True
            'call connect
            .Tool.Connect m_cAddInApp
        End If
    End With
    Add = True
    Exit Function
ErrHandler:
    cDialog.ErrHandler Err, "Unable to load Add-In: " & Error & " Please ensure that the add-in is installed correctly, and all the files are correctly registered.", "clsAddIns.Add"
    Add = False
    Exit Function
End Function
Public Function RaiseMenuClicks(lIndex As Long) As Boolean
Dim i As Long
    For i = 1 To m_lAddInCount
        If m_cAddInInfo(i).Loaded Then
            If m_cAddInInfo(i).Tool.MenuClick(lIndex) Then
                'menu handled
                RaiseMenuClicks = True
                Exit For
            End If
        End If
    Next
End Function
Public Sub Remove(ByVal lAddInIndex As Long)
    If lAddInIndex > 0 Then
        With m_cAddInInfo(lAddInIndex)
            If Not .Tool Is Nothing Then
                'close object
                .Tool.HideDialog
                .Tool.Disconnect
            End If
            Set .Tool = Nothing
            .Loaded = False
        End With
    End If
End Sub
 
Public Property Get Tool(lIndex As Long) As IDevPadTools
    If lIndex > 0 Then
        Set Tool = m_cAddInInfo(lIndex).Tool
    End If
End Property

Public Sub UnloadAll()
    Dim i As Long
    'loop through all loaded items, and remove them
    For i = 1 To m_lAddInCount
        If m_cAddInInfo(i).Loaded Then Remove i
    Next
End Sub
Public Sub Terminate()
    
End Sub

'*** Add-Ins ***
Public Sub ProcessAddIns()
    Dim iFileNum As Integer
    Dim sLine As String
    Dim sInfo As String
    Dim vFields() As String
    Dim lPos As Long
    On Error Resume Next
    iFileNum = FreeFile
    Open App.Path & "\addins.ini" For Input As iFileNum
    'it would be much easier to use a binary file from this end...
    'but not so easy for allowing install programs to easily
    'add/remove values from this file!
    
    If Err Then
        'cDialog.ErrHandler vbObjectError + 999, "AddIn Information Not Found"
    Else
        On Error GoTo ErrHandler
        m_lAddInCount = 0
        Do While Not EOF(iFileNum)
            Line Input #iFileNum, sLine
            If Left$(sLine, 1) <> ";" And sLine <> "" Then
                If InStr(1, sLine, "=") Then
                    'an add-in
                    m_lAddInCount = m_lAddInCount + 1
                    ReDim Preserve m_cAddInInfo(1 To m_lAddInCount)
                    sInfo = Right$(sLine, Len(sLine) - InStr(1, sLine, "="))
                    vFields = Split(sInfo, "|")
                    If UBound(vFields) = 5 Then
                        With m_cAddInInfo(m_lAddInCount)
                            'Name|ClassName|Description|Icon|Author|Website
                            lPos = InStr(1, vFields(0), "\")
                            .Name = Right$(vFields(0), Len(vFields(0)) - lPos)
                            If lPos > 0 Then .Category = Left$(vFields(0), lPos - 1)
                            .ClassName = vFields(1)
                            .Description = vFields(2)
                            .Icon = vFields(3)
                            .Author = vFields(4)
                            .Website = vFields(5)
                            .ShowInMenu = GetSetting(REG_KEY & "AddIns", .ClassName, "ShowInMenu", 1)
                            .ShowInTB = GetSetting(REG_KEY & "AddIns", .ClassName, "ShowInTB", 0)
                            .LoadAtStartup = GetSetting(REG_KEY & "AddIns", .ClassName, "LoadAtStartup", 0)
                        End With
                        'we've had an item before...
                        'do stuff
                        InitAddIn (m_lAddInCount)
                    Else
                        'ignore
                    End If
                End If
            End If
        Loop
        Close iFileNum
    End If
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "AddIns.ProcessAddIns"
End Sub

Private Sub InitAddIn(lIndex As Long)
Dim lNew As Long
    With m_cAddInInfo(lIndex)
        If .ShowInMenu Or .ShowInTB Then AddAddInMenu .Name, .Icon, lIndex, .ShowInMenu, .ShowInTB
        If .LoadAtStartup Then
            lNew = Add(lIndex)
        End If
    End With
End Sub

Private Sub AddAddInMenu(sName As String, sIcon As String, lIndex As Long, bShowMenu As Boolean, bShowTB As Boolean)
    Dim sImageKey   As String
    Dim lImgIndex   As Long
    'adds an add-in to the tools menu
    If sIcon = "-1" Or sIcon = "" Then
        'no image
    Else
        'get its key
        sImageKey = UCase$("AddInIcon" & sName)
        'load the image from the specified file, and get its index
        If Left$(sIcon, 1) = "\" Then sIcon = App.Path & sIcon
        lImgIndex = frmMainForm.vbalMain.AddFromFile(sIcon, 1, sImageKey)
        If lImgIndex <> 0 Then lImgIndex = IndexForKey(sImageKey)
        'If lImgIndex = 0 Then cDialog.ErrHandler 0, "Error loading icon (" & sIcon & ") for " & sName & " add-in.", "Startup.LoadMenus"
    End If
    If lImgIndex = 0 Then lImgIndex = -1
    If bShowMenu Then
        'add to the menu
        frmMainForm.ctlPopMenu.InsertItem sName, "ToolsSep2", "ToolsAddIn" & CStr(lIndex), , , lImgIndex
    End If
    If bShowTB Then
        If lImgIndex = -1 Then lImgIndex = IndexForKey("FILE_BAT")
        'add to the tb
        frmMainForm.tbrMain(3).AddButton sName, lImgIndex, , , , , "BTN" & CStr(lIndex)
    End If
End Sub

Private Sub Class_Terminate()
    Dim i As Long
    If m_bChanged Then
        For i = 1 To m_lAddInCount
            With m_cAddInInfo(i)
                SaveSetting REG_KEY & "AddIns", .ClassName, "ShowInMenu", Abs(.ShowInMenu)
                SaveSetting REG_KEY & "AddIns", .ClassName, "ShowInTB", Abs(.ShowInTB)
                SaveSetting REG_KEY & "AddIns", .ClassName, "LoadAtStartup", Abs(.LoadAtStartup)
            End With
        Next
    End If
    Set m_cAddInApp = Nothing
    m_lAddInCount = 0
    Erase m_cAddInInfo
End Sub

Private Function SaveINISetting(Section As String, Key As String, Value As String, IniFile As String) As Boolean
    If WritePrivateProfileString(Section, Key, Value, IniFile) Then
        SaveINISetting = True
    End If
End Function
Private Function GetINISetting(Section As String, Key As String, IniFile As String) As String
    Dim sBuffer As String
    sBuffer = String$(255, 0)
    If GetPrivateProfileString(Section, Key, Empty, sBuffer, Len(sBuffer), IniFile) Then
        GetINISetting = Left$(sBuffer, InStr(sBuffer, Chr$(0)) - 1)
    End If
End Function
