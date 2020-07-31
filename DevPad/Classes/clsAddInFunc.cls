VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAddInApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Private cAddInDlg As clsAddInDialog
Private cAddInMenu As clsAddInMenu

Implements IDevPadApp

Private Property Get IDevPadApp_ApphWnd() As Long
    IDevPadApp_ApphWnd = frmMainForm.hWnd
End Property

Private Property Get IDevPadApp_AppPath() As String
    IDevPadApp_AppPath = App.Path
End Property

Private Property Get IDevPadApp_Dialogs() As DevPadAddInTlb.IDevPadDialog
    'init the class if we haven't already
    If cDialog Is Nothing Then Set cDialog = New clsDialog
    'return it
    Set IDevPadApp_Dialogs = cDialog
End Property

'Private Property Get IDevPadApp_DocumentByhWnd(ByVal hWnd As Long) As DevPadAddInTlb.IDevPadDocument
'
'End Property
'
'Private Property Get IDevPadApp_DocumentByIndex(ByVal ID As Long) As DevPadAddInTlb.IDevPadDocument
'
'End Property
'
Private Property Get IDevPadApp_Documents() As DevPadAddInTlb.IDevPadDocuments
    Set IDevPadApp_Documents = cDocuments
End Property

Private Property Get IDevPadApp_Menus() As DevPadAddInTlb.IDevPadMenu
    'init the class if we haven't already
    If cAddInMenu Is Nothing Then Set cAddInMenu = New clsAddInMenu
    'return it
    Set IDevPadApp_Menus = cAddInMenu
End Property
'
'
'Private Property Get IDevPadApp_DocumentCount() As Long
'Dim lCount As Long
'Dim i As Long
'    For i = 0 To Forms.Count
'        If Forms(i).Name = "frmDocument" Then
'            lCount = lCount + 1
'        End If
'    Next
'    IDevPadApp_DocumentCount = lCount
'End Property

