VERSION 5.00
Begin VB.Form frmSave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Developers Pad"
   ClientHeight    =   2805
   ClientLeft      =   30
   ClientTop       =   240
   ClientWidth     =   6000
   Icon            =   "frmSave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstDocuments 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   600
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   720
      Width           =   3855
   End
   Begin VB.CommandButton cmdSaveSelected 
      Caption         =   "1024"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4560
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "1002"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4560
      TabIndex        =   2
      Top             =   900
      Width           =   1335
   End
   Begin VB.CommandButton cmdDiscardAll 
      Caption         =   "1025"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4560
      TabIndex        =   1
      Top             =   510
      Width           =   1335
   End
   Begin VB.Image imgImage 
      Height          =   480
      Index           =   2
      Left            =   0
      Picture         =   "frmSave.frx":000C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblLabel 
      Caption         =   "1178"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
Dim colForms As Collection
Dim bFormsSpecified As Boolean

Public Sub PrepareForExit(Optional frmForms As Variant)
On Error GoTo ErrHandler

    'Dim frmForm As Form
    Dim i As Long
    'prepares devpad for closing
    Set colForms = New Collection
    
    If IsMissing(frmForms) Then
        'we haven't been passed a group of forms...
        'get all the loaded ones
        Set frmForms = Forms
        bFormsSpecified = False
    Else
        'we have been given a group of forms to unload for..
        bFormsSpecified = True
    End If
    'clear the list of documents
    lstDocuments.Clear
    For i = 1 To cDocuments.Count
        If cDocuments.Item(i).Modified = True Then
            'if it is a document, and it has been modified,
            'add it to the list
            lstDocuments.AddItem cDocuments.Item(i).DocumentCaption
            'select the item
            lstDocuments.Selected(lstDocuments.NewIndex) = True
            'add the hWnd to the collection
            colForms.Add cDocuments.Item(i).DocID
            'reset save option
            cDocuments.Item(i).SaveOption = vbwNotSet
        End If
    Next
    'unassign reference
    Set frmForms = Nothing
    If lstDocuments.ListCount <> 0 Then
        'if devpad is minimized, we HAVE TO restore!
        'otherwise, for some reason, this form can't get the focus
        'and the user is stuck!
        If frmMainForm.WindowState = vbMinimized Then frmMainForm.WindowState = vbNormal
        'display the form
        Show vbModal, frmMainForm
    Else
        'no unsaved docs... unload them
        UnloadForms
    End If
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Save.PrepareForExit"
End Sub

Private Sub cmdCancel_Click()
    'abort exit...
    frmMainForm.bCancelClose = True
    'unload this form
    Unload Me
End Sub

Private Sub cmdDiscardAll_Click()
    Dim i As Long
    Hide
    On Error Resume Next
    'loop through all the forms, and set their flag to discard
    For i = 0 To lstDocuments.ListCount - 1
        cDocuments.ItemByID(colForms(i + 1)).SaveOption = vbwDiscard
    Next
    'flags set...unload the forms
    UnloadForms
End Sub

Private Sub cmdSaveSelected_Click()
    Dim i As Long
    'hide the form
    Hide
    For i = 0 To lstDocuments.ListCount - 1
        'loop through all the forms
        If lstDocuments.Selected(i) = True Then
            'item is selected... set flag to save
            cDocuments.ItemByID(colForms(i + 1)).SaveOption = vbwSave
        Else
            'item is not selected... set flag to discard
            cDocuments.ItemByID(colForms(i + 1)).SaveOption = vbwDiscard
        End If
    Next
    'flags set... unload the forms
    UnloadForms
End Sub
Private Sub UnloadForms()
On Error GoTo ErrHandler
    'unloads all the forms in the collection
   ' Dim frmForm As Form
    Dim i As Long
    'give way for a sec... makes a tidier exit!
    DoEvents
    'reset abort flag
    frmMainForm.bCancelClose = False
    'we are closing multiple... don't worry about the statusbar
    bClosingMultiple = True
    If bFormsSpecified Then
        'only deal with specified forms
        For i = 1 To colForms.Count
            cDocuments.ItemByID(colForms(i + 1)).Close
        Next
    Else
        'unload all of them
        For i = cDocuments.Count To 1 Step -1
            cDocuments.Item(i).Close
            If frmMainForm.bCancelClose = True Then Exit For
            DoEvents
        Next
    End If
    'reset flag
    bClosingMultiple = False
    'unload this form
    Unload Me
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Save.UnloadForms"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then cDialog.ShowHelpTopic 13, hWnd
End Sub

Private Sub Form_Load()
    'load the resource strings
    LoadResStrings Controls
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'user has clicked X button... simulate clicking on Cancel
    If UnloadMode = vbFormControlMenu Then cmdCancel_Click
End Sub
