VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFavourites"
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

Public Sub AddFavourite(sPath As String, cboFav As ComboBox)
    Dim i As Long
    For i = 0 To cboFav.ListCount
        If LCase$(cboFav.List(i)) = LCase$(sPath) Then
            cDialog.ErrHandler vbObjectError + 999, "That favourite path already exists", "Open.AddFav"
            Exit Sub
        End If
    Next
    cboFav.AddItem sPath
End Sub
Public Sub LoadFavourites(cboFav As ComboBox, frmForm As Form)  'agghh!
    Dim iFileNum As Integer
    Dim sLine As String
    Dim lLargestLine As Long
    cboFav.Clear
    On Error Resume Next
    iFileNum = FreeFile
    Open App.Path & "\favourites.txt" For Input As iFileNum
    If Err = 0 Then
        Do While Not EOF(iFileNum)
            Line Input #iFileNum, sLine
            cboFav.AddItem sLine
            If frmForm.TextWidth(sLine) > lLargestLine Then lLargestLine = frmForm.TextWidth(sLine)
        Loop
        Close iFileNum
        SendMessage cboFav.hWnd, CB_SETDROPPEDWIDTH, (lLargestLine / Screen.TwipsPerPixelX) + 30, 0&
    Else
        Open App.Path & "\favourites.txt" For Output As iFileNum
        Close iFileNum
    End If
'    cboFav.Text = ""
End Sub
Public Sub SaveFavourites(cboFav As ComboBox)
    Dim iFileNum As Integer
    Dim i As Long
    On Error GoTo ErrHandler
    iFileNum = FreeFile
    Open App.Path & "\favourites.txt" For Output As iFileNum
    For i = 0 To cboFav.ListCount - 1
        Print #iFileNum, cboFav.List(i)
    Next
    Close iFileNum
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "clsFavourites.SaveFavourites"
End Sub

