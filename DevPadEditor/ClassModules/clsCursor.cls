VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCursor"
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
Option Explicit

Private m_lngPrevCursor As MousePointerConstants

Private Sub Class_Initialize()
    ' Store the current Screen.MousePointer
    m_lngPrevCursor = Screen.MousePointer
End Sub
Private Sub Class_Terminate()
    ResetCursor
End Sub
Public Sub SetCursor(vCursor As MousePointerConstants)
    ResetCursor
    ' Apply the new cursor
    Screen.MousePointer = vCursor
End Sub

Public Sub ResetCursor()
    ' Reset the cursor to the one stored in Class_Initalize
    Screen.MousePointer = m_lngPrevCursor
End Sub

