Attribute VB_Name = "modDeclares"
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

Public cFunc As DevPadAddInTlb.IDevPadApp

Public cHTMLHelp As clsHTMLHelp
Private Sub InitHTMLHelp()
    If cHTMLHelp Is Nothing Then Set cHTMLHelp = New clsHTMLHelp
    With cHTMLHelp
        .HelpPath = App.Path & "\ToolsHelp.chm"
        .Window = "MainWindow"
    End With
End Sub
Public Sub ShowHTMLHelpTopic(lTopic As Long, lhWnd As Long)
    InitHTMLHelp
    cHTMLHelp.hWnd = lhWnd
    cHTMLHelp.HTMLShowTopicByID lTopic
End Sub
