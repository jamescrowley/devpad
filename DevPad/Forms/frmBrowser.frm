VERSION 5.00
Begin VB.Form frmBrowser 
   Caption         =   "Form1"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim IE As SHDocVw.InternetExplorer

Private Sub Form_Load()
    Dim obj As New MSHTML.HTMLWindow2
    obj.alert ("hello")
    obj.showModalDialog ("http://localnetwork/")
'    Set IE = CreateObject("InternetExplorer.Application")
'    IE.Visible = True
'    ie.Parent=
End Sub
