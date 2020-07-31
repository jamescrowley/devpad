VERSION 5.00
Begin VB.Form frmPrinting 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Printing..."
   ClientHeight    =   1005
   ClientLeft      =   30
   ClientTop       =   255
   ClientWidth     =   5040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
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
      Left            =   2010
      TabIndex        =   1
      Top             =   615
      Width           =   1092
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   4875
   End
End
Attribute VB_Name = "frmPrinting"
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
Private m_bCancel As Boolean
'set the printing message...
Public Property Let Message(sNew As String)
    lblMsg.Caption = sNew
    DoEvents
End Property
'return cancel flag
Public Property Get Cancel() As Boolean
    Cancel = m_bCancel
End Property
'update cancel flag
Public Property Let Cancel(bNew As Boolean)
    m_bCancel = bNew
End Property
Private Sub cmdCancel_Click()
    'dialog cancelled
    m_bCancel = True
End Sub
Private Sub Form_Load()
    LoadResStrings Controls
End Sub
