VERSION 5.00
Begin VB.Form frmEntry 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Entry"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   289
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   359
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "1001"
      Height          =   345
      Left            =   150
      TabIndex        =   5
      Top             =   3900
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "1002"
      Height          =   345
      Left            =   1455
      TabIndex        =   6
      Top             =   3900
      Width           =   1215
   End
   Begin VB.PictureBox picCursorEnd 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1920
      ScaleHeight     =   375
      ScaleWidth      =   15
      TabIndex        =   1
      ToolTipText     =   "SelEnd"
      Top             =   1320
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "1009"
      Height          =   330
      Left            =   4095
      Picture         =   "frmEntry.frx":000C
      TabIndex        =   8
      Top             =   825
      Width           =   1092
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "1010"
      Height          =   330
      Left            =   2925
      Picture         =   "frmEntry.frx":00C6
      TabIndex        =   7
      Top             =   825
      Width           =   1092
   End
   Begin VB.CheckBox chkAfterText 
      Caption         =   "1076"
      Height          =   270
      Left            =   3555
      TabIndex        =   3
      Top             =   465
      Width           =   1545
   End
   Begin VB.PictureBox picCursorPos 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1485
      ScaleHeight     =   375
      ScaleWidth      =   15
      TabIndex        =   11
      ToolTipText     =   "SelStart"
      Top             =   1320
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.TextBox txtValue 
      Height          =   2415
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   1215
      Width           =   5055
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   1245
      TabIndex        =   0
      Top             =   120
      Width           =   3930
   End
   Begin DevPad.vbwSpinner txtPos 
      Height          =   330
      Left            =   1245
      TabIndex        =   13
      Top             =   465
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   582
   End
   Begin DevPad.vbwSpinner txtLen 
      Height          =   330
      Left            =   2760
      TabIndex        =   2
      Top             =   465
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   582
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   10
      X2              =   347
      Y1              =   250
      Y2              =   250
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      Caption         =   "1240"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   4770
      MouseIcon       =   "frmEntry.frx":0180
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   3960
      Width           =   360
   End
   Begin VB.Label lblLabel 
      Caption         =   "1203"
      Height          =   255
      Index           =   3
      Left            =   2025
      TabIndex        =   12
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblLabel 
      Caption         =   "1074"
      Height          =   255
      Index           =   0
      Left            =   135
      TabIndex        =   10
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblLabel 
      Caption         =   "1130"
      Height          =   255
      Index           =   2
      Left            =   135
      TabIndex        =   9
      Top             =   135
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   10
      X2              =   346
      Y1              =   251
      Y2              =   251
   End
End
Attribute VB_Name = "frmEntry"
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
' You may not redisibute this source code,
' or disibute re-compiled versions of
' Developers Pad
'
Option Explicit
Implements ISubclass
Private Const WM_HSCROLL = &H114

Private lCurItem    As Long    'current item
Private bIgnore     As Boolean 'ignore triggered event
Private bChanged    As Boolean 'data changed?
Public bNoUnload    As Boolean 'don't unload form
Public Sub DisplayEntry(bAdd As Boolean, sDescription As String, sValue As String, lPos As Long, lLen As Long, lEntry As Long)
On Error GoTo ErrHandler
    
    If bAdd Then
        cmdOK.Caption = LoadResString(1020) 'Add
    Else
        cmdOK.Caption = LoadResString(112) 'Save
    End If
    'fill the fields
    txtDescription = sDescription
    'ignore triggered events
    bIgnore = True
    txtValue = sValue
    'allow triggered events...
    bIgnore = False
    'need to set max/min before setting value
    txtPos.Max = Len(sValue)
    'check the checkbox if necessary...
    chkAfterText.Value = Abs(lPos = 0)
    'trigger it's click event
    Call chkAfterText_Click
    'if lpos = 0, then we want the end of the text
    If lPos = 0 Then lPos = Len(sValue)
    'update the position...
    txtPos.Value = lPos
    'set the max selection length value
    txtLen.Max = Len(sValue) - txtPos.Value
    txtLen.Value = lLen
    'update spinner values
    UpdateValues
    'data not changed
    bChanged = False
    lCurItem = lEntry
    cmdCancel.Caption = LoadResString(1002) 'cancel
    'show form
    If Visible = False Then Show , frmMainForm
    txtDescription.SetFocus
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Entry.Display"
End Sub
Private Sub chkAfterText_Click()
    'update the start and len spinners...
    UpdateValues
    'something has changed!
    bChanged = True
End Sub
'move to the next/last entry
Private Sub cmdLast_Click()
    frmProject.LastEntry
End Sub
Private Sub cmdNext_Click()
    frmProject.NextEntry
End Sub
'validates the current entry
Public Function ValidateEntry(Optional bNoPrompt As Boolean = False) As Boolean
Dim vAns As ShowYesNoResult
    If bChanged = True Then
        If cmdOK.Caption = LoadResString(1020) Or bNoPrompt Then 'add
            vAns = Yes
        Else
            vAns = cDialog.ShowYesNo(LoadResString(1237), True) '"Save Changes to entry?"
        End If
        Select Case vAns
        Case Yes
            If txtValue.Text = "" And txtDescription.Text <> "" Then
                cDialog.ShowWarning LoadResString(1238), "Entry.cmdOK_Click()" '"You cannot have an empty entry"
                Exit Function
            ElseIf txtDescription.Text = "" And txtValue.Text <> "" Then
                cDialog.ShowWarning LoadResString(1239) '"You must enter a name for the entry"
                Exit Function
            End If
            ' are both required fields filled?
            If txtValue.Text <> "" And txtDescription.Text <> "" Then
                If cmdOK.Caption = LoadResString(112) Then  'Save
                    'save the entry
                    frmProject.ChangeEntry txtDescription, txtValue.Text, IIf(chkAfterText.Value = 1, 0, txtPos.Value), txtLen.Value, lCurItem
                    'unload unless told otherwise!
                    If bNoUnload = False Then Unload Me
                    bNoUnload = False
                Else
                    'add the entry
                    frmProject.AddEntry txtDescription, txtValue.Text, IIf(chkAfterText.Value = 1, 0, txtPos.Value), txtLen.Value
                    'reset the fields...
                    txtDescription = ""
                    txtPos.Value = 0
                    txtValue = ""
                    chkAfterText.Value = 1
                    Call chkAfterText_Click
                    txtDescription.SetFocus
                End If
                bChanged = False
            End If
            'ValidateEntry = True
        'Case No
            'ValidateEntry = True
        Case Cancelled
            'actually failed..., but you'll see!
            ValidateEntry = True
        End Select
    'Else
        'ValidateEntry = True
    End If
    'flip result flag
    ValidateEntry = Not ValidateEntry
    'change Cancel button to Close if successful
    If ValidateEntry Then cmdCancel.Caption = LoadResString(1000) 'close
End Function
Private Sub cmdOK_Click()
    'validate the current entry
    ValidateEntry True
End Sub
Private Sub Form_Load()
    'attach subclassing...
    AttachMessage Me, txtValue.hWnd, WM_VSCROLL
    AttachMessage Me, txtValue.hWnd, WM_HSCROLL
    'load the res strings
    LoadResStrings Controls
    'set the highlight pictureboxes to the right height
    picCursorPos.Height = TextHeight("Test")
    picCursorEnd.Height = TextHeight("Test")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'detach subclassing...
    DetachMessage Me, txtValue.hWnd, WM_VSCROLL
    DetachMessage Me, txtValue.hWnd, WM_HSCROLL
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As vbwSubClass.EMsgResponse)
End Property

Private Property Get ISubclass_MsgResponse() As vbwSubClass.EMsgResponse
    ISubclass_MsgResponse = emrPreprocess
End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    UpdateValues
End Function

Private Sub lblHelp_Click()
    cDialog.ShowHelpTopic 1, hWnd
End Sub
Private Sub txtDescription_Change()
    bChanged = True
End Sub
Private Sub txtLen_Change()
    UpdateValues True
End Sub
Private Sub txtPos_Change()
    UpdateValues True
End Sub
'updates the cursor pos/len entries...
Private Sub UpdateValues(Optional bSpinTriggered As Boolean = False)
    'it has changed...
    bChanged = True
    'set the maximum values for the spinners
    txtPos.Max = Len(txtValue.Text)
    txtLen.Max = Len(txtValue.Text) - txtPos.Value
    If chkAfterText.Value = 1 And bSpinTriggered = False Then
        'AfterText checkbox has been checked...
        'Set length to 0, and pos to the end of text
        txtLen.Value = 0
        txtPos.Value = Len(txtValue.Text)
        'hide cursor sel pic boxes...
        picCursorPos.Visible = False
        picCursorEnd.Visible = False
    Else
        'show cursor sel pic boxes
        picCursorPos.Visible = True
        'move them to the correct position...
        ShowCursorPos picCursorPos, txtPos.Value
        If txtLen.Value <> 0 Then
            'move the end sel pic box too...
            ShowCursorPos picCursorEnd, txtPos.Value + txtLen.Value
        Else
            'no selection length to display... hide it instead
            picCursorEnd.Visible = False
        End If
    End If
    'Adjust the checkbox state...
    chkAfterText.Value = Abs(txtLen.Value = 0 And txtPos.Value = Len(txtValue.Text))
End Sub
Private Sub txtValue_Change()
    If bIgnore Then Exit Sub
    UpdateValues
End Sub
'used to move display cursors
Private Sub ShowCursorPos(picObj As PictureBox, lSelStart As Long)
    Dim xPixels As Long
    Dim yPixels As Long
    'Get xy co-ordinates from char pos...
    GetPosFromChar lSelStart, xPixels, yPixels
    'don't ask me why... but sometimes it returns this..
    If xPixels = 65535 Then
        'if it does, reduce the pos by 1
        GetPosFromChar txtPos.Value - 1, xPixels, yPixels
        xPixels = xPixels + TextWidth(Right$(txtValue.Text, 1))
    End If
    'move the sel object to correct position
    picObj.Left = txtValue.Left + xPixels ' + 1
    picObj.Top = txtValue.Top + yPixels + 2
    picObj.Visible = True
End Sub
'returns xy co-ordinates from sel pos
Private Sub GetPosFromChar(ByVal lIndex As Long, ByRef xPixels As Long, ByRef yPixels As Long)
Dim lxy As Long
Const EM_POSFROMCHAR = &HD6&
   lxy = SendMessageLong(txtValue.hWnd, EM_POSFROMCHAR, lIndex, 0)
   xPixels = (lxy And &HFFFF&)
   yPixels = (lxy \ &H10000) And &HFFFF&
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub
