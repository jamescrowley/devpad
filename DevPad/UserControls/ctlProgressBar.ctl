VERSION 5.00
Begin VB.UserControl vbwProgressBar 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5310
   ScaleHeight     =   960
   ScaleWidth      =   5310
   Begin VB.PictureBox picStatus 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      FillColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   384
      Left            =   0
      ScaleHeight     =   390
      ScaleWidth      =   5310
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5310
   End
End
Attribute VB_Name = "vbwProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//////////////////////////////////////////////////////////////
'// This is the example program for the Progress Bar Control //
'//                                                          //
'//                 Copyright VB Web 1999                    //
'//                   www.vbweb.f9.co.uk                     //
'//                 support@vbweb.f9.co.uk                   //
'//               PLEASE REPORT ANY BUGS YOU FIND            //
'//////////////////////////////////////////////////////////////

Option Explicit
'// Property Variables:
Dim m_bShowStatus            As Boolean
Dim m_fValue                 As Single
Dim m_sText                  As String
Dim m_sTextAfterProgress     As String

'///////////////////////////////////////////////////////////
Public Property Get Value() As Single
    Value = m_fValue
End Property
Public Property Let Value(ByVal New_Value As Single)
    If Abs(m_fValue - New_Value) < 0.5 Then
        Exit Property
    End If
    m_fValue = New_Value
    DoEvents
    UpdateStatus
    PropertyChanged "Value"
End Property

'///////////////////////////////////////////////////////////
Public Property Get ShowStatus() As Boolean
    ShowStatus = m_bShowStatus
End Property
Public Property Let ShowStatus(ByVal New_ShowStatus As Boolean)
    m_bShowStatus = New_ShowStatus
    UpdateStatus
    PropertyChanged "ShowStatus"
End Property
'///////////////////////////////////////////////////////////
Public Property Get Text() As String
    Text = m_sText
End Property
Public Property Let Text(ByVal New_Text As String)
    m_sText = New_Text
    UpdateStatus
    PropertyChanged "Text"
End Property
'///////////////////////////////////////////////////////////
Public Property Get Font() As Font
    Set Font = picStatus.Font
    UpdateStatus
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set picStatus.Font = New_Font
    UpdateStatus
    PropertyChanged "Font"
End Property
'///////////////////////////////////////////////////////////
Public Property Get ProgressColor() As OLE_COLOR
    ProgressColor = picStatus.ForeColor
End Property
Public Property Let ProgressColor(ByVal New_ProgressColor As OLE_COLOR)
On Error GoTo ErrHandler
    '// For this to work well, we need a white background and any color foreground
    picStatus.ForeColor = New_ProgressColor
    
    UpdateStatus
    PropertyChanged "ProgressColor"
    Exit Property
ErrHandler:
    Err.Raise Err, "vbwProgressBar", Error
End Property

Public Property Get TextAfterProgress() As String
    TextAfterProgress = m_sTextAfterProgress
End Property
Public Property Let TextAfterProgress(ByVal New_TextAfterProgress As String)
    m_sTextAfterProgress = New_TextAfterProgress
    UpdateStatus
    PropertyChanged "TextAfterProgress"
End Property

Private Sub UserControl_Initialize()
     SetThin3DBorder UserControl.hWnd
End Sub

'///////////////////////////////////////////////////////////
Private Sub UserControl_Resize()
    picStatus.Height = ScaleHeight
    SetProgress m_fValue
End Sub
'///////////////////////////////////////////////////////////
Private Sub SetProgress(Optional ByVal fPercent As Single = 0, Optional ByVal dblCurrentValue As Single = 0)
On Error GoTo ErrHandler
    Dim sPercent As String
    Dim iX As Integer
    Dim iY As Integer
    Dim iWidth As Integer
    Dim iHeight As Integer
    Dim iPercent As Integer
    Dim sString As String
    '// we need to update backcolor (even though it doesn't change)
    '// to stop flickering
    picStatus.BackColor = vbWhite
    '// Format percentage and get attributes of text
    iPercent = Int(fPercent + 0.5) 'Int(100 * fPercent + 0.5)
    fPercent = fPercent / 100
    
    '// Never allow the percentage to be 0 or 100 unless it is exactly that value.  This
    '// prevents, for instance, the status bar from reaching 100% until we are entirely done.
    If m_bShowStatus = True Then sPercent = Format$(iPercent) & "%"
    sString = m_sText & IIf(m_bShowStatus, " " & sPercent, "") & IIf(m_sTextAfterProgress <> "", " " & m_sTextAfterProgress, "")

    'iWidth = picStatus.TextWidth(sString)
    iHeight = picStatus.TextHeight(sString)
    
    '// Now set iX and iY to the starting location for priing the percentage
    iX = 50 '(picStatus.ScaleWidth / 2) - (iWidth / 2)
    iY = (picStatus.ScaleHeight / 2) - (iHeight / 2)

    '// Need to draw a filled box with the picStatus background color to wipe out previous
    '// percentage display (if any)
    picStatus.DrawMode = 13 '// Copy Pen
    picStatus.Line (iX, iY)-(iWidth, iHeight), picStatus.BackColor

    '// Back to the center pri position and pri the text
  '  picStatus.CurrentX = 30
    picStatus.CurrentX = iX
    picStatus.CurrentY = iY
    
    picStatus.Print m_sText & " " & sPercent & " " & m_sTextAfterProgress
    '// Now fill in the box with the ribbon color to the desired percentage
    '// If percentage is 0, fill the whole box with the background color to clear it
    '// Use the "Not XOR" pen so that we change the color of the text to white
    '// wherever we touch it, and change the color of the background to blue
    '// wherever we touch it.
    picStatus.DrawMode = 10 '// Not XOR Pen
    If fPercent > 0 Then
        picStatus.Line (0, 0)-(picStatus.Width * fPercent, picStatus.Height), picStatus.ForeColor, BF
    Else
        picStatus.Line (0, 0)-(picStatus.Width, picStatus.Height), picStatus.BackColor, BF
    End If
    Exit Sub
ErrHandler:
    Err.Raise Err, "vbwProgressBar", Error
End Sub
Private Sub UpdateStatus()
    SetProgress m_fValue
End Sub
'// Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_bShowStatus = True
    picStatus.ForeColor = &H800000
    m_sTextAfterProgress = ""
End Sub

'// Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_fValue = PropBag.ReadProperty("Value", 0)
    m_bShowStatus = PropBag.ReadProperty("ShowStatus", True)
    m_sText = PropBag.ReadProperty("Text", Empty)
    Set picStatus.Font = PropBag.ReadProperty("Font", Ambient.Font)
    picStatus.ForeColor = PropBag.ReadProperty("ProgressColor", &H800000)
    m_sTextAfterProgress = PropBag.ReadProperty("TextAfterProgress", "")
End Sub

'// Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Value", m_fValue, 0)
    Call PropBag.WriteProperty("ShowStatus", m_bShowStatus, True)
    Call PropBag.WriteProperty("Text", m_sText, Empty)
    Call PropBag.WriteProperty("Font", picStatus.Font, Ambient.Font)
    Call PropBag.WriteProperty("ProgressColor", picStatus.ForeColor, &H800000)
    Call PropBag.WriteProperty("TextAfterProgress", m_sTextAfterProgress, "")
End Sub



