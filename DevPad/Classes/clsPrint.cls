VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrint"
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
Private m_frmPrint As frmPrinting
Private Type CHARRANGE
  cpMin             As Long      ' First character of range (0 for start of doc)
  cpMax             As Long      ' Last character of range (-1 for end of doc)
End Type
Private Type FormatRange
  hdc               As Long      ' Actual DC to draw on
  hdcTarget         As Long      ' Target DC for determining text formatting
  rc                As RECT      ' Region of the DC to draw to (in twips)
  rcPage            As RECT      ' Region of the entire DC (page size) (in twips)
  chrg              As CHARRANGE ' Range of text to draw (see above declaration)
End Type

Private Const EM_FORMATRANGE As Long = WM_USER + 57
Private Const PHYSICALOFFSETX As Long = 112
Private Const PHYSICALOFFSETY As Long = 113

Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As Long, ByVal lpInitData As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Public Function ShowPrintDialog(vFlags As Long) As Boolean
    On Error GoTo ErrHandler
    
    Dim cCmDlg As clsHookDlg
    Dim bErrOnce As Boolean

    Set cCmDlg = New clsHookDlg
    With cCmDlg
        .CancelError = True
        ' Set flags
        .Flags = vFlags
        ' set owner
        .hWnd = frmMainForm.hWnd
        .PrinterObject = Printer
        On Error Resume Next
        'display the dialog
        .ShowPrinter
        ' Dialog cancelled?
        If Err = 0 Then ShowPrintDialog = True
    End With

    Set cCmDlg = Nothing
    Exit Function
ErrHandler:
    cDialog.ErrHandler Err, Error, "Core.ShowPrintDialog"
End Function
Public Sub PrintPage()
    Dim PrintFlags As EPrintDialog
    On Error GoTo ErrHandler
    
    ' printing not cancelled
    'StopPrinting = False
    ' no documents open?
    If DocOpen = False Then Exit Sub

    ' Set print range
    If ActiveDoc.SelLength = 0 Then
        PrintFlags = PD_ALLPAGES
    Else
        PrintFlags = PD_SELECTION
    End If
    PrintFlags = PrintFlags + PD_DISABLEPRINTTOFILE
    If ShowPrintDialog(PrintFlags) = False Then Exit Sub
    ' Load the printing dialog
    If m_frmPrint Is Nothing Then
        Set m_frmPrint = New frmPrinting
        Load m_frmPrint
    End If
    With m_frmPrint
        .Cancel = False
        .Show vbModeless, frmMainForm
        .Refresh
        DoEvents
        ' Print the contents of the RichTextBox with a one inch margin
        PrintRTF ActiveDoc, 1440, 1440, 1440, 1440, ActiveDoc.DocumentCaption ' 1440 Twips = 1 Inch
        .Hide
    End With
Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Print Page"
End Sub
'
' PrintRTF - Prints the contents of a RichTextBox control using the
'            provided margins

Private Sub PrintRTF(rtf As IDevPadDocument, LeftMarginWidth As Long, TopMarginHeight As Long, RightMarginWidth As Long, BottomMarginHeight As Long, sCaption As String)
On Error GoTo ErrHandler
    Dim LeftOffset As Long, TopOffset As Long
    Dim LeftMargin As Long, TopMargin As Long
    Dim RightMargin As Long, BottomMargin As Long
    Dim fr As FormatRange
    Dim rcDrawTo As RECT
    Dim rcPage As RECT
    Dim TextLength As Long
    Dim NextCharPosition As Long
    Dim r As Long
    
    ' Start a print job to get a valid Printer.hDC
    Printer.Print Space(1)
    Printer.ScaleMode = vbTwips
    
    ' Get the offsett to the printable area on the page in twips
    LeftOffset = Printer.ScaleX(GetDeviceCaps(Printer.hdc, _
       PHYSICALOFFSETX), vbPixels, vbTwips)
    TopOffset = Printer.ScaleY(GetDeviceCaps(Printer.hdc, _
       PHYSICALOFFSETY), vbPixels, vbTwips)
    
    ' Calculate the Left, Top, Right, and Bottom margins
    LeftMargin = LeftMarginWidth - LeftOffset
    TopMargin = TopMarginHeight - TopOffset
    RightMargin = (Printer.Width - RightMarginWidth) - LeftOffset
    BottomMargin = (Printer.Height - BottomMarginHeight) - TopOffset
    
    ' Set printable area rect
    rcPage.Left = 0
    rcPage.Top = 0
    rcPage.Right = Printer.ScaleWidth
    rcPage.Bottom = Printer.ScaleHeight
    
    ' Set rect in which to print (relative to printable area)
    rcDrawTo.Left = LeftMargin
    rcDrawTo.Top = TopMargin
    rcDrawTo.Right = RightMargin
    rcDrawTo.Bottom = BottomMargin
    
    ' Set up the print instructions
    fr.hdc = Printer.hdc   ' Use the same DC for measuring and rendering
    fr.hdcTarget = Printer.hdc  ' Point at printer hDC
    fr.rc = rcDrawTo            ' Indicate the area on page to draw to
    fr.rcPage = rcPage          ' Indicate entire size of page
    fr.chrg.cpMin = 0           ' Indicate start of text through
    fr.chrg.cpMax = -1          ' end of the text
    
    ' Get length of text in RTF
    TextLength = rtf.CharacterCount
    Printer.Font.Name = vDefault.sFont
    Printer.Font.Size = vDefault.nFontSize
    
    ' Loop printing each page until done
    Do
         m_frmPrint.Message = LoadResString(1198) & ActiveDoc.DocumentCaption & LoadResString(1199) & Printer.DeviceName & " ..." & vbCrLf & LoadResString(1200) & Str(Printer.Page)
         Printer.Print "       " & sCaption & " " & LoadResString(1200) & Str(Printer.Page)
         ' Print the page by sending EM_FORMATRANGE message
         NextCharPosition = SendMessage(rtf.hWnd, EM_FORMATRANGE, True, fr)
         If NextCharPosition >= TextLength Then Exit Do  ' If done then exit
         DoEvents
         If m_frmPrint.Cancel Then
            Exit Do
         End If
         fr.chrg.cpMin = NextCharPosition ' Starting position for next page
         
         Printer.NewPage                  ' Move on to next page
         Printer.Print Space(1) ' Re-initialize hDC
    
         fr.hdc = Printer.hdc
         fr.hdcTarget = Printer.hdc
    Loop
TheEnd:
    If m_frmPrint.Cancel Then
        Printer.KillDoc
    Else
        ' Commit the print job
        Printer.EndDoc
    End If
    ' Allow the RTF to free up memory
    r = SendMessage(rtf.hWnd, EM_FORMATRANGE, False, ByVal CLng(0))
    ' finished
    Unload m_frmPrint
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "PrintRTF"
    GoTo TheEnd
End Sub
