VERSION 5.00
Object = "{C4925FC3-1606-11D4-82BB-004005AAE138}#5.2#0"; "VBWIML.OCX"
Object = "{017E002E-D7CC-11D2-8E21-44B10AC10000}#22.0#0"; "VBWGRID.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmCOMInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "COM Information"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCOMInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select..."
      Height          =   345
      Left            =   4965
      TabIndex        =   13
      Top             =   345
      Width           =   1215
   End
   Begin VB.TextBox txtLibrary 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1005
      TabIndex        =   8
      Top             =   90
      Width           =   5895
   End
   Begin VB.TextBox txtFile 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1005
      TabIndex        =   7
      Top             =   375
      Width           =   3870
   End
   Begin VB.TextBox txtGUID 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1005
      TabIndex        =   6
      Top             =   660
      Width           =   3870
   End
   Begin VB.TextBox txtVersion 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1005
      TabIndex        =   5
      Top             =   930
      Width           =   3870
   End
   Begin RichTextLib.RichTextBox rtfInfo 
      Height          =   780
      Left            =   2280
      TabIndex        =   4
      Top             =   3840
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   1376
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmCOMInfo.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vbwIml.vbalImageList imlIcons 
      Left            =   6570
      Top             =   660
      _ExtentX        =   953
      _ExtentY        =   953
      Size            =   5640
      Images          =   "frmCOMInfo.frx":0090
      KeyCount        =   6
      Keys            =   "CLASSÿENUMÿPROPÿMETHODÿEVENTÿENUMITEM"
   End
   Begin vbAcceleratorGrid.vbalGrid lstClass 
      Height          =   3225
      Left            =   60
      TabIndex        =   2
      Top             =   1380
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   5689
      RowMode         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Header          =   0   'False
      DisableIcons    =   -1  'True
      DefaultRowHeight=   15
   End
   Begin VB.PictureBox picHolder 
      Align           =   2  'Align Bottom
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
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   7110
      TabIndex        =   0
      Top             =   4635
      Width           =   7110
      Begin VB.Label lblLabel 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Alpha 1.0 Release"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00B1523B&
         Height          =   225
         Index           =   0
         Left            =   5400
         TabIndex        =   14
         Top             =   30
         Width           =   1605
      End
      Begin VB.Label lblStatus 
         Height          =   210
         Left            =   45
         TabIndex        =   1
         Top             =   60
         Width           =   8460
      End
   End
   Begin vbAcceleratorGrid.vbalGrid lstMembers2 
      Height          =   2415
      Left            =   2295
      TabIndex        =   3
      Top             =   1380
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   4260
      RowMode         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Header          =   0   'False
      DisableIcons    =   -1  'True
      DefaultRowHeight=   15
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Library"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   60
      TabIndex        =   12
      Top             =   120
      Width           =   825
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "GUID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   60
      TabIndex        =   11
      Top             =   675
      Width           =   825
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   60
      TabIndex        =   10
      Top             =   390
      Width           =   825
   End
   Begin VB.Label lblLabel 
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   60
      TabIndex        =   9
      Top             =   960
      Width           =   825
   End
End
Attribute VB_Name = "frmCOMInfo"
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
Private m_cTLI As TypeLibInfo

Private m_sInterfaces() As String
Private m_iTypeInfoIndex() As Long
Private m_iBelongsToInterface() As Long
Private m_bHidden() As Boolean
Private m_iCount As Long

Dim sMembers() As String
Dim sHelp() As String
Dim iMemberID() As Long
Dim iMemberCount As Long
Dim sEvents() As String
Dim iEventID() As Long
Dim sEventHelp() As String
Dim iEventCount As Long

Private Function pbGetTypeLibInfo( _
        ByVal sFIle As String _
    ) As Boolean
On Error GoTo pGetTypeLibInfoError
    
    ' Clear up info we're holding about previous TypeLib, if any:
    m_iCount = 0
    Erase m_sInterfaces
    Erase m_iTypeInfoIndex
    Erase m_iBelongsToInterface
    lstClass.Clear
    
    
    lstClass.AddRow
    lstClass.CellText(1, 1) = "<No Interfaces>"
    lstClass.SelectedRow = 1
    lstClass.Enabled = False
    lstClass.Redraw = False
    
'    cboClass.AddItem "<No Interfaces>"
'    cboClass.ListIndex = 0
'    cboClass.Enabled = False
 '   lstGeneral.Clear
    'lstMembers.Clear
    lstMembers2.Clear
    
    ' Generate a TypeLibInfo object for the specified file.
    Status "Linking to Type Library..."
    Set m_cTLI = TLI.TypeLibInfoFromFile(sFIle)
    
    'Me.Caption = App.Title & " (" & sFIle & ")"
    ' If we succeed, then organize the TypeInfo members.
    ' VB classes have a number of components which are normally hidden from you:
        ' -the CoClass, which has the correct name but is empty because all its functions
        '   are performed by the members with _ before the name,
        ' -one or two DispInterface items, which underscores first.  The first has one underscore
        '   and contains the non-event interfaces.  The second has two and contains the events.
    Dim iTypeInfo As Long
    Dim sName As String
    Dim sBelongsTo As String
    Dim iCheckOwner As Long
    
    ' Populate general information:
    With m_cTLI
        txtLibrary.Text = .Name & " (" & .HelpString & ")"
        'lstGeneral.ItemData(lstGeneral.NewIndex) = &HFFFFFFF
        'lstGeneral.AddItem "File:" & vbTab & sFIle
        txtFile.Text = sFIle
      '  lstGeneral.ItemData(lstGeneral.NewIndex) = &HFFFFFFF
        txtGUID = .Guid
       ' lstGeneral.AddItem "GUID:" & vbTab & .Guid
      '  lstGeneral.ItemData(lstGeneral.NewIndex) = &HFFFFFFF
        txtVersion = .MajorVersion & "." & .MinorVersion
        
'        lstGeneral.AddItem "Version:" & vbTab & .MajorVersion & "." & .MinorVersion
'        lstGeneral.ItemData(lstGeneral.NewIndex) = &HFFFFFFF
    End With
    
    Status "Counting Type Library Members..."
    With m_cTLI
        ' Items with an attribute mask = 16 are old interfaces:
        m_iCount = .TypeInfoCount
        ReDim Preserve m_sInterfaces(1 To m_iCount) As String
        ReDim Preserve m_iTypeInfoIndex(1 To m_iCount) As Long
        ReDim Preserve m_iBelongsToInterface(1 To m_iCount) As Long
        ReDim Preserve m_bHidden(1 To m_iCount) As Boolean
        For iTypeInfo = 1 To m_iCount
            Debug.Print .TypeInfos(iTypeInfo).Name, .TypeInfos(iTypeInfo).AttributeMask
            m_sInterfaces(iTypeInfo) = .TypeInfos(iTypeInfo).Name
            m_iTypeInfoIndex(iTypeInfo) = iTypeInfo
            m_bHidden(iTypeInfo) = (.TypeInfos(iTypeInfo).AttributeMask = 16)
        Next iTypeInfo
    End With
    
    Status "Checking for Related VB Type Libraries and Parsing..."
    For iTypeInfo = 1 To m_iCount
        If Not (m_bHidden(iTypeInfo)) Then
            If (Left$(m_sInterfaces(iTypeInfo), 1) = "_") Then
                sBelongsTo = Mid$(m_sInterfaces(iTypeInfo), 2)
                If (Left$(sBelongsTo, 1) = "_") Then
                    sBelongsTo = Mid$(sBelongsTo, 2)
                End If
                For iCheckOwner = 1 To m_iCount
                    If (iCheckOwner <> iTypeInfo) Then
                        If Not (m_bHidden(iCheckOwner)) Then
                            If (m_sInterfaces(iCheckOwner) = sBelongsTo) Then
                                m_iBelongsToInterface(iTypeInfo) = iCheckOwner
                                Exit For
                            End If
                        End If
                    End If
                Next iCheckOwner
            End If
        End If
    Next iTypeInfo
    
    ' Add to the combo box:
    If (iTypeInfo > 0) Then
        Status "Adding Type Library Members..."
        lstClass.Clear
       ' cboClass.Clear
'        lstClass.AddRow
'        lstClass.CellText(lstClass.Rows, 1) = "<All Interfaces>"
'       ' cboClass.AddItem "<All Interfaces>"
'        lstClass.RowTag(lstClass.Rows) = &HFFFFFFF
       ' cboClass.ItemData(cboClass.NewIndex) = &HFFFFFFF
        For iTypeInfo = 1 To m_iCount
            If Not (m_bHidden(iTypeInfo)) Then
                If (m_iBelongsToInterface(iTypeInfo) = 0) Then
                    lstClass.AddRow
                    If (m_cTLI.TypeInfos(iTypeInfo).TypeKindString = "enum") Then
                        
                        lstClass.CellText(lstClass.Rows, 1) = "Enum " & m_sInterfaces(iTypeInfo) '"<All Interfaces>"
                        lstClass.CellIcon(lstClass.Rows, 1) = 1
                        'cboClass.AddItem "Enum " & m_sInterfaces(iTypeInfo)
                    Else
                        lstClass.CellText(lstClass.Rows, 1) = m_sInterfaces(iTypeInfo)
                        lstClass.CellIcon(lstClass.Rows, 1) = 0
                        'cboClass.AddItem m_sInterfaces(iTypeInfo)
                    End If
                    lstClass.RowTag(lstClass.Rows) = iTypeInfo
                    
                 '   cboClass.ItemData(cboClass.NewIndex) = iTypeInfo
                End If
            End If
        Next iTypeInfo
        lstClass.Enabled = True
        lstClass.Redraw = True
        lstClass.SelectedRow = 1
        lstClass.AutoWidthColumn 1
        
        
'        cboClass.Enabled = True
'        cboClass.ListIndex = 0
    End If
            
    Status "Ready."
    
    pbGetTypeLibInfo = True
    Exit Function
pGetTypeLibInfoError:
    MsgBox "Failed to get type lib info for file: '" & sFIle & "'" & vbCrLf & vbCrLf & Err.Description, vbExclamation
    Set m_cTLI = Nothing
    Exit Function
End Function

Private Sub Status(ByVal sStatus As String)
    lblStatus.Caption = sStatus
    lblStatus.Refresh
End Sub
'
'Private Sub cboClass_Click()
'Dim iTypeInfo As Long
'Dim i As Long
'Dim sRtf As String
'Dim sTypeLibName As String
'Dim sDateString As String
'Dim sTypeLibString As String
'
'    ' Clear list
'    Status "Getting Type Library Information..."
'    lstMembers2.Clear
'
'    If (cboClass.ListIndex > -1) Then
'        ' Evaluate the contents:
'        iTypeInfo = cboClass.ItemData(cboClass.ListIndex)
'
'        If (iTypeInfo > 0) Then
'
'            Screen.MousePointer = vbHourglass
'
'            ' Prepare the RTF header:
'            sDateString = "yr" & Year(Now) & "\mo" & Month(Now) & "\dy" & Day(Now) & "\hr" & Hour(Now) & "\min" & Minute(Now)
'            sTypeLibString = m_cTLI.Name
'
''            sRtf = "{\rtf1\ansi\ansicpg1252\uc1 \deff0\deflang1033\deflangfe1033{\fonttbl{\f0\froman\fcharset0\fprq2{\*\panose 02020603050405020304}Times New Roman;}{\f1\fswiss\fcharset0\fprq2{\*\panose 020b0604020202020204}Arial;}" & vbCrLf
''            sRtf = sRtf & "{\f2\fmodern\fcharset0\fprq1{\*\panose 02070309020205020404}Courier New;}{\f15\fswiss\fcharset0\fprq2{\*\panose 020b0604030504040204}Verdana;}}{\colortbl;\red0\green0\blue0;\red0\green0\blue255;\red0\green255\blue255;\red0\green255\blue0;" & vbCrLf
''            sRtf = sRtf & "\red255\green0\blue255;\red255\green0\blue0;\red255\green255\blue0;\red255\green255\blue255;\red0\green0\blue128;\red0\green128\blue128;\red0\green128\blue0;\red128\green0\blue128;\red128\green0\blue0;\red128\green128\blue0;\red128\green128\blue128;" & vbCrLf
''            sRtf = sRtf & "\red192\green192\blue192;}{\stylesheet{\widctlpar\adjustright \fs20\lang2057\cgrid \snext0 Normal;}{\s1\sb240\sa60\keepn\widctlpar\adjustright \b\f15\fs28\lang2057\kerning28\cgrid \sbasedon0 \snext0 heading 1;}{\s3\sb240\sa60\keepn\widctlpar\adjustright" & vbCrLf
''            sRtf = sRtf & "\b\f15\lang2057\cgrid \sbasedon0 \snext0 heading 3;}{\*\cs10 \additive Default Paragraph Font;}{\s15\qc\widctlpar\adjustright \b\f15\fs16\lang2057\cgrid \sbasedon0 \snext0 caption;}{\s16\li720\widctlpar\adjustright \f2\fs16\lang2057\cgrid" & vbCrLf
''            sRtf = sRtf & "\sbasedon0 \snext16 Code;}{\*\cs17 \additive \ul\cf12 \sbasedon10 FollowedHyperlink;}{\*\cs18 \additive \ul\cf2 \sbasedon10 Hyperlink;}{\s19\widctlpar\adjustright \f15\fs20\lang2057\cgrid \sbasedon0 \snext19 Paragraph;}}{\info" & vbCrLf
''            sRtf = sRtf & "{\title " & sTypeLibName & "Interface Definition}{\author ActiveX Documenter}{\operator ActiveX Documenter}{\creatim\" & sDateString & "}{\revtim\" & sDateString & "}{\printim\" & sDateString & "}{\version1}{\edmins8}" & vbCrLf
''            sRtf = sRtf & "{\nofchars1789}{\*\company vbaccelerator}{\nofcharsws2197}{\vern89}}\paperw11906\paperh16838 \widowctrl\ftnbj\aenddoc\formshade\viewkind1\viewscale100\pgbrdrhead\pgbrdrfoot \fet0\sectd \linex0\headery709\footery709\colsx709\endnhere\sectdefaultcl {\*\pnseclvl1" & vbCrLf
''            sRtf = sRtf & "\pnucrm\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl2\pnucltr\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl3\pndec\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl4\pnlcltr\pnstart1\pnindent720\pnhang{\pntxta )}}{\*\pnseclvl5" & vbCrLf
''            sRtf = sRtf & "\pndec\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl6\pnlcltr\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl7\pnlcrm\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl8\pnlcltr\pnstart1\pnindent720\pnhang" & vbCrLf
''            sRtf = sRtf & "{\pntxtb (}{\pntxta )}}{\*\pnseclvl9\pnlcrm\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}\pard\plain \widctlpar\adjustright \fs20\lang2057\cgrid {\b\f1\fs24" & vbCrLf
''
''            sRtf = sRtf & sTypeLibString & " Interface Definition \par " & vbCrLf & "\par }"
''
''            sRtf = sRtf & "{\b\f1 General Information" & vbCrLf & "\par }"
''            sRtf = sRtf & "\pard \widctlpar\tx993\adjustright {\f1 " & psGetGeneralInfoRtf(0) & vbCrLf
''            sRtf = sRtf & "\par " & psGetGeneralInfoRtf(1) & vbCrLf
''            sRtf = sRtf & "\par " & psGetGeneralInfoRtf(2) & vbCrLf
''            sRtf = sRtf & "\par " & psGetGeneralInfoRtf(3) & vbCrLf
''            sRtf = sRtf & "\par }\pard \widctlpar\adjustright {" & vbCrLf
''            sRtf = sRtf & "\par }"
''
'            If (iTypeInfo = &HFFFFFFF) Then
'                ' Do all the enums:
''                sRtf = sRtf & "{\b\f1 Enumerations" & vbCrLf
''                sRtf = sRtf & "\par }{\f1 This section lists enumerations exposed by " & sTypeLibString & "." & vbCrLf
''                sRtf = sRtf & "\par }{\f1" & vbCrLf
''
'                Status "Reading enums..."
'               ' ProgressMax = (cboClass.ListCount - 1) * 2
'                For i = 0 To cboClass.ListCount - 1
'                    If (i <> cboClass.ListIndex) Then
'                       ' ProgressValue = i
'                        iTypeInfo = cboClass.ItemData(i)
'                        If (m_cTLI.TypeInfos(iTypeInfo).TypeKindString = "enum") Then
'                            Status "Getting Information for " & m_cTLI.TypeInfos(iTypeInfo).Name & " ..."
'                            pDisplayInterfaces iTypeInfo
'                        End If
'                    End If
'                Next i
'
''                sRtf = sRtf & "}{" & vbCrLf
''                sRtf = sRtf & "\par" & vbCrLf
''                sRtf = sRtf & "\par }{\b\f1 Interfaces}{\b\f1" & vbCrLf
''                sRtf = sRtf & "\par }{\f1 This section lists }{\f1 the Classes exposed by " & sTypeLibString & ".  For each class, the methods and events are listed.}{\f1" & vbCrLf
''                sRtf = sRtf & "\par }{" & vbCrLf
''                sRtf = sRtf & "\par }" & vbCrLf
''
'                ' Do all the interfaces:
'                For i = 0 To cboClass.ListCount - 1
'                    If (i <> cboClass.ListIndex) Then
'                        iTypeInfo = cboClass.ItemData(i)
'                       ' ProgressValue = i + cboClass.ListCount - 1
'                        If (m_cTLI.TypeInfos(iTypeInfo).TypeKindString <> "enum") Then
'                            Status "Getting Information for " & m_cTLI.TypeInfos(iTypeInfo).Name & " ..."
'                            pDisplayInterfaces iTypeInfo
'                          '  sRtf = sRtf & "{\par}" & vbCrLf
'                        End If
'                    End If
'                Next i
'
'            Else
'                Status "Getting Information for " & m_cTLI.TypeInfos(iTypeInfo).Name & " ..."
'               ' sRtf = sRtf & " {\f1 "
'                pDisplayInterfaces iTypeInfo ', sRtf
'               ' sRtf = sRtf & "\par }" & vbCrLf
'            End If
'
'            ' Complete the RTF:
'        '    sRtf = sRtf & "\par }}"
'
'            Status "Displaying the TypeLibrary Document..."
'            ' DIsplay the Rtf:
'          '  rtfDocument.Contents(SF_RTF) = sRtf
'
'            Screen.MousePointer = vbDefault
'            Status "Ready."
'           ' ProgressValue = 0
'        Else
'            Status "No Type Library Information."
'        End If
'    Else
'        Status "No Type Library Information."
'    End If
'
'End Sub


Private Sub pDisplayInterfaces(iTypeInfo As Long)
Dim sJunk As String
Dim sType As String
Dim sGUID As String


Dim iBelongsTo As Long
Dim sParseItem As String
Dim iMember As Long
    iEventCount = 0
    iMemberCount = 0
    pEvaluateMember m_cTLI.TypeInfos(iTypeInfo), sJunk, sGUID, sJunk, sType, sMembers(), sHelp(), iMemberID(), iMemberCount
    
    For iBelongsTo = 1 To m_iCount
        If (m_iBelongsToInterface(iBelongsTo) = iTypeInfo) Then
            If (Left$(m_sInterfaces(iBelongsTo), 2) = "__") Then
                ' events:
                pEvaluateMember m_cTLI.TypeInfos(iBelongsTo), sJunk, sJunk, sJunk, sJunk, sEvents(), sEventHelp(), iEventID(), iEventCount
            Else
                ' methods/properties:
                pEvaluateMember m_cTLI.TypeInfos(iBelongsTo), sJunk, sJunk, sJunk, sJunk, sMembers(), sHelp(), iMemberID(), iMemberCount
            End If
        End If
    Next iBelongsTo

    ' Add the information to the class list:
    If (sType = "enum") Then
        'lstMembers2.AddRow
        'lstMembers.AddItem "Public Enum " & m_sInterfaces(iTypeInfo)
        'lstMembers2.CellText(lstMembers2.Rows, 1) = "Public Enum " & m_sInterfaces(iTypeInfo)
        'lstMembers2.RowTag(lstMembers2.Rows) = &HFFFFFFF
        'lstMembers.ItemData(lstMembers.NewIndex) = &HFFFFFFF
        'sRtf = sRtf & "\par " & "Public Enum " & m_sInterfaces(iTypeInfo) & vbCrLf
        For iMember = 1 To iMemberCount
            lstMembers2.AddRow
        'lstMembers.AddItem "Public Enum " & m_sInterfaces(iTypeInfo)
            lstMembers2.CellText(lstMembers2.Rows, 1) = sMembers(iMember) '"Public Enum " & m_sInterfaces(iTypeInfo)
            lstMembers2.CellIcon(lstMembers2.Rows, 1) = IndexForKey("ENUMITEM")
            lstMembers2.RowTag(lstMembers2.Rows) = iMember '&HFFFFFFF
        
            'lstMembers.AddItem vbTab & sMembers(iMember)
            'lstMembers.ItemData(lstMembers.NewIndex) = iMember
            'sRtf = sRtf & "\par \tab " & sMembers(iMember) & vbCrLf
        Next iMember
        'lstMembers.AddItem "End Enum"
        'lstMembers.ItemData(lstMembers.NewIndex) = &HFFFFFFF
        'sRtf = sRtf & "\par End Enum" & vbCrLf
    Else
'        lstMembers2.AddRow
''        'lstMembers.AddItem "Public Enum " & m_sInterfaces(iTypeInfo)
'        lstMembers2.CellText(lstMembers2.Rows, 1) = m_sInterfaces(iTypeInfo) & " " & sGUID '"Public Enum " & m_sInterfaces(iTypeInfo)
'        lstMembers2.CellIcon(lstMembers2.Rows, 1) = IndexForKey("CLASS")
'        lstMembers2.RowTag(lstMembers2.Rows) = -1
        'lstMembers.AddItem "Class:" & vbTab & m_sInterfaces(iTypeInfo) & " " & sGUID
        'lstMembers.ItemData(lstMembers.NewIndex) = -1
        pParseGUID sGUID
        'sRtf = sRtf & "{\b\f1 " & m_sInterfaces(iTypeInfo) & " " & sGUID & vbCrLf
        'sRtf = sRtf & "\par }{" & vbCrLf
        
        'sRtf = sRtf & "\par }{\f1\ul Methods" & vbCrLf & "}"
        If (iMemberCount > 0) Then
            Dim bProperty As Boolean
            'lstMembers.AddItem "Methods:"
            'lstMembers.ItemData(lstMembers.NewIndex) = -1
            For iMember = 1 To iMemberCount
                
                sParseItem = sMembers(iMember)
                
               ' sRtf = sRtf & "{\b\f1 " & vbCrLf & "\par " & sMembers(iMember) & vbCrLf & "\par }"
                If Trim$(Len(sHelp(iMember))) > 0 Then
                    sParseItem = sParseItem & " '" & sHelp(iMember)
                   ' sRtf = sRtf & "{\f1 " & sHelp(iMember) & "}"
                End If
                lstMembers2.AddRow
        'lstMembers.AddItem "Public Enum " & m_sInterfaces(iTypeInfo)
                
                If Left$(sParseItem, 1) = "^" Then
                    sParseItem = Right$(sParseItem, Len(sParseItem) - 1)
                    lstMembers2.CellIcon(lstMembers2.Rows, 1) = IndexForKey("PROP")
                Else
                    
                    lstMembers2.CellIcon(lstMembers2.Rows, 1) = IndexForKey("METHOD")
                End If
                lstMembers2.CellText(lstMembers2.Rows, 1) = sParseItem
                lstMembers2.RowTag(lstMembers2.Rows) = iMemberID(iMember) '&HFFFFFFF
'
'                lstMembers.AddItem sParseItem
'                lstMembers.ItemData(lstMembers.NewIndex) = iMember
'
            Next iMember
        Else
           ' lstMembers.AddItem "No Methods."
           ' lstMembers.ItemData(lstMembers.NewIndex) = -1
           ' sRtf = sRtf & "{\f1" & "None " & vbCrLf & "\par}"
        End If
        
       ' sRtf = sRtf & "{\f1" & vbCrLf & "\par}{\f1\ul Events" & vbCrLf & "\par}"
        If (iEventCount > 0) Then
'            lstMembers.AddItem "Events:"
'            lstMembers.ItemData(lstMembers.NewIndex) = -1

            For iMember = 1 To iEventCount
                'sParseItem = "Public Event " & Mid$(sEvents(iMember), 5)
                sParseItem = sEvents(iMember)
            
                'sRtf = sRtf & "{\b\f1 " & vbCrLf & "\par " & sParseItem & vbCrLf & "\par }"
                If Trim$(Len(sEventHelp(iMember))) > 0 Then
                    sParseItem = sParseItem & " '" & sEventHelp(iMember)
                    'sRtf = sRtf & "{\f1 " & sEventHelp(iMember) & "}"
                End If
                lstMembers2.AddRow
        'lstMembers.AddItem "Public Enum " & m_sInterfaces(iTypeInfo)
                lstMembers2.CellText(lstMembers2.Rows, 1) = sParseItem
                lstMembers2.CellIcon(lstMembers2.Rows, 1) = IndexForKey("EVENT")
                lstMembers2.RowTag(lstMembers2.Rows) = iEventID(iMember) '&HFFFFFFF
'                lstMembers.AddItem sParseItem
'                lstMembers.ItemData(lstMembers.NewIndex) = iMember
            Next iMember
            
        Else
'            lstMembers.AddItem "No Events."
'            lstMembers.ItemData(lstMembers.NewIndex) = -1
            'sRtf = sRtf & "{\f1" & "None " & vbCrLf & "\par}"
            
        End If
        
    End If

End Sub

Private Sub pEvaluateMember( _
        ByRef tI As TypeInfo, _
        ByRef sName As String, _
        ByRef sGUID As String, _
        ByRef sHelpString As String, _
        ByRef sType As String, _
        ByRef sMembers() As String, _
        ByRef sHelp() As String, _
        ByRef iMemberID() As Long, _
        ByRef iMemberCount As Long _
    )
Dim iTypeInfo As Long

    With tI
        sName = .Name
        sGUID = .Guid
        sHelpString = .HelpString
        sType = .TypeKindString
        
        If (.TypeKind = TKIND_ENUM) Then
            ' do enum:
            pEvaluateEnum tI, sMembers(), iMemberCount
        Else
            ' do class:
            pEvaluateClass tI, sMembers(), sHelp(), iMemberID(), iMemberCount
        End If
    End With
    
End Sub

Private Sub pEvaluateEnum( _
        ByRef tI As TypeInfo, _
        ByRef sMembers() As String, _
        ByRef iMemberCount As Long _
    )
Dim iMember As Long

    iMemberCount = 0
    Erase sMembers
    
    With tI
        On Error Resume Next
        iMemberCount = .Members.Count
        If (Err.Number <> 0) Then
            iMemberCount = 0
        End If
        Err.Clear
        
        On Error GoTo 0
        If (iMemberCount > 0) Then
            ReDim sMembers(1 To iMemberCount) As String
            For iMember = 1 To iMemberCount
                With .Members(iMember)
                    sMembers(iMember) = .Name & "=" & .Value
                End With
            Next iMember
        End If
    End With
    
End Sub
Private Sub pEvaluateClass( _
        ByRef tI As TypeInfo, _
        ByRef sMembers() As String, _
        ByRef sHelp() As String, _
        ByRef iMemberID() As Long, _
        ByRef iMemberCount As Long _
    )
Dim iMember As Long
Dim iMemCount As Long
        
    ' Initialise:
    iMemberCount = 0
    Erase sMembers
    Erase sHelp
    Erase iMemberID
    ' Find out the contents of the TypeInfo:
    With tI
        
        ' Get number of members in this class:
        On Error Resume Next
        iMemCount = .Members.Count
        If (Err.Number <> 0) Then
            iMemCount = 0
        End If
        Err.Clear
        
        On Error GoTo 0
        If (iMemCount > 0) Then
        
            For iMember = 1 To iMemCount
                If (.Members(iMember).AttributeMask = 0) Then ' Not hidden
                    iMemberCount = iMemberCount + 1
                    ReDim Preserve sMembers(1 To iMemberCount) As String
                    ReDim Preserve sHelp(1 To iMemberCount) As String
                    ReDim Preserve iMemberID(1 To iMemberCount) As Long
                    iMemberID(iMemberCount) = iMember
                    pEvaluateClassMember .Members(iMember), sMembers(iMemberCount), iMemberCount
                End If
            Next iMember
        End If
        
    End With
End Sub

Private Sub pParseGUID(ByRef sThis As String)
    ReplaceSection sThis, "{", "\{"
    ReplaceSection sThis, "}", "\}"
End Sub

Private Sub ReplaceSection( _
        ByRef sToModify As String, _
        ByVal sToReplace As String, _
        ByVal sReplaceWith As String _
    )
' ==================================================================
' Replaces all occurrences of sToReplace with
' sReplaceWidth in sToModify.
' ==================================================================
' ==================================================================
' Replaces all occurrences of sToReplace with
' sReplaceWidth in sToModify.
' ==================================================================
Dim iLastPos As Long
Dim iNextPos As Long
Dim iReplaceLen As Long
Dim sOut As String
    iReplaceLen = Len(sToReplace)
    iLastPos = 1
    iNextPos = InStr(iLastPos, sToModify, sToReplace)
    sOut = ""
    Do While (iNextPos > 0)
        If (iNextPos > 1) Then
            sOut = sOut & Mid$(sToModify, iLastPos, (iNextPos - iLastPos))
        End If
        sOut = sOut & sReplaceWith
        iLastPos = iNextPos + iReplaceLen
        iNextPos = InStr(iLastPos, sToModify, sToReplace)
    Loop
    If (iLastPos <= Len(sToModify)) Then
        sOut = sOut & Mid$(sToModify, (iLastPos))
    End If
    sToModify = sOut
End Sub

Private Sub pEvaluateClassMember( _
        ByRef tM As MemberInfo, _
        ByRef sMember As String, _
        ByRef iMember As Long _
    )
Dim iParam As Long
Dim iParamCount As Long
Dim lType As TliVarType
Dim bOptional As Boolean
Dim bIsLet As Boolean
Dim sDefault As String
Dim sName As String
Dim sPrefix As String

'On Error Resume Next

    With tM
        
        ' Type of member (sub, function, property..):
        sMember = psGetMemberType(tM, bIsLet)
        sName = .Name
        If (Left$(sName, 1) = "_") Then
            ' check for standard prefixes:
            sPrefix = Left$(sName, 7)
            If (sPrefix = "_B_var_") Then
                sName = Mid$(sName, 8)
            ElseIf (sPrefix = "_B_str_") Then
                sName = Mid$(sName, 8) & "$"
            End If
        End If
        'If bDetailed = False Then
            If Left$(sMember, 8) = "Property" Then
                sMember = "^" & sName & " (" & IIf(bIsLet, "Let", "Get") & ")"
            Else
                sMember = sName
            End If
            'Exit Sub
        'End If
'        sMember = sMember & " " & sName
'
'
'        ' Any parameters?
'        iParamCount = .Parameters.Count
'        If (Err.Number <> 0) Then
'            iParamCount = 0
'        End If
'        Err.Clear
'
'        ' If we have parameters then add the function description:
'        For iParam = 1 To iParamCount
'
'            bOptional = False
'
'            With .Parameters(iParam)
'                ' Add open bracket first time:
'                If (iParam = 1) Then
'                    sMember = sMember & "("
'                End If
'
'                ' .HasCustomData or .Optional implies the parameter is optional:
'                If (.HasCustomData() = True) Then
'                   sMember = sMember & "Optional "
'                   bOptional = True
'                Else
'                    If .Optional Then
'                        sMember = sMember & "Optional "
'                    End If
'                End If
'
'                ' Check Byref/Byval status of member:
'                If ((lType And VT_BYREF) = VT_BYREF) Then
'                Else
'                    sMember = sMember & "ByVal "
'                End If
'
'                ' Name of parameter:
'                sMember = sMember & .Name
'
'                ' Evaluate the parameter type:
'                If (.VarTypeInfo.VarType = 0) Then
'                    ' Custom type:
'                    sMember = sMember & " As " & .VarTypeInfo.TypeInfo.Name
'                Else
'                    lType = .VarTypeInfo.VarType
'                    sMember = sMember & psTranslateType(lType)
'                End If
'
'                ' Add default value if there is one:
'                If (bOptional) Then
'                   If (.Default) Then
'                    On Error Resume Next
'                    sDefault = CStr(.DefaultValue)
'                    If (Err.Number = 0) Then
'                        sMember = sMember & "=" & sDefault
'                    Else
'                        sMember = sMember & "=Nothing"
'                    End If
'                    Err.Clear
'                    On Error GoTo 0
'                   End If
'                End If
'
'                ' If this is the last parameter then close the declaration,
'                ' otherwise put a comma in front of the next one:
'                If (iParam < iParamCount) Then
'                    sMember = sMember & ", "
'                Else
'                    If Not (bIsLet) Then
'                        sMember = sMember & ")"
'                    End If
'                End If
'
'            End With
'        Next iParam
'
'        ' Now add the return type and fix up Property Lets as required:
'        If (.ReturnType.VarType <> 0) Then
'            ' Returns a Standard type:
'            If (.ReturnType.VarType = VT_VOID) Then
'                ' sub
'            Else
'                ' If a constant, we want to get the constant value:
'                If (Left$(sMember, 5) = "Const") Then
'                    Debug.Print sMember, .ReturnType.VarType
'                    On Error Resume Next
'                    If (.ReturnType.VarType = VT_BSTR) Or (.ReturnType.VarType = VT_LPSTR) Then
'                        sMember = sMember & " = " & psParseForNonPrintable(.Value) & ""
'                    Else
'                        sMember = sMember & " = " & .Value
'                    End If
'                Else
'                    ' If property let, must put in the RHS argument:
'                    If (bIsLet) Then
'                        If (iParamCount = 0) Then
'                            ' property let has only one var:
'                            sMember = sMember & "(RHS "
'                        Else
'                            ' more than on property let var:
'                            sMember = sMember & ", RHS"
'                        End If
'                    End If
'
'                    ' No paramters, put the open close in:
'                    If (iParamCount = 0) Then
'                        If Not (bIsLet) Then
'                            sMember = sMember & "()"
'                        End If
'                    End If
'
'                    ' Add the return type:
'                    sMember = sMember & psTranslateType(.ReturnType.VarType)
'
'                    ' Close the property let statement:
'                    If (bIsLet) Then
'                        sMember = sMember & ")"
'                    End If
'                End If
'            End If
'        Else
'            ' If property let, must put in the RHS argument:
'            If (bIsLet) Then
'                If (iParamCount = 0) Then
'                    ' property let has only one var:
'                    sMember = sMember & "(RHS "
'                Else
'                    ' more than on property let var:
'                    sMember = sMember & ", RHS"
'                End If
'            End If
'
'            ' No paramters, put the open close in:
'            If (iParamCount = 0) Then
'                If Not (bIsLet) Then
'                    sMember = sMember & "()"
'                End If
'            End If
'
'            ' Returns a custom type:
'            sMember = sMember & " As " & .ReturnType.TypeInfo.Name
'
'            ' Close the property let statement:
'            If (bIsLet) Then
'                sMember = sMember & ")"
'            End If
'        End If
'
'        sHelp = .HelpString
    
    End With
    
End Sub
Private Sub pDisplayClassMemberInfo(ByRef tM As MemberInfo)
        
Dim iParam As Long
Dim iParamCount As Long
Dim lType As TliVarType
Dim bOptional As Boolean
Dim bIsLet As Boolean
Dim sDefault As String
Dim sName As String
Dim sPrefix As String
Dim sMember As String
'On Error Resume Next
    With tM
        
        ' Type of member (sub, function, property..):
        sMember = psGetMemberType(tM, bIsLet)
        sName = .Name
        If (Left$(sName, 1) = "_") Then
            ' check for standard prefixes:
            sPrefix = Left$(sName, 7)
            If (sPrefix = "_B_var_") Then
                sName = Mid$(sName, 8)
            ElseIf (sPrefix = "_B_str_") Then
                sName = Mid$(sName, 8) & "$"
            End If
        End If
        
        sMember = sMember & " " & sName
        
        
        ' Any parameters?
        iParamCount = .Parameters.Count
        If (Err.Number <> 0) Then
            iParamCount = 0
        End If
        Err.Clear
        
        ' If we have parameters then add the function description:
        For iParam = 1 To iParamCount
            
            bOptional = False
            
            With .Parameters(iParam)
                ' Add open bracket first time:
                If (iParam = 1) Then
                    sMember = sMember & "("
                End If
                
                ' .HasCustomData or .Optional implies the parameter is optional:
                If (.HasCustomData() = True) Then
                   sMember = sMember & "Optional "
                   bOptional = True
                Else
                    If .Optional Then
                        sMember = sMember & "Optional "
                    End If
                End If
                
                ' Check Byref/Byval status of member:
                If ((lType And VT_BYREF) = VT_BYREF) Then
                Else
                    sMember = sMember & "ByVal "
                End If
                
                ' Name of parameter:
                sMember = sMember & .Name
                
                ' Evaluate the parameter type:
                If (.VarTypeInfo.VarType = 0) Then
                    ' Custom type:
                    sMember = sMember & " As " & .VarTypeInfo.TypeInfo.Name
                Else
                    lType = .VarTypeInfo.VarType
                    sMember = sMember & psTranslateType(lType)
                End If
                
                ' Add default value if there is one:
                If (bOptional) Then
                   If (.Default) Then
                    On Error Resume Next
                    sDefault = CStr(.DefaultValue)
                    If (Err.Number = 0) Then
                        sMember = sMember & "=" & sDefault
                    Else
                        sMember = sMember & "=Nothing"
                    End If
                    Err.Clear
                    On Error GoTo 0
                   End If
                End If
                
                ' If this is the last parameter then close the declaration,
                ' otherwise put a comma in front of the next one:
                If (iParam < iParamCount) Then
                    sMember = sMember & ", "
                Else
                    If Not (bIsLet) Then
                        sMember = sMember & ")"
                    End If
                End If
                
            End With
        Next iParam
                    
        ' Now add the return type and fix up Property Lets as required:
        If (.ReturnType.VarType <> 0) Then
            ' Returns a Standard type:
            If (.ReturnType.VarType = VT_VOID) Then
                ' sub
            Else
                ' If a constant, we want to get the constant value:
                If (Left$(sMember, 5) = "Const") Then
                    Debug.Print sMember, .ReturnType.VarType
                    On Error Resume Next
                    If (.ReturnType.VarType = VT_BSTR) Or (.ReturnType.VarType = VT_LPSTR) Then
                        sMember = sMember & " = " & psParseForNonPrintable(.Value) & ""
                    Else
                        sMember = sMember & " = " & .Value
                    End If
                Else
                    ' If property let, must put in the RHS argument:
                    If (bIsLet) Then
                        If (iParamCount = 0) Then
                            ' property let has only one var:
                            sMember = sMember & "(RHS "
                        Else
                            ' more than on property let var:
                            sMember = sMember & ", RHS"
                        End If
                    End If
                    
                    ' No paramters, put the open close in:
                    If (iParamCount = 0) Then
                        If Not (bIsLet) Then
                            sMember = sMember & "()"
                        End If
                    End If
                    
                    ' Add the return type:
                    sMember = sMember & psTranslateType(.ReturnType.VarType)
                    
                    ' Close the property let statement:
                    If (bIsLet) Then
                        sMember = sMember & ")"
                    End If
                End If
            End If
        Else
            ' If property let, must put in the RHS argument:
            If (bIsLet) Then
                If (iParamCount = 0) Then
                    ' property let has only one var:
                    sMember = sMember & "(RHS "
                Else
                    ' more than on property let var:
                    sMember = sMember & ", RHS"
                End If
            End If
            
            ' No paramters, put the open close in:
            If (iParamCount = 0) Then
                If Not (bIsLet) Then
                    sMember = sMember & "()"
                End If
            End If
            
            ' Returns a custom type:
            sMember = sMember & " As " & .ReturnType.TypeInfo.Name
            
            ' Close the property let statement:
            If (bIsLet) Then
                sMember = sMember & ")"
            End If
        End If
        rtfInfo.Text = sMember & vbCrLf & .HelpString
        rtfInfo.SelStart = InStr(1, sMember, sName) - 1
        rtfInfo.SelLength = Len(sName)
        rtfInfo.SelBold = True
        rtfInfo.SelStart = 0
       ' sHelp = .HelpString
    
    End With
    
End Sub
Private Function psGetMemberType( _
        ByRef tM As MemberInfo, _
        ByRef bIsLet As Boolean _
    ) As String
    
    bIsLet = False
    
    Select Case tM.InvokeKind
    Case INVOKE_EVENTFUNC
        psGetMemberType = "Event"
    Case INVOKE_FUNC
        If (tM.ReturnType.VarType = VT_VOID) Then
            psGetMemberType = "Sub"
        Else
            psGetMemberType = "Function"
        End If
    Case INVOKE_PROPERTYGET
        psGetMemberType = "Property Get"
    Case INVOKE_PROPERTYPUT
        psGetMemberType = "Property Let"
        bIsLet = True
    Case INVOKE_PROPERTYPUTREF
        psGetMemberType = "Property Set"
    Case INVOKE_UNKNOWN
        psGetMemberType = "Const"
    Case Else
        Debug.Assert 1 = 0
    End Select
    
End Function

Private Function psTranslateType(ByVal lType As Long)
Dim sType As String
    Select Case (lType And &HFF&)
    Case VT_BOOL
        sType = "Boolean"
    Case VT_BSTR, VT_LPSTR
        sType = "String"
    Case VT_DATE
        sType = "Date"
    Case VT_INT
        sType = "Integer"
    Case VT_VARIANT
        sType = "Variant"
    Case VT_DECIMAL
        sType = "Decimal"
    Case VT_I4
        sType = "Long"
    Case VT_I2
        sType = "Integer"
    Case VT_I8
        sType = "Unknown"
    Case VT_SAFEARRAY
        sType = "SafeArray"
    Case VT_CLSID
        sType = "CLSID"
    Case VT_UINT
        sType = "UInt"
    Case VT_UI4
        sType = "ULong"
    Case VT_UNKNOWN
        sType = "Unknown"
    Case VT_VECTOR
        sType = "Vector"
    Case VT_R4
        sType = "Single"
    Case VT_R8
        sType = "Double"
    Case VT_DISPATCH
        sType = "Object"
    Case VT_UI1
        sType = "Byte"
    Case VT_CY
        sType = "Currency"
    Case Else
        'sType = "???"
     '   Debug.Assert 1 = 0
    End Select
    If (lType And VT_ARRAY) = VT_ARRAY Then
        sType = "() As " & sType
    ElseIf sType <> "" Then
        sType = " As " & sType
    End If
    psTranslateType = sType

End Function

Private Function psParseForNonPrintable(ByVal vThis As Variant) As String
Dim iPos As Long
Dim sRet As String
Dim iLen As Long
Dim iChar As Integer
Dim sChar As String
Dim bLastNonPrintable As Boolean

    iLen = Len(vThis)
    For iPos = 1 To iLen
        sChar = Mid$(vThis, iPos, 1)
        iChar = Asc(sChar)
        If (iChar < 32) Then
            If (iPos <> 1) Then
                sRet = sRet & "& "
            End If
            If (bLastNonPrintable) Then
                sChar = "Chr$(" & iChar & ") "
            Else
                If (iPos = 1) Then
                    sChar = "Chr$(" & iChar & ") "
                Else
                    sChar = """ & Chr$(" & iChar & ") "
                End If
            End If
            bLastNonPrintable = True
        Else
            If (bLastNonPrintable) Or (iPos = 1) Then
                If (iPos <> 1) Then
                    sRet = sRet & "& "
                End If
                sChar = """" & sChar
            End If
            bLastNonPrintable = False
        End If
        sRet = sRet & sChar
    Next iPos
    If Not (bLastNonPrintable) Then
        sRet = sRet & """"
    End If
    psParseForNonPrintable = sRet

End Function

Private Sub cmdSelect_Click()
    If cFunc.Dialogs.ShowOpenSaveDialog(False, "Select Component...", "ActiveX Files (*.OCX;*.DLL;*.TLB;*.OLB;*.EXE)|*.OCX;*.DLL;*.TLB;*.OLB;*.EXE|ActiveX Controls (*.OCX)|*.OCX|ActiveX DLLs (*.DLL)|*.DLL|Type Libraries (*.TLB;*.OLB)|*.TLB;*.OLB|ActiveX Executables (*.EXE)|*.EXE|All Files (*.*)|*.", , Me) = True Then
        txtFile.Text = cFunc.Dialogs.FileName
        txtFile.SelStart = Len(txtFile.Text)
        pbGetTypeLibInfo txtFile.Text
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then ShowHTMLHelpTopic 3, hWnd
End Sub

Private Sub Form_Load()
    lstClass.ImageList = imlIcons.hIml
    lstClass.AddColumn
    lstMembers2.ImageList = imlIcons.hIml
    lstMembers2.AddColumn
    lstMembers2.SortObject.SortColumn(1) = 1
    lstMembers2.SortObject.SortType(1) = CCLSortString
End Sub

Private Sub lstClass_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
    Dim iTypeInfo As Long
Dim i As Long
Dim sRtf As String
Dim sTypeLibName As String
Dim sDateString As String
Dim sTypeLibString As String

    ' Clear list
    Status "Getting Type Library Information..."
    lstMembers2.Clear
    lstMembers2.Redraw = False
    If (lRow > -1) Then
        ' Evaluate the contents:
        iTypeInfo = CLng(lstClass.RowTag(lRow))
        'iTypeInfo = cboClass.ItemData(cboClass.ListIndex)
        
        If (iTypeInfo > 0) Then
                
            Screen.MousePointer = vbHourglass
        
            ' Prepare the RTF header:
            sDateString = "yr" & Year(Now) & "\mo" & Month(Now) & "\dy" & Day(Now) & "\hr" & Hour(Now) & "\min" & Minute(Now)
            sTypeLibString = m_cTLI.Name
            
'            sRtf = "{\rtf1\ansi\ansicpg1252\uc1 \deff0\deflang1033\deflangfe1033{\fonttbl{\f0\froman\fcharset0\fprq2{\*\panose 02020603050405020304}Times New Roman;}{\f1\fswiss\fcharset0\fprq2{\*\panose 020b0604020202020204}Arial;}" & vbCrLf
'            sRtf = sRtf & "{\f2\fmodern\fcharset0\fprq1{\*\panose 02070309020205020404}Courier New;}{\f15\fswiss\fcharset0\fprq2{\*\panose 020b0604030504040204}Verdana;}}{\colortbl;\red0\green0\blue0;\red0\green0\blue255;\red0\green255\blue255;\red0\green255\blue0;" & vbCrLf
'            sRtf = sRtf & "\red255\green0\blue255;\red255\green0\blue0;\red255\green255\blue0;\red255\green255\blue255;\red0\green0\blue128;\red0\green128\blue128;\red0\green128\blue0;\red128\green0\blue128;\red128\green0\blue0;\red128\green128\blue0;\red128\green128\blue128;" & vbCrLf
'            sRtf = sRtf & "\red192\green192\blue192;}{\stylesheet{\widctlpar\adjustright \fs20\lang2057\cgrid \snext0 Normal;}{\s1\sb240\sa60\keepn\widctlpar\adjustright \b\f15\fs28\lang2057\kerning28\cgrid \sbasedon0 \snext0 heading 1;}{\s3\sb240\sa60\keepn\widctlpar\adjustright" & vbCrLf
'            sRtf = sRtf & "\b\f15\lang2057\cgrid \sbasedon0 \snext0 heading 3;}{\*\cs10 \additive Default Paragraph Font;}{\s15\qc\widctlpar\adjustright \b\f15\fs16\lang2057\cgrid \sbasedon0 \snext0 caption;}{\s16\li720\widctlpar\adjustright \f2\fs16\lang2057\cgrid" & vbCrLf
'            sRtf = sRtf & "\sbasedon0 \snext16 Code;}{\*\cs17 \additive \ul\cf12 \sbasedon10 FollowedHyperlink;}{\*\cs18 \additive \ul\cf2 \sbasedon10 Hyperlink;}{\s19\widctlpar\adjustright \f15\fs20\lang2057\cgrid \sbasedon0 \snext19 Paragraph;}}{\info" & vbCrLf
'            sRtf = sRtf & "{\title " & sTypeLibName & "Interface Definition}{\author ActiveX Documenter}{\operator ActiveX Documenter}{\creatim\" & sDateString & "}{\revtim\" & sDateString & "}{\printim\" & sDateString & "}{\version1}{\edmins8}" & vbCrLf
'            sRtf = sRtf & "{\nofchars1789}{\*\company vbaccelerator}{\nofcharsws2197}{\vern89}}\paperw11906\paperh16838 \widowctrl\ftnbj\aenddoc\formshade\viewkind1\viewscale100\pgbrdrhead\pgbrdrfoot \fet0\sectd \linex0\headery709\footery709\colsx709\endnhere\sectdefaultcl {\*\pnseclvl1" & vbCrLf
'            sRtf = sRtf & "\pnucrm\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl2\pnucltr\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl3\pndec\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl4\pnlcltr\pnstart1\pnindent720\pnhang{\pntxta )}}{\*\pnseclvl5" & vbCrLf
'            sRtf = sRtf & "\pndec\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl6\pnlcltr\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl7\pnlcrm\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl8\pnlcltr\pnstart1\pnindent720\pnhang" & vbCrLf
'            sRtf = sRtf & "{\pntxtb (}{\pntxta )}}{\*\pnseclvl9\pnlcrm\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}\pard\plain \widctlpar\adjustright \fs20\lang2057\cgrid {\b\f1\fs24" & vbCrLf
'
'            sRtf = sRtf & sTypeLibString & " Interface Definition \par " & vbCrLf & "\par }"
'
'            sRtf = sRtf & "{\b\f1 General Information" & vbCrLf & "\par }"
'            sRtf = sRtf & "\pard \widctlpar\tx993\adjustright {\f1 " & psGetGeneralInfoRtf(0) & vbCrLf
'            sRtf = sRtf & "\par " & psGetGeneralInfoRtf(1) & vbCrLf
'            sRtf = sRtf & "\par " & psGetGeneralInfoRtf(2) & vbCrLf
'            sRtf = sRtf & "\par " & psGetGeneralInfoRtf(3) & vbCrLf
'            sRtf = sRtf & "\par }\pard \widctlpar\adjustright {" & vbCrLf
'            sRtf = sRtf & "\par }"
'
'            If (iTypeInfo = &HFFFFFFF) Then
'                ' Do all the enums:
''                sRtf = sRtf & "{\b\f1 Enumerations" & vbCrLf
''                sRtf = sRtf & "\par }{\f1 This section lists enumerations exposed by " & sTypeLibString & "." & vbCrLf
''                sRtf = sRtf & "\par }{\f1" & vbCrLf
''
'                Status "Reading enums..."
'               ' ProgressMax = (cboClass.ListCount - 1) * 2
'                'For i = 0 To cboClass.ListCount - 1
'                For i = 1 To lstClass.Rows
'                    If (i <> lRow) Then
'                       ' ProgressValue = i
'                        iTypeInfo = CLng(lstClass.RowTag(i)) 'cboClass.ItemData(i)
'                        If (m_cTLI.TypeInfos(iTypeInfo).TypeKindString = "enum") Then
'                            Status "Getting Information for " & m_cTLI.TypeInfos(iTypeInfo).Name & " ..."
'                            pDisplayInterfaces iTypeInfo
'                        End If
'                    End If
'                Next i
'
''                sRtf = sRtf & "}{" & vbCrLf
''                sRtf = sRtf & "\par" & vbCrLf
''                sRtf = sRtf & "\par }{\b\f1 Interfaces}{\b\f1" & vbCrLf
''                sRtf = sRtf & "\par }{\f1 This section lists }{\f1 the Classes exposed by " & sTypeLibString & ".  For each class, the methods and events are listed.}{\f1" & vbCrLf
''                sRtf = sRtf & "\par }{" & vbCrLf
''                sRtf = sRtf & "\par }" & vbCrLf
''
'                ' Do all the interfaces:
'                'For i = 0 To cboClass.ListCount - 1
'                For i = 1 To lstClass.Rows
'                    'If (i <> cboClass.ListIndex) Then
'                    If (i <> lRow) Then
'                        iTypeInfo = CLng(lstClass.RowTag(i)) 'cboClass.ItemData(i)
'                       ' ProgressValue = i + cboClass.ListCount - 1
'                        If (m_cTLI.TypeInfos(iTypeInfo).TypeKindString <> "enum") Then
'                            Status "Getting Information for " & m_cTLI.TypeInfos(iTypeInfo).Name & " ..."
'                            pDisplayInterfaces iTypeInfo
'                          '  sRtf = sRtf & "{\par}" & vbCrLf
'                        End If
'                    End If
'                Next i
'
'            Else
                Status "Getting Information for " & m_cTLI.TypeInfos(iTypeInfo).Name & " ..."
               ' sRtf = sRtf & " {\f1 "
                pDisplayInterfaces iTypeInfo ', sRtf
               ' sRtf = sRtf & "\par }" & vbCrLf
          '  End If
            
            ' Complete the RTF:
        '    sRtf = sRtf & "\par }}"
                
           ' Status "Displaying the TypeLibrary Document..."
            ' DIsplay the Rtf:
          '  rtfDocument.Contents(SF_RTF) = sRtf

            Screen.MousePointer = vbDefault
            Status "Ready."
           ' ProgressValue = 0
        Else
            Status "No Type Library Information."
        End If
    Else
        Status "No Type Library Information."
    End If
    lstMembers2.Sort
    lstMembers2.Redraw = True
    
    rtfInfo.Text = ""
    lstMembers2.AutoWidthColumn 1
    If lstMembers2.Rows > 0 Then
        lstMembers2.SelectedRow = 1
        lstMembers2_SelectionChange 1, 1
    End If
    
End Sub
Private Function IndexForKey(sKey As String) As Long
    If sKey = "CLASS" Then
        IndexForKey = 0
    Else
        IndexForKey = imlIcons.ItemIndex(UCase$(sKey))
    End If
End Function

Private Sub lstMembers2_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
Dim iTypeInfo As Long
Dim iMember As Long
Dim iBelongsTo As Long
Dim bFound As Boolean
    iTypeInfo = CLng(lstClass.RowTag(lstClass.SelectedRow))
    If lstMembers2.RowTag(lRow) = "" Then Exit Sub
    iMember = CLng(lstMembers2.RowTag(lRow))
    
    For iBelongsTo = 1 To m_iCount
        
        If (m_iBelongsToInterface(iBelongsTo) = iTypeInfo) Then
            If (Left$(m_sInterfaces(iBelongsTo), 2) = "__") Then
                ' events:
                'pEvaluateMember m_cTLI.TypeInfos(iBelongsTo), sJunk, sJunk, sJunk, sJunk, sEvents(), sEventHelp(), iEventID(), iEventCount
                
                If iMember <= iEventCount Then
                    pDisplayClassMemberInfo m_cTLI.TypeInfos(iBelongsTo).Members(iMember)
                    bFound = True
                End If
            Else
                ' methods/properties:
                pDisplayClassMemberInfo m_cTLI.TypeInfos(iBelongsTo).Members(iMember)
                bFound = True
            '    Exit For
                'pEvaluateMember m_cTLI.TypeInfos(iBelongsTo), sJunk, sJunk, sJunk, sJunk, sMembers(), sHelp(), iMemberID(), iMemberCount
            End If
        End If
    Next iBelongsTo
    If bFound = False Then
        'enum
        pDisplayClassMemberInfo m_cTLI.TypeInfos(iTypeInfo).Members(iMember)
    End If
End Sub


Private Sub picHolder_Resize()
    lblLabel(0).Left = picHolder.ScaleWidth - lblLabel(0).Width
End Sub
