VERSION 5.00
Object = "{017E002E-D7CC-11D2-8E21-44B10AC10000}#22.1#0"; "vbwGrid.ocx"
Begin VB.Form frmTaskList 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Task List - 3 tasks"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin vbAcceleratorGrid.vbalGrid vbalGrid1 
      Height          =   2070
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   7005
      _extentx        =   12356
      _extenty        =   3651
      font            =   "frmTaskList.frx":0000
      borderstyle     =   0
      disableicons    =   -1  'True
      defaultrowheight=   15
   End
End
Attribute VB_Name = "frmTaskList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim i As Long
    With vbalGrid1
        .BorderStyle = ecgBorderStyle3dThin
        .GridLines = True
        .Editable = True
        .HeaderButtons = False
        .DefaultRowHeight = 17
        .RowMode = True
        .AddColumn "Importance", "!", ecgHdrTextALignCentre, -1, 10, True, True
        .AddColumn "Icon", "", ecgHdrTextALignCentre, -1, 20, True, True
        .AddColumn "Done", "", ecgHdrTextALignCentre, -1, 20, True, False
        .AddColumn "Description", "Description", ecgHdrTextALignLeft, , (vbalGrid1.Width / Screen.TwipsPerPixelX) - 52
        .AddColumn "File", "File", ecgHdrTextALignLeft, , 200
        .AddColumn "Line", "Line", ecgHdrTextALignLeft, , 100
        
        .AddRow -1, "Task1", True
        .CellText(.rows, 4) = "Click here to add a new task"
        .CellForeColor(.rows, 4) = &H80000011
        '.CellTextAlign(.rows, 4) = DT_VCENTER
        .AddRow -1, "AddRowSep", True, 2
        For i = 1 To .Columns
            .CellBackColor(.rows, i) = &H80000011
        Next

        .AddRow -1, "AddRow", True, -1
        .CellText(.rows, 4) = "My Task"
    End With
End Sub

Private Sub vbalGrid1_ColumnWidthStartChange(ByVal lCol As Long, ByVal lWidth As Long, bCancel As Boolean)
    'don't allow first 3 columns to be resized...
    If lCol < 4 Then bCancel = True
End Sub

Private Sub vbalGrid1_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
    If KeyCode = vbKeyDown And vbalGrid1.selectedrow = 1 Then
        bDoDefault = False
        vbalGrid1.selectedrow = 3
    End If
End Sub

Private Sub vbalGrid1_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
bCancel = False
End Sub

Private Sub vbalGrid1_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
    If lRow = 2 Then vbalGrid1.selectedrow = 1
End Sub
