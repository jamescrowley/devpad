VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ICaptureBF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get CapturehWnd() As Long
Attribute CapturehWnd.VB_Description = "Requests the window handle of the window you will use to capture the folder browse dialog into.  You must show your form before returning the window handle."

End Property
Public Sub SelectionChanged(ByVal sPath As String)
Attribute SelectionChanged.VB_Description = "Fires when the selection in the folder browse dialog changes."

End Sub
Public Property Let CaptureBrowseForFolder(ByRef c As Object)
Attribute CaptureBrowseForFolder.VB_Description = "Provides you with a reference to the browse for folder object being used in the dialog.  If you copy a reference to this object, make sure you set it to nothing as your form unloads."
   
End Property
Public Sub Unload()
Attribute Unload.VB_Description = "Fired when the browse for folder dialog is closed.  Unload your form at this time to ensure everything clears up."

End Sub
