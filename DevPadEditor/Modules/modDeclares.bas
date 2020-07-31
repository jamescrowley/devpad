Attribute VB_Name = "modDeclares"
Option Explicit
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

Public Function TranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then TranslateColor = -1
End Function
Public Function StripChar(ByVal sChar As String, ByVal sString As String) As String
    If Left$(sString, Len(sChar)) = sChar Then sString = Right$(sString, Len(sString) - Len(sChar))
    If Right$(sString, Len(sChar)) = sChar Then sString = Left$(sString, Len(sString) - Len(sChar))
    StripChar = sString
End Function
