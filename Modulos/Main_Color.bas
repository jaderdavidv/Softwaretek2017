Attribute VB_Name = "Main_Color"
Option Compare Database
Option Explicit
 
Declare PtrSafe Sub wlib_AccColorDialog _
  Lib "msaccess.exe" _
    Alias "#53" (ByVal Hwnd As Long, lngRGB As Long)
 
Public Function ChooseWebColor(DefaultWebColor As Variant) As String
  Dim lngColor As Long
  lngColor = CLng("&H" & Right("000000" + _
                  Replace(Nz(DefaultWebColor, ""), "#", ""), 6))
  wlib_AccColorDialog Screen.ActiveForm.Hwnd, lngColor
  ChooseWebColor = "#" & Right("000000" & Hex(lngColor), 6)
End Function

Function HexToLong(ByVal sHex As String) As Long
        HexToLong = val("&H" & sHex & "&")
End Function

Public Function obtenerColorEnfoque(focus As Boolean) As Long

If (focus) Then
    Dim Color As String
    Color = "#F3A647"
    obtenerColorEnfoque = HexToLong((Right(Color, Len(Color) - 1)))
Else
    Color = "#A29D96"
    obtenerColorEnfoque = HexToLong((Right(Color, Len(Color) - 1)))
End If

End Function
