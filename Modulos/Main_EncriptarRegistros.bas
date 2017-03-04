Attribute VB_Name = "Main_EncriptarRegistros"
Option Compare Database
Option Explicit

'////////////////////////////

Private Function ConvToHex(X As Integer) As String
    If X > 9 Then
        ConvToHex = Chr(X + 55)
    Else
        ConvToHex = CStr(X)
    End If
End Function
  

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 Function Encriptar(DataValue As Variant) As Variant
      
    Dim X As Long
    Dim Temp As String
    Dim TempNum As Integer
    Dim TempChar As String
    Dim TempChar2 As String
      
    For X = 1 To Len(DataValue)
        TempChar2 = Mid(DataValue, X, 1)
        TempNum = Int(Asc(TempChar2) / 16)
          
        If ((TempNum * 16) < Asc(TempChar2)) Then
                 
            TempChar = ConvToHex(Asc(TempChar2) - (TempNum * 16))
            Temp = Temp & ConvToHex(TempNum) & TempChar
        Else
            Temp = Temp & ConvToHex(TempNum) & "0"
          
        End If
    Next X
      
      
    Encriptar = Temp
End Function

Private Function ConvToInt(X As String) As Integer
      
    Dim x1 As String
    Dim x2 As String
    Dim Temp As Integer
      
    x1 = Mid(X, 1, 1)
    x2 = Mid(X, 2, 1)
      
    If IsNumeric(x1) Then
        Temp = 16 * Int(x1)
    Else
        Temp = (Asc(x1) - 55) * 16
    End If
      
    If IsNumeric(x2) Then
        Temp = Temp + Int(x2)
    Else
        Temp = Temp + (Asc(x2) - 55)
    End If
      
    ' retorno
    ConvToInt = Temp
      
End Function
  
' función que decodifica el dato
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function Desencriptar(DataValue As Variant) As Variant
      
    Dim X As Long
    Dim Temp As String
    Dim HexByte As String
      
    For X = 1 To Len(DataValue) Step 2
          
        HexByte = Mid(DataValue, X, 2)
        Temp = Temp & Chr(ConvToInt(HexByte))
          
    Next X
    ' retorno
    Desencriptar = Temp
      
End Function
