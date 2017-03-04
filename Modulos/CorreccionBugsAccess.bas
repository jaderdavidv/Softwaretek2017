Attribute VB_Name = "CorreccionBugsAccess"
Option Compare Database

' BUG # 1: Access separa los decimales utilizando coma (","), mientras MySQL lo hace con puntos "." por tanto debe reemplezarse

Public Function reemplazarSeparadorDecimal(numero As Double) As String
    reemplazarSeparadorDecimal = (Replace(CStr(numero), ",", "."))
End Function
