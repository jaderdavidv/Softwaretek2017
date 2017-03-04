Attribute VB_Name = "ConversionUnidades"
Option Compare Database
'lol
Public Function obtenerTasaConversionUnidades(medidaIni As Integer, medidaFin As Integer) As Double

Dim tasa As Double
tasa = Nz(DLookup("[tarifaconversion]", "UnidadesMedidasConversion", "[IDMedidaInicial]=" & medidaIni & " And [IDMedidaFinal]=" & medidaFin), 0)
If (tasa <> 0) Then
    obtenerTasaConversionUnidades = tasa
Else
   ' Buscamos el inverso a la conversion
   tasa = Nz(DLookup("[tarifaconversion]", "UnidadesMedidasConversion", "[IDMedidaInicial]=" & medidaFin & " And [IDMedidaFinal]=" & medidaIni), 0)
   If (tasa <> 0) Then
    obtenerTasaConversionUnidades = (1 / tasa)
   Else
    MsgBox "No se ha definido una tasa de conversión de la unidad base del producto a la unidad seleccionada, por tanto no se realizará ninguna conversión"
    obtenerTasaConversionUnidades = 1
   End If
End If

End Function
