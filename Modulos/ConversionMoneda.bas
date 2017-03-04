Attribute VB_Name = "ConversionMoneda"
Option Compare Database

Public Function obtenerTasaMonedaActual(Fecha As Date, IDMoneda As Integer) As Double
    If (DLookup("[IDMoneda]", "Monedas", "[esmonedalocal?]=-1 And [IDMoneda]=" & IDMoneda)) Then
        obtenerTasaMonedaActual = 1
    Else
        obtenerTasaMonedaActual = Nz(DLookup("[VrEquivalente]", "MonedasTasaDiaria_v"), 0)
    End If
End Function
