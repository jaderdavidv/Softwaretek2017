Attribute VB_Name = "Modulo_NombresImpuestos"
Option Compare Database

Private Function getNombre(nombreimp As String) As String
    getNombre = DLookup("[Abreviacion]", "ImpuestosConfiguracion", "[Impuesto]='" & nombreimp & "'")
End Function

Public Function getNombreImpuesto(Indice As Integer) As String
    getNombreImpuesto = Nz(getNombre(("Imp" & Indice)), "Imp" & Indice)
End Function

Public Function getNombreRetencion(Indice As Integer) As String
    getNombreRetencion = Nz(getNombre("Ret" & Indice), "Ret" & Indice)
End Function





