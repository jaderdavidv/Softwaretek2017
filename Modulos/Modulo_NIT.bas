Attribute VB_Name = "Modulo_NIT"
Option Compare Database

Public Function obtenerPrecioNIT(IDNIT As Integer) As Integer
    obtenerPrecioNIT = Nz(DLookup("[precio]", "nit_v", "[id_nit]=" & IDNIT), 1)
End Function
