Attribute VB_Name = "ValoresPredeterminados"
Option Compare Database

Dim IDDocumento As Integer
Dim IDActivo As Integer
Dim IDComprobante As Integer

Public Function getIDDocumento() As Integer
    getIDDocumento = IDDocumento
End Function

Public Sub setIDDocumento(id As Integer)
    IDDocumento = id
End Sub

'//////////////////////////////////////////////////////////

Public Function getIDActivo() As Integer
    getIDActivo = IDActivo
End Function

Public Sub setIDActivo(id As Integer)
    IDActivo = id
End Sub

'/////////////////////////////////////////////////////////

Public Function getIDComprobante() As Integer
    getIDComprobante = IDComprobante
End Function

Public Sub setIDComprobante(id As Integer)
    IDComprobante = id
End Sub

'/////////////////////////////////////////////////////////

Public Function getIDNITPredeterminado() As Integer
    getIDNITPredeterminado = 1
End Function

Public Function getIDFormaPagoPredeterminada() As Integer
    getIDFormaPagoPredeterminada = 3206
End Function

' ////////////////////////////////////////////////////////
