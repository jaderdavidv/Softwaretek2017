Attribute VB_Name = "CalcularValoresMovimiento"
Option Compare Database

' ////////////////////////////////////////////////////////////////////

Public Function obtenerVrBrutoVentaActivo(IDActivo As Integer, Precio As Integer, IDSucursal As Integer) As Double
        obtenerVrBrutoVentaActivo = Nz(DLookup("[VentaVrBruto" & Precio & "]", "ActivosDatosSucursal_v", "activos.idactivo=" & IDActivo & " And [IDSucursal]=" & IDSucursal), 0)
End Function

Public Function obtenerVrCostoActivo(IDActivo As Integer) As Double
        obtenerVrCostoActivo = Nz(DLookup("[VrCostoEstandar]", "ActivosDatosSucursal_v", "activos.[idactivo]=" & IDActivo), 0)
End Function

' /////////////////////////////////////////////////////////////////////

Public Function impuestoAfectaTotal(impuesto As String) As Boolean
    impuestoAfectaTotal = Nz(DLookup("[AfectaTotales]", "ImpuestosConfiguracion", "[Impuesto]='" & impuesto & "'"), False)
End Function

Public Function aplicarPorcentaje(valor As Double, porcentaje As Double)
        aplicarPorcentaje = (valor * porcentaje / 100)
End Function

Public Function calcularVrDescuento(VrBruto As Double, PDescuento As Double) As Double
    calcularVrDescuento = aplicarPorcentaje(VrBruto, PDescuento)
End Function

'///////////////////////// CALCULO DE IMPUESTOS ///////////////////////

Public Function calcularVrIVA(VrBase As Double, PIVA As Double)
    calcularVrIVA = aplicarPorcentaje(VrBase, PIVA)
End Function

Public Function calcularVrImpoconsumo(VrBase As Double, pImpo As Double)
    calcularVrImpoconsumo = aplicarPorcentaje(VrBase, pImpo)
End Function

Public Function calcularVrImpuesto1(VrBase As Double, pImp1 As Double)
    calcularVrImpuesto1 = aplicarPorcentaje(VrBase, pImp1)
End Function

Public Function calcularVrImpuesto2(VrBase As Double, pImp2 As Double)
    calcularVrImpuesto2 = aplicarPorcentaje(VrBase, pImp2)
End Function

Public Function calcularVrImpuesto3(VrBase As Double, pImp3 As Double)
    calcularVrImpuesto3 = aplicarPorcentaje(VrBase, pImp3)
End Function

Public Function calcularVrImpuesto4(VrBase As Double, pImp4 As Double)
    calcularVrImpuesto4 = aplicarPorcentaje(VrBase, pImp4)
End Function

' ////////////////////////// CALCULO DE RETENCIONES  ///////////////////////////

Public Function calcularVrRTFT(VrBase As Double, PRTFT As Double)
    calcularVrRTFT = aplicarPorcentaje(VrBase, PRTFT)
End Function

Public Function calcularVrRTIVA(VrIVA As Double, PRTIVA As Double)
    calcularVrRTIVA = aplicarPorcentaje(VrIVA, PRTIVA)
End Function

Public Function calcularVrRTICA(VrBase As Double, PRTICA As Double)
    calcularVrRTICA = aplicarPorcentaje(VrBase, PRTICA)
End Function

Public Function calcularVrRTCREE(VrBase As Double, PRTCREE As Double)
    calcularVrRTCREE = aplicarPorcentaje(VrBase, PRTCREE)
End Function

Public Function calcularVrRetencion1(VrBase As Double, pRet1 As Double) As Double
    calcularVrRetencion1 = aplicarPorcentaje(VrBase, pRet1)
End Function

Public Function calcularVrRetencion2(VrBase As Double, pRet2 As Double) As Double
    calcularVrRetencion2 = aplicarPorcentaje(VrBase, pRet2)
End Function

Public Function calcularVrRetencion3(VrBase As Double, pRet3 As Double) As Double
    calcularVrRetencion3 = aplicarPorcentaje(VrBase, pRet3)
End Function

Public Function calcularVrRetencion4(VrBase As Double, pRet4 As Double) As Double
    calcularVrRetencion4 = aplicarPorcentaje(VrBase, pRet4)
End Function

' ////////////////////////////////////////////////////////////////////////////////
