VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_DocumentoSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Function obtenerVrBruto(MueveValor As String, IDActivo As Integer, Precio As Integer) As Double
    If (MueveValor = "VR VENTA") Then
        obtenerVrBruto = Nz(obtenerVrBrutoVentaActivo(IDActivo, Precio, Me.Sucursal), 0)
    Else
        obtenerVrBruto = Nz(obtenerVrCostoActivo(IDActivo), 0)
    End If
End Function

Private Sub calcularVrMoneda(valor As String)
    Me.Controls(valor) = Me.Controls(valor & "Moneda") * CDbl(Nz((Me.Parent!TasaMoneda), 1))
End Sub

Private Sub ValidarCantidadExcedeExistencias()
    If (Me.RestringirCantSegunReferencia = -1 And Me.Salidas > Me.SaldoReferencia) Then
        MsgBox "Cantidad Excede Existencias"
        Me.CantidadSalidas = 0
        CalcularCantidades
    End If
End Sub

Private Sub TotalizarFactura()

        calcularVrMoneda "VrBruto"
    Me.Cantidad = Abs(Me.CantidadEntradas - Me.CantidadSalidas)
    Me.VrDescuentoMoneda = calcularVrDescuento(Me.VrBrutoMoneda, Me.PDescuento)
        calcularVrMoneda "VrDescuento"
    Me.VrBaseMoneda = Me.VrBrutoMoneda - Me.VrDescuentoMoneda
        calcularVrMoneda "VrBase"
    Me.VrIVAMoneda = calcularVrIVA(Me.VrBaseMoneda, Me.PIVA)
        calcularVrMoneda "VrIVA"
    Me.VrImpuesto1Moneda = calcularVrImpuesto1(Me.VrBaseMoneda, Me.PImpuesto1)
        calcularVrMoneda "VrImpuesto1"
    Me.VrImpuesto2Moneda = calcularVrImpuesto2(Me.VrBaseMoneda, Me.PImpuesto1)
        calcularVrMoneda "VrImpuesto2"
    Me.VrImpuesto3Moneda = calcularVrImpuesto3(Me.VrBaseMoneda, Me.PImpuesto1)
        calcularVrMoneda "VrImpuesto3"
    Me.VrImpuesto4Moneda = calcularVrImpuesto4(Me.VrBaseMoneda, Me.PImpuesto1)
        calcularVrMoneda "VrImpuesto4"
    Me.VrRTFTMoneda = calcularVrRTFT(Me.VrBaseMoneda, Me.PRTFT)
        calcularVrMoneda "VrRTFT"
    Me.VrRTIVAMoneda = calcularVrRTIVA(Me.VrIVAMoneda, Me.PRTIVA)
        calcularVrMoneda "VrRTIVA"
    Me.VrRTICAMoneda = calcularVrRTICA(Me.VrBaseMoneda, Me.PRTICA)
        calcularVrMoneda "VrRTICA"
    Me.VrRTCREEMoneda = calcularVrRTCREE(Me.VrBaseMoneda, Me.PRTCREE)
        calcularVrMoneda "VrRTCREE"
    Me.VrRetencion1Moneda = calcularVrRetencion1(Me.VrBaseMoneda, Me.PRetencion1)
        calcularVrMoneda "VrRetencion1"
    Me.VrRetencion2Moneda = calcularVrRetencion2(Me.VrBaseMoneda, Me.PRetencion2)
        calcularVrMoneda "VrRetencion2"
    Me.VrRetencion3Moneda = calcularVrRetencion3(Me.VrBaseMoneda, Me.PRetencion3)
        calcularVrMoneda "VrRetencion3"
    Me.VrRetencion4Moneda = calcularVrRetencion4(Me.VrBaseMoneda, Me.PRetencion4)
        calcularVrMoneda "VrRetencion4"
    Me.TotalImpuestosMoneda = IIf(impuestoAfectaTotal("IVA"), Me.VrIVAMoneda, 0) + IIf(impuestoAfectaTotal("IMP1"), Me.VrImpuesto1Moneda, 0) + IIf(impuestoAfectaTotal("IMP2"), Me.VrImpuesto2Moneda, 0) + IIf(impuestoAfectaTotal("IMP3"), Me.VrImpuesto3Moneda, 0) + IIf(impuestoAfectaTotal("IMP4"), Me.VrImpuesto4Moneda, 0)
        calcularVrMoneda "TotalImpuestos"
    Me.TotalRetencionesMoneda = IIf(impuestoAfectaTotal("RTFT"), Me.VrRTFTMoneda, 0) + IIf(impuestoAfectaTotal("RTIVA"), Me.VrRTIVAMoneda, 0) + IIf(impuestoAfectaTotal("RTICA"), Me.VrRTICAMoneda, 0) + IIf(impuestoAfectaTotal("RTCREE"), Me.VrRTCREEMoneda, 0) + IIf(impuestoAfectaTotal("RET1"), Me.VrRetencion1Moneda, 0) + IIf(impuestoAfectaTotal("RET2"), Me.VrRetencion2Moneda, 0) + IIf(impuestoAfectaTotal("RET3"), Me.VrRetencion3Moneda, 0) + IIf(impuestoAfectaTotal("RET4"), Me.VrRetencion4Moneda, 0)
        calcularVrMoneda "TotalRetenciones"
    Me.VrTotalMoneda = Me.VrBaseMoneda + Me.TotalImpuestosMoneda - Me.TotalRetencionesMoneda
        calcularVrMoneda "VrTotal"
    Me.VrTotalGenMoneda = Me.VrTotalMoneda * Me.Cantidad
        calcularVrMoneda "VrTotalGen"
    
End Sub

' Verificamos si un impuesto o retención afecta el total del documento o no

Private Sub CalcularCantidades()
    If (Me.IDActivo > 0) Then
        Me.Entradas = Me.TasaConversionMedida * Me.CantidadEntradas
        Me.Salidas = Me.TasaConversionMedida * Me.CantidadSalidas
    End If
    Me.Cantidad = Abs(Me.Entradas - Me.Salidas)
    TotalizarFactura
End Sub

Private Sub establecerPUCBase(tipo As Integer)
    Select Case tipo:
        Case 1:
            If (Me.Parent!MueveValor = "VR VENTA") Then
                Me.PUCBase = Me.Parent!IDPUCBase
                Me.PUCBaseNIIF = Me.Parent!IDPUCBaseNIIF
            Else
                Me.PUCBase = DLookup("[IDCostoPUC]", "Activos_v", "[IDActivo]=" & Me.IDActivo)
                Me.PUCBaseNIIF = DLookup("[IDCostoPUCniif]", "Activos_v", "[IDActivo]=" & Me.IDActivo)
            End If
        Case 2:
            Me.PUCBase = DLookup("[idcuentapuc]", "codigomovimiento_v", "[idcodigomovimiento]=" & Me.IDCodigoMovimiento)
            Me.PUCBaseNIIF = DLookup("[idcuentapucNIIF]", "codigomovimiento_v", "[idcodigomovimiento]=" & Me.IDCodigoMovimiento)
    End Select

    establecerCuentasPUC
    
End Sub

Private Sub establecerCuentasPUC()

    ' SI ES ACTIVO TRAEMOS LA CUENTA DEL COSTO DEL ACTIVO
    
    If (Me.IDActivo <> 0) Then
        Me.PUCCosto = DLookup("[idcostopuc]", "activos_v", "Activos.[idactivo]=" & Me.IDActivo)
        Me.PUCCostoContraPartida = DLookup("[IDCostoPUCContraPartida]", "activos_v", "Activos.[idactivo]=" & Me.IDActivo)
        Me.PUCCostoNIIF = DLookup("[idcostopucNIIF]", "activos_v", "Activos.[idactivo]=" & Me.IDActivo)
        Me.PUCCostoContraPartidaNIIF = DLookup("[IDCostoPUCContraPartidaNIIF]", "activos_v", "Activos.[idactivo]=" & Me.IDActivo)
    End If

    If (Me.Parent!ManejaIVA = -1) Then
        Me.PUCIVA = DLookup("[idcuentapuciva]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase)
        Me.PUCIVANIIF = DLookup("[idcuentapucivaniif]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase)
        Me.PIVA = Nz(DLookup("[porcentajeiva]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase), 0)
    End If
    
    If (Me.Parent!ManejaImp1 = -1) Then
        Me.PUCImpuesto1 = DLookup("[idcuentapucimp1]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase)
        Me.PUCImpuesto1NIIF = DLookup("[idcuentapucimp1niif]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase)
        Me.PImpuesto1 = Nz(DLookup("[porcentajeimp1]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase), 0)
    End If
    
    If (Me.Parent!ManejaImp2 = -1) Then
        Me.PUCImpuesto2 = DLookup("[idcuentapucimp2]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase)
        Me.PUCImpuesto2NIIF = DLookup("[idcuentapucimp2niif]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase)
        Me.PImpuesto2 = Nz(DLookup("[porcentajeimp2]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase), 0)
    End If
    
    If (Me.Parent!ManejaImp3 = -1) Then
        Me.PUCImpuesto3 = DLookup("[idcuentapucimp3]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase)
        Me.PUCImpuesto3NIIF = DLookup("[idcuentapucimp3niif]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase)
        Me.PImpuesto3 = Nz(DLookup("[porcentajeimp3]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase), 0)
    End If
    
    If (Me.Parent!ManejaImp4 = -1) Then
        Me.PUCImpuesto4 = DLookup("[idcuentapucimp4]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase)
        Me.PUCImpuesto4NIIF = DLookup("[idcuentapucimp4niif]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase)
        Me.PImpuesto4 = Nz(DLookup("[porcentajeimp4]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase), 0)
    End If
    
    If (Me.Parent!ManejaRTFT = -1) Then
        Me.PUCRTFT = DLookup("[idcuentapucRTFT]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase)
        Me.PUCRTFTNIIF = DLookup("[idcuentapucRTFTNIIF]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase)
        Me.PRTFT = Nz(DLookup("[porcentajeRTFT]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase), 0)
    End If
    
    If (Me.Parent!ManejaRTIVA = -1) Then
        Me.PUCRTIVA = DLookup("[idcuentapucRTIVA]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase)
        Me.PUCRTIVANIIF = DLookup("[idcuentapucRTIVANIIF]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase)
        Me.PRTIVA = Nz(DLookup("[porcentajeRTIVA]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase), 0)
    End If
    
    If (Me.Parent!ManejaRTICA = -1) Then
        Me.PUCRTICA = DLookup("[idcuentapucRTICA]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase)
        Me.PUCRTICANIIF = DLookup("[idcuentapucRTICANIIF]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase)
        Me.PRTICA = Nz(DLookup("[porcentajeRTICA]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase), 0)
    End If
    
    If (Me.Parent!ManejaRTCREE = -1) Then
        Me.PUCRTCREE = DLookup("[IDCuentaPUCRTCREE]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase)
        Me.PUCRTCREENIIF = DLookup("[IDCuentaPUCRTCREENIIF]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase)
        Me.PRTCREE = Nz(DLookup("[porcentajeRTCREE]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase), 0)
    End If
    
    If (Me.Parent!ManejaRetencion1 = -1) Then
        Me.PUCRetencion1 = DLookup("[idcuentapucRetencion1]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase)
        Me.PUCRetencion1NIIF = DLookup("[idcuentapucRetencion1niif]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase)
        Me.PRetencion1 = Nz(DLookup("[porcentajeRetencion1]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase), 0)
    End If
    
    If (Me.Parent!ManejaRetencion2 = -1) Then
        Me.PUCRetencion2 = DLookup("[idcuentapucRetencion2]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase)
        Me.PUCRetencion2NIIF = DLookup("[idcuentapucRetencion2niif]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase)
        Me.PRetencion2 = Nz(DLookup("[porcentajeRetencion2]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase), 0)
    End If
    
    If (Me.Parent!ManejaRetencion3 = -1) Then
        Me.PUCRetencion3 = DLookup("[idcuentapucRetencion3]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase)
        Me.PUCRetencion3NIIF = DLookup("[idcuentapucRetencion3niif]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase)
        Me.PRetencion3 = Nz(DLookup("[porcentajeRetencion3]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase), 0)
    End If
    
    If (Me.Parent!ManejaRetencion4 = -1) Then
        Me.PUCRetencion4 = DLookup("[idcuentapucRetencion4]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase)
        Me.PUCRetencion4NIIF = DLookup("[idcuentapucRetencion4niif]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase)
        Me.PRetencion4 = Nz(DLookup("[porcentajeRetencion4]", "cuentaspuc_v", "[idcuentaPUC]=" & Me.PUCBase), 0)
    End If
    
End Sub


Private Sub RegistroIngresado()

   Me.IDNIT = Nz((Me.Parent!IDNIT), 1)
   Me.CDC = Nz((Me.Parent!IDCDC), 1)
   Me.Sucursal = Nz((Me.Parent!IDSucursal), 1)
   Me.NoDocRef = Nz((Me.Parent!NoDocumento), 1)
   Me.RestringirCantSegunReferencia = Nz(Me.Parent!RestringirSalidasSegunSaldo, 0)
   
End Sub

Private Sub AsignarNombreMostrar(tipo As Integer)
    Select Case tipo
        Case 1: Me.NombreParaMostrar = DLookup("[codigoactivo]", "activos_v", "[idactivo]=" & Me.IDActivo) & " - " & DLookup("[nombreactivo]", "activos_v", "[idactivo]=" & Me.IDActivo)
        Case 2: Me.NombreParaMostrar = DLookup("[nombremovimiento]", "codigomovimiento_v", "[idcodigomovimiento]=" & Me.IDCodigoMovimiento)
    End Select
End Sub

' SERIE DE EVENTOS QUE SE REALIZAN AL INGRESAR EL ACTIVO

Private Sub ActivoIngresado()
    
    Me.Precio = Nz(Me.Parent!ListaPrecios, 1)
    Me.VrBruto = Nz(obtenerVrBruto(Me.Parent!MueveValor, Me.IDActivo, Me.Precio), 0)
    Me.VrBrutoMoneda = Me.VrBruto * CDbl(Nz((Me.Parent!TasaMoneda), 1))
    Me.MedidaEntradas = DLookup("[IDUnidadPredeterminada]", "activosdatossucursal", "[idactivo]=" & Me.IDActivo & " And [idsucursal]=" & Me.Sucursal)
    Me.MedidaSalidas = DLookup("[IDUnidadPredeterminada]", "activosdatossucursal", "[idactivo]=" & Me.IDActivo & " And [idsucursal]=" & Me.Sucursal)
    Me.SaldoReferencia = obtenerSaldoActivoSucursal(Me.IDActivo, Me.Sucursal)
    
    ' Si el documento no afecta costos, el costo es el promedio actual
    If (Me.Parent!AfectaCostoPromedio = 0) Then
        Me.VrCosto = Nz(DLookup("[VrCostoPromedio]", "activosdatossucursal", "[idactivo]=" & Me.IDActivo & " And [idsucursal]=" & Me.Sucursal), 0)
    End If
    
    TotalizarFactura
    
    If (Me.IDActivo > 0) Then
        establecerPUCBase 1
        AsignarNombreMostrar 1
    End If
    
    
    
End Sub

Private Sub Form_Load()

    ' NOMBRA TODOS LOS CAMBIOS DE IMPUESTOS Y RETENCIONES SEGUN EL NOMBRE ASIGNADO EN LA TABLA CONFIGURACIONIMPUESTOS
    
    Dim i As Integer
    For i = 1 To 4
        Me.Controls("PImpuesto" & i & "_Etiqueta").Caption = "%" & getNombreImpuesto(i)
        Me.Controls("VrImpuesto" & i & "_Etiqueta").Caption = getNombreImpuesto(i)
        Me.Controls("PRetencion" & i & "_Etiqueta").Caption = "%" & getNombreImpuesto(i)
        Me.Controls("VrRetencion" & i & "_Etiqueta").Caption = getNombreImpuesto(i)
    Next i
    
End Sub

Private Sub CantidadEntradas_AfterUpdate()
    CalcularCantidades
End Sub

Private Sub CantidadSalidas_AfterUpdate()
    CalcularCantidades
    ValidarCantidadExcedeExistencias
End Sub

Private Sub DescripcionActivo_AfterUpdate()
    RegistroIngresado
    ActivoIngresado
End Sub

Private Sub IDActivo_AfterUpdate()
    RegistroIngresado
    ActivoIngresado
End Sub

Private Sub IDCodigoMovimiento_AfterUpdate()
    If (Me.IDCodigoMovimiento > 0 And Me.IDActivo = 0) Then
        establecerPUCBase 2
    End If
    AsignarNombreMostrar 2
    RegistroIngresado
End Sub

Private Sub IDNIT_AfterUpdate()
    If (Me.IDNIT2 = 0) Then
        Me.IDNIT2 = Me.IDNIT
    End If
End Sub

Private Sub MedidaEntradas_Change()

    Dim unidadinicial As Integer
    Dim unidadfinal As Integer
    unidadinicial = DLookup("[IDMedidaBase]", "Activos_v", "[Activos].[IDActivo]=" & Me.IDActivo)
    unidadfinal = Me.MedidaEntradas
    Me.TasaConversionMedida = obtenerTasaConversionUnidades(unidadinicial, unidadfinal)
    CalcularCantidades
    TotalizarFactura

End Sub

Private Sub MedidaSalidas_Change()

    Dim unidadinicial As Integer
    Dim unidadfinal As Integer
    unidadinicial = DLookup("[IDMedidaBase]", "Activos_v", "[Activos].[IDActivo]=" & Me.IDActivo)
    unidadfinal = Me.MedidaSalidas
    Me.TasaConversionMedida = obtenerTasaConversionUnidades(unidadinicial, unidadfinal)
    CalcularCantidades
    TotalizarFactura
    
End Sub

Private Sub PDescuento_AfterUpdate()
    TotalizarFactura
End Sub

Private Sub PImpuesto1_AfterUpdate()
     TotalizarFactura
End Sub

Private Sub PImpuesto2_AfterUpdate()
     TotalizarFactura
End Sub

Private Sub PImpuesto3_AfterUpdate()
     TotalizarFactura
End Sub

Private Sub PImpuesto4_AfterUpdate()
     TotalizarFactura
End Sub

Private Sub PIVA_AfterUpdate()
     TotalizarFactura
     Me.PUCIVA = Nz(DLookup("[IDCuentaPUC]", "Comprobante_PUCTarifas", "[Impuesto]='IVA' And [IDComprobante]=" & Me.Parent!IDComprobante & " And [Tarifa]=" & Me.PIVA), getCuentaIVAPredeterminada)
End Sub

Private Sub PRTCREE_AfterUpdate()
     TotalizarFactura
End Sub

Private Sub PRTFT_AfterUpdate()
     TotalizarFactura
End Sub

Private Sub PRTICA_AfterUpdate()
     TotalizarFactura
End Sub

Private Sub PRetencion1_AfterUpdate()
    TotalizarFactura
End Sub

Private Sub PRetencion2_AfterUpdate()
     TotalizarFactura
End Sub

Private Sub PRetencion3_AfterUpdate()
    TotalizarFactura
End Sub

Private Sub PRetencion4_AfterUpdate()
     TotalizarFactura
End Sub

Private Sub PRTIVA_AfterUpdate()
    TotalizarFactura
End Sub

Private Sub VrBrutoMoneda_AfterUpdate()
     TotalizarFactura
End Sub
