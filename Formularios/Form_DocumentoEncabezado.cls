VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_DocumentoEncabezado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

' Insertar Movimiento en las tablas en una sola transacción

Private Sub TransaccionInsertarDocumento(id As Double)
    
    Dim wrk As DAO.Workspace
    Dim dbC As DAO.Database
    Dim dbX As DAO.Database
    Set wrk = DBEngine(0)
    Set dbC = CurrentDb
    Set dbX = wrk.OpenDatabase(CurrentProject.Path & "\Softwaretek.accdb")
    
    On Error GoTo trans_Err
    
    wrk.BeginTrans
    dbC.Execute "update DocumentoEncabezado set DocumentoEncabezado.eliminadosistema = -1 where DocumentoEncabezado.IDDocumento=" & id & ";", dbFailOnError
    dbC.Execute "update DocumentoSub set DocumentoSub.eliminadoSistema = -1 where DocumentoSub.IDDocumentoEncabezado=" & id & ";", dbFailOnError
    dbC.Execute "insert into DocumentoEncabezado select DocumentoTotalizarEncabezado.* from DocumentoTotalizarEncabezado where DocumentoTotalizarEncabezado.IDDocumento=" & id & ";", dbFailOnError
    dbC.Execute "insert into DocumentoSub select DocumentoSubUser.* from DocumentoSubUser where (((DocumentoSubUser.IDDocumentoEncabezado)=" & id & "));", dbFailOnError
    wrk.CommitTrans dbForceOSFlush
    
trans_ExitOK:

    wrk.Close
    Set dbC = Nothing
    Set dbX = Nothing
    Set wrk = Nothing
    Exit Sub
    
trans_ExitError:

    wrk.Close
    Set dbC = Nothing
    Set dbX = Nothing
    Set wrk = Nothing
    End
    
trans_Err:
    wrk.Rollback
    MsgBox "Ocurrio un error al momento de realizar esta transacción" & vbNewLine & Err.Description, vbCritical, "Error Code: " & Err.Number
    Resume trans_ExitError
    
End Sub


Private Function ActivoExcedeExistencias() As Boolean

If (Nz(DCount("[IDActivo]", "Documento_ValidarSaldosActivosEnMovimiento02"), 0) > 0) Then
    ActivoExcedeExistencias = True
Else
    ActivoExcedeExistencias = False
End If

End Function


Private Sub finalizarMovimiento()
    ValidarCamposRequeridos
    Me.Valor1 = Me.TotalGeneral
    DoCmd.RunCommand acCmdRefresh
    
    ' VALIDAMOS QUE LOS ITEMS NO EXCEDAN SALDO EN INVENTARIO
    If (ActivoExcedeExistencias) Then
        DoCmd.OpenForm "Documento_ListaActivosExcedenSaldo"
        End
    End If
    
    setIDDocumento (Me.IDDocumento)
    TransaccionInsertarDocumento Me.IDDocumento
    ContabilizarDocumento
    ActualizarDatosActivoSucursal getBoolean(Me.AfectaCostoPromedio), getBoolean(Me.MueveInventarios)
    DoCmd.RunSQL "delete movimientocontableuser.* from movimientocontableuser;", -1
    DoCmd.OpenReport "Documento_MediaCarta01", acViewPreview
End Sub

' Determina si el documento maneja retenciones segun el nit y el comprobante

Private Sub ConfigurarManejoRetenciones()

    Me.ManejaRTFT = AplicarRetencionEnDocumento(Me.IDNIT, Me.IDComprobante, "RTFT")
    Me.ManejaRTIVA = AplicarRetencionEnDocumento(Me.IDNIT, Me.IDComprobante, "RTIVA")
    Me.ManejaRTICA = AplicarRetencionEnDocumento(Me.IDNIT, Me.IDComprobante, "RTICA")
    Me.ManejaRTCREE = AplicarRetencionEnDocumento(Me.IDNIT, Me.IDComprobante, "RTCREE")
    Me.ManejaRetencion1 = AplicarRetencionEnDocumento(Me.IDNIT, Me.IDComprobante, "Retencion1")
    Me.ManejaRetencion2 = AplicarRetencionEnDocumento(Me.IDNIT, Me.IDComprobante, "Retencion2")
    Me.ManejaRetencion3 = AplicarRetencionEnDocumento(Me.IDNIT, Me.IDComprobante, "Retencion3")
    Me.ManejaRetencion4 = AplicarRetencionEnDocumento(Me.IDNIT, Me.IDComprobante, "Retencion4")
    
End Sub


' Hace visibles/invisibles campos en el sub formulario

Private Sub HacerCampoVisibleSubFormulario(nombreCampo As String, visibilidad As Boolean)
    If (visibilidad) Then
        Me.DocumentoSub(nombreCampo).ColumnHidden = False
    Else
        Me.DocumentoSub(nombreCampo).ColumnHidden = True
    End If
End Sub

' Hace visibles/invisibles campos en el formulario

Private Sub HacerCampoVisibleFormulario(nombreCampo As String, visibilidad As Boolean)
    If (visibilidad) Then
        Me(nombreCampo).Visible = True
    Else
        Me(nombreCampo).Visible = False
    End If
End Sub

Private Sub ConfigurarFormulario()

    Me.NoDocumento = Me.Prefijo & (Nz(Me.Consecutivo, 0) + 1)
    Me.DocumentoSub!NaturalezaBase = Me.NaturalezaBase
    
    If (Me.ConsecutivoAutomatico = 1) Then
        Me.NoDocumento.Enabled = False
        Me.NoDocumento.Locked = True
    Else
        Me.NoDocumento.Enabled = True
        Me.NoDocumento.Locked = False
    End If
    
    Me.RestringirSalidasSegunSaldo = Nz(DLookup("[RestringirSalidasSegunSaldo]", "comprobante_v", "[idcomprobante]=" & Me.IDComprobante), False)

    ' CAMPOS DEL FORMULARIO
    
    HacerCampoVisibleFormulario "DocReferencia", getBoolean(Me.ManejaDocReferencia)
    HacerCampoVisibleFormulario "IDMoneda", getBoolean(Me.ManejaMoneda)
    HacerCampoVisibleFormulario "TasaMoneda", getBoolean(Me.ManejaMoneda)
    HacerCampoVisibleFormulario "Dato1", getBoolean(Me.ManejaDatoEncabezado1)
    HacerCampoVisibleFormulario "Dato2", getBoolean(Me.ManejaDatoEncabezado2)
    HacerCampoVisibleFormulario "Dato3", getBoolean(Me.ManejaDatoEncabezado3)
    HacerCampoVisibleFormulario "Dato4", getBoolean(Me.ManejaDatoEncabezado4)
    HacerCampoVisibleFormulario "IDPedido", getBoolean(Me.ManejaPedido)
    HacerCampoVisibleFormulario "FechaEmision", getBoolean(Me.ManejaFechaEmision)
    HacerCampoVisibleFormulario "IDUsuario", getBoolean(Me.ExigeUsuario)
    HacerCampoVisibleFormulario "Contraseña", getBoolean(Me.ExigeUsuario)
    HacerCampoVisibleFormulario "ListaPrecios", getBoolean(Me.ManejaListaPrecios)
    HacerCampoVisibleFormulario "IDVendedor", getBoolean(Me.ManejaVendedor)
    
    ' CAMPOS DEL SUBFORMULARIO

    HacerCampoVisibleSubFormulario "CodigoBarra", getBoolean(Me.ManejaCodigoBarra)
    HacerCampoVisibleSubFormulario "IDCodigoMovimiento", getBoolean(Me.ManejaCodigoMovimiento)
    HacerCampoVisibleSubFormulario "IDActivo", getBoolean(Me.ManejaIDActivo)
    HacerCampoVisibleSubFormulario "IDProceso", getBoolean(Me.ManejaProceso)
    HacerCampoVisibleSubFormulario "DescripcionActivo", getBoolean(Me.ManejaIDActivo)
    HacerCampoVisibleSubFormulario "IDNIT", getBoolean(Me.ManejaNIT)
    HacerCampoVisibleSubFormulario "Observación", getBoolean(Me.ManejaObservacion)
    HacerCampoVisibleSubFormulario "Lote", getBoolean(Me.ManejaLote)
    HacerCampoVisibleSubFormulario "Talla", getBoolean(Me.ManejaTalla)
    HacerCampoVisibleSubFormulario "Color", getBoolean(Me.ManejaColor)
    HacerCampoVisibleSubFormulario "CDC", getBoolean(Me.ManejaCDC)
    HacerCampoVisibleSubFormulario "SubCDC", getBoolean(Me.ManejaSubCDC)
    HacerCampoVisibleSubFormulario "Sucursal", getBoolean(Me.ManejaSucursal)
    HacerCampoVisibleSubFormulario "Bodega", getBoolean(Me.ManejaBodega)
    HacerCampoVisibleSubFormulario "VidaUtil", getBoolean(Me.ManejaVidaUtil)
    HacerCampoVisibleSubFormulario "SubDatoExtra1", getBoolean(Me.ManejaDatoExtra1)
    HacerCampoVisibleSubFormulario "SubDatoExtra2", getBoolean(Me.ManejaDatoExtra2)
    HacerCampoVisibleSubFormulario "SubDatoExtra3", getBoolean(Me.ManejaDatoExtra3)
    HacerCampoVisibleSubFormulario "SubDatoExtra4", getBoolean(Me.ManejaDatoExtra4)
    HacerCampoVisibleSubFormulario "FechaHoraInicio", getBoolean(Me.ManejaFechaHoraInicio)
    HacerCampoVisibleSubFormulario "FechaHoraFin", getBoolean(Me.ManejaFechaHoraFin)
    HacerCampoVisibleSubFormulario "NoCheque", getBoolean(Me.ManejaNoCheque)
    HacerCampoVisibleSubFormulario "EsCajaMenor", getBoolean(Me.ManejaEsCajaMenor)
    HacerCampoVisibleSubFormulario "Imprimir", getBoolean(Me.ManejaOpcionImprimir)
    
    HacerCampoVisibleSubFormulario "Plazo", getBoolean(Me.ManejaPagoPorCuotas)
    HacerCampoVisibleSubFormulario "NoCuotas", getBoolean(Me.ManejaPagoPorCuotas)
    HacerCampoVisibleSubFormulario "VrCuota", getBoolean(Me.ManejaPagoPorCuotas)
    
    HacerCampoVisibleSubFormulario "VrCosto", getBoolean(Me.ManejaVrCosto)
    HacerCampoVisibleSubFormulario "PDescuento", getBoolean(Me.ManejaDescuento)
    HacerCampoVisibleSubFormulario "PIVA", getBoolean(Me.ManejaIVA)
    HacerCampoVisibleSubFormulario "PRTFT", getBoolean(Me.ManejaRTFT)
    HacerCampoVisibleSubFormulario "PRTIVA", getBoolean(Me.ManejaRTIVA)
    HacerCampoVisibleSubFormulario "PRTICA", getBoolean(Me.ManejaRTICA)
    HacerCampoVisibleSubFormulario "PRTCREE", getBoolean(Me.ManejaRTCREE)
    HacerCampoVisibleSubFormulario "PImpuesto1", getBoolean(Me.ManejaImp1)
    HacerCampoVisibleSubFormulario "PImpuesto2", getBoolean(Me.ManejaImp2)
    HacerCampoVisibleSubFormulario "PImpuesto3", getBoolean(Me.ManejaImp3)
    HacerCampoVisibleSubFormulario "PImpuesto4", getBoolean(Me.ManejaImp4)
    HacerCampoVisibleSubFormulario "PRetencion1", getBoolean(Me.ManejaRetencion1)
    HacerCampoVisibleSubFormulario "PRetencion2", getBoolean(Me.ManejaRetencion2)
    HacerCampoVisibleSubFormulario "PRetencion3", getBoolean(Me.ManejaRetencion3)
    HacerCampoVisibleSubFormulario "PRetencion4", getBoolean(Me.ManejaRetencion4)
    
    HacerCampoVisibleSubFormulario "MedidaEntradas", getBoolean(Me.ManejaUnidad * Me.ManejaEntradas)
    HacerCampoVisibleSubFormulario "MedidaSalidas", getBoolean(Me.ManejaUnidad * Me.ManejaSalidas)
    
    HacerCampoVisibleSubFormulario "CantEntradasReferencia", getBoolean((Me.ManejaCantidadReferencia * Me.ManejaEntradas))
    HacerCampoVisibleSubFormulario "CantidadEntradas", getBoolean(Me.ManejaEntradas)
    HacerCampoVisibleSubFormulario "CantidadEntradasInv", getBoolean(Me.ManejaDobleCantidad)
    HacerCampoVisibleSubFormulario "CantSalidasReferencia", getBoolean((Me.ManejaCantidadReferencia * Me.ManejaSalidas))
    HacerCampoVisibleSubFormulario "CantidadSalidas", getBoolean(Me.ManejaSalidas)
    HacerCampoVisibleSubFormulario "CantidadSalidasInv", getBoolean(Me.ManejaDobleCantidad)
    HacerCampoVisibleSubFormulario "SaldoReferencia", getBoolean(Me.MostrarExistencias)
    
End Sub

Private Sub MostrarBotones()
    If (Me.btnGuardarPendiente.Visible = False) Then
        Me.btnGuardarPendiente.Visible = True
        Me.btnImprimir.Visible = True
        Me.btnAnular.Visible = True
        Me.btnCopiarDocumento.Visible = True
        Me.btnCargarPlantilla.Visible = True
    End If
End Sub



Private Sub ContabilizarDocumento()

On Error GoTo fun_Err

    DoCmd.OpenQuery "ContabilizarMov_01CostoMercVendidaDB"
    DoCmd.OpenQuery "ContabilizarMov_02CostoMercVendidaCD"
    DoCmd.OpenQuery "ContabilizarMov_03VrBase"
    DoCmd.OpenQuery "ContabilizarMov_17FormaPago1"
    
    If (Me.ManejaIVA) Then
        DoCmd.OpenQuery "ContabilizarMov_04VrIVA"
    End If
    If (Me.ManejaImp1) Then
        DoCmd.OpenQuery "ContabilizarMov_05Imp1"
    End If
    If (Me.ManejaImp2) Then
        DoCmd.OpenQuery "ContabilizarMov_06Imp2"
    End If
    If (Me.ManejaImp3) Then
        DoCmd.OpenQuery "ContabilizarMov_07Imp3"
    End If
    If (Me.ManejaImp4) Then
        DoCmd.OpenQuery "ContabilizarMov_08Imp4"
    End If
    If (Me.ManejaRTFT) Then
        DoCmd.OpenQuery "ContabilizarMov_09RTFT"
    End If
    If (Me.ManejaRTIVA) Then
        DoCmd.OpenQuery "ContabilizarMov_10RTIVA"
    End If
    If (Me.ManejaRTICA) Then
        DoCmd.OpenQuery "ContabilizarMov_11RTICA"
    End If
    If (Me.ManejaRTCREE) Then
        DoCmd.OpenQuery "ContabilizarMov_12RTCREE"
    End If
    If (Me.ManejaRetencion1) Then
        DoCmd.OpenQuery "ContabilizarMov_13Ret1"
    End If
    If (Me.ManejaRetencion2) Then
        DoCmd.OpenQuery "ContabilizarMov_14Ret2"
    End If
    If (Me.ManejaRetencion3) Then
        DoCmd.OpenQuery "ContabilizarMov_15Ret3"
    End If
    If (Me.ManejaRetencion4) Then
        DoCmd.OpenQuery "ContabilizarMov_16Ret4"
    End If
    If (Me.ManejaFormaPago2) Then
        DoCmd.OpenQuery "ContabilizarMov_18FormaPago2"
    End If
    If (Me.ManejaFormaPago3) Then
        DoCmd.OpenQuery "ContabilizarMov_19FormaPago3"
    End If
    If (Me.ManejaFormaPago4) Then
        DoCmd.OpenQuery "ContabilizarMov_20FormaPago4"
    End If
    If (Me.ManejaFormaPago5) Then
        DoCmd.OpenQuery "ContabilizarMov_21FormaPago5"
    End If

    If (Me.MueveInventarios) Then
        DoCmd.OpenQuery "ContabilizarMov_22ActualizarMovMueveInventario"
    End If
    
    DoCmd.RunSQL "insert into movimientocontable select movimientocontableuser.* from movimientocontableuser;", -1
    Exit Sub
    
fun_Err:
    MsgBox "Ocurrio un error al momento de contabilizar este documento" & vbNewLine & Err.Description, vbCritical, "Error Code: " & Err.Number
    End

End Sub

Private Sub ValidarCamposRequeridos()
    
    If (Me.ExigeDocReferencia = -1 And Me.DocReferencia = 0) Then
        MsgBox "Favor Ingresar No Documento de Referencia"
        DoCmd.GoToControl Me.DocReferencia.Name
        End
    End If
    
    If (Me.ExigeMoneda = -1 And Me.IDMoneda = 1) Then
        MsgBox "Favor Ingresar Moneda para este documento"
        DoCmd.GoToControl Me.IDMoneda.Name
        End
    End If
    
    If (Me.ExigePedido = -1 And Me.IDPedido = 0) Then
        MsgBox "Favor Ingresar No Pedido"
        DoCmd.GoToControl Me.IDPedido.Name
        End
    End If
    
    If (Me.ExigeProceso = -1 And Nz(DLookup("[idproceso]", "[documentosubuser]", "[idproceso]=0"), 0) <> 0) Then
        MsgBox "Favor Ingresar IDProceso de cada rubro de este documento"
        End
    End If
    
    If (Me.ExigeNIT = -1 And Me.IDNIT = "0") Then
        MsgBox "Favor Ingresar Tercero Válido"
        DoCmd.GoToControl Me.IDNIT.Name
        End
    End If
    
    If (Me.ExigeUsuario = -1 And VerificarUsuario(Me.IDUsuario, Encriptar(Me.Contraseña)) = 0) Then
        MsgBox "Favor Ingresar Usuario y Contraseña"
        DoCmd.GoToControl Me.IDUsuario.Name
        End
    End If
    
    If (Me.ExigeVendedor = -1 And Me.IDVendedor = 0) Then
        MsgBox "Favor Ingresar Vendedor"
        DoCmd.GoToControl Me.IDVendedor.Name
        End
    End If

End Sub

Private Sub btnGuardar_Click()
 finalizarMovimiento
End Sub

Private Sub Comando305_Click()
    DoCmd.Close
End Sub

Private Sub Comando1_Click()
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub Form_Close()
    DoCmd.SetWarnings True
    DoCmd.RunSQL "delete documentoencabezadouser.* from documentoencabezadouser;", -1
End Sub

Private Sub Form_Current()
ConfigurarFormulario
End Sub

Private Sub Form_Load()
DoCmd.SetWarnings False
End Sub

Private Sub IDComprobante_AfterUpdate()
    ConfigurarFormulario
End Sub

' EVENTOS DE RECIBIR Y PERDER ENFOQUE

Private Sub Dato1_GotFocus()
    Me.Dato1_Etiqueta.ForeColor = obtenerColorEnfoque(True)
End Sub

Private Sub Dato1_LostFocus()
    Me.Dato1_Etiqueta.ForeColor = obtenerColorEnfoque(False)
End Sub

Private Sub Dato2_GotFocus()
    Me.Dato2_Etiqueta.ForeColor = obtenerColorEnfoque(True)
End Sub

Private Sub Dato2_LostFocus()
    Me.Dato2_Etiqueta.ForeColor = obtenerColorEnfoque(False)
End Sub

Private Sub Dato3_GotFocus()
    Me.Dato3_Etiqueta.ForeColor = obtenerColorEnfoque(True)
End Sub

Private Sub Dato3_LostFocus()
    Me.Dato3_Etiqueta.ForeColor = obtenerColorEnfoque(False)
End Sub

Private Sub Dato4_GotFocus()
    Me.Dato4_Etiqueta.ForeColor = obtenerColorEnfoque(True)
End Sub

Private Sub Dato4_LostFocus()
    Me.Dato4_Etiqueta.ForeColor = obtenerColorEnfoque(False)
End Sub

Private Sub contraseña_GotFocus()
    Me.Contraseña_Etiqueta.ForeColor = obtenerColorEnfoque(True)
End Sub

Private Sub contraseña_LostFocus()
    Me.Contraseña_Etiqueta.ForeColor = obtenerColorEnfoque(False)
End Sub

Private Sub DocReferencia_GotFocus()
    Me.DocReferencia_Etiqueta.ForeColor = obtenerColorEnfoque(True)
End Sub

Private Sub DocReferencia_LostFocus()
    Me.DocReferencia_Etiqueta.ForeColor = obtenerColorEnfoque(False)
End Sub

Private Sub Fecha_GotFocus()
    Me.Fecha_Etiqueta.ForeColor = obtenerColorEnfoque(True)
End Sub

Private Sub Fecha_LostFocus()
    Me.Fecha_Etiqueta.ForeColor = obtenerColorEnfoque(False)
End Sub

Private Sub FechaEmision_GotFocus()
    Me.FechaEmision_Etiqueta.ForeColor = obtenerColorEnfoque(True)
End Sub

Private Sub FechaEmision_LostFocus()
    Me.FechaEmision_Etiqueta.ForeColor = obtenerColorEnfoque(False)
End Sub

Private Sub IDCDC_GotFocus()
    Me.IDCDC_Etiqueta.ForeColor = obtenerColorEnfoque(True)
End Sub

Private Sub IDCDC_LostFocus()
    Me.IDCDC_Etiqueta.ForeColor = obtenerColorEnfoque(False)
End Sub

Private Sub IDComprobante_GotFocus()
    Me.IDComprobante_Etiqueta.ForeColor = obtenerColorEnfoque(True)
    Me.ActiveControl.Dropdown
End Sub

Private Sub IDComprobante_LostFocus()
    Me.IDComprobante_Etiqueta.ForeColor = obtenerColorEnfoque(False)
End Sub

Private Sub IDMoneda_Change()
    Me.TasaMoneda = obtenerTasaMonedaActual(Me.Fecha, Me.IDMoneda)
    DoCmd.OpenQuery "Movimiento_RecalcularVrMoneda"
End Sub

Private Sub IDMoneda_GotFocus()
    Me.IDMoneda_Etiqueta.ForeColor = obtenerColorEnfoque(True)
End Sub

Private Sub IDMoneda_LostFocus()
    Me.IDMoneda_Etiqueta.ForeColor = obtenerColorEnfoque(False)
End Sub

Private Sub IDNIT_AfterUpdate()
    Me.ListaPrecios = obtenerPrecioNIT(Me.IDNIT)
    Me.IDPUCFormaPago1 = Nz(DLookup("[IDPUCFormaPago]", "nit_v", "[id_nit]=" & Me.IDNIT), 3206)
    ConfigurarManejoRetenciones
    ConfigurarFormulario
End Sub

Private Sub IDNIT_GotFocus()
    Me.IDNIT_Etiqueta.ForeColor = obtenerColorEnfoque(True)
End Sub

Private Sub IDNIT_LostFocus()
    Me.IDNIT_Etiqueta.ForeColor = obtenerColorEnfoque(False)
End Sub

Private Sub IDPedido_GotFocus()
    Me.IDPedido_Etiqueta.ForeColor = obtenerColorEnfoque(True)
End Sub

Private Sub IDPedido_LostFocus()
    Me.IDPedido_Etiqueta.ForeColor = obtenerColorEnfoque(False)
End Sub

Private Sub IDSucursal_GotFocus()
    Me.IDSucursal_Etiqueta.ForeColor = obtenerColorEnfoque(True)
End Sub

Private Sub IDSucursal_LostFocus()
    Me.IDSucursal_Etiqueta.ForeColor = obtenerColorEnfoque(False)
End Sub

Private Sub IDUsuario_GotFocus()
    Me.IDUsuario_Etiqueta.ForeColor = obtenerColorEnfoque(True)
End Sub

Private Sub IDUsuario_LostFocus()
    Me.IDUsuario_Etiqueta.ForeColor = obtenerColorEnfoque(False)
End Sub

Private Sub IDVendedor_GotFocus()
    Me.IDVendedor_Etiqueta.ForeColor = obtenerColorEnfoque(True)
End Sub

Private Sub IDVendedor_LostFocus()
    Me.IDVendedor_Etiqueta.ForeColor = obtenerColorEnfoque(False)
End Sub

Private Sub ListaPrecios_GotFocus()
     Me.ListaPrecios_Etiqueta.ForeColor = obtenerColorEnfoque(True)
End Sub

Private Sub ListaPrecios_LostFocus()
     Me.ListaPrecios_Etiqueta.ForeColor = obtenerColorEnfoque(False)
End Sub

Private Sub masopciones_btn_Click()
    DoCmd.OpenForm "documento_masopciones", acNormal
End Sub

Private Sub NoDocumento_GotFocus()
    Me.NoDocumento_Etiqueta.ForeColor = obtenerColorEnfoque(True)
End Sub

Private Sub NoDocumento_LostFocus()
    Me.NoDocumento_Etiqueta.ForeColor = obtenerColorEnfoque(False)
End Sub

Private Sub Observacion_GotFocus()
    Me.Observacion_Etiqueta.ForeColor = obtenerColorEnfoque(True)
End Sub
Private Sub Observacion_LostFocus()
    Me.Observacion_Etiqueta.ForeColor = obtenerColorEnfoque(False)
End Sub
