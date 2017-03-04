VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ActivosDatosSucursal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Function aplicarPoliticaPrecioVenta(Precio As Integer, politica As Integer, putil As Double, venta As Double) As Double
    Select Case politica
        Case 1: aplicarPoliticaPrecioVenta = Nz(venta, 0)
        Case 2: aplicarPoliticaPrecioVenta = Me.VrCostoEstandar + (Me.VrCostoEstandar * (Me.Controls("PUtilidadPrecio" & Precio) / 100))
        Case 3: aplicarPoliticaPrecioVenta = Me.VrUltimoPrecioCompra + (Me.VrUltimoPrecioCompra * (Me.Controls("PUtilidadPrecio" & Precio) / 100))
        Case 4: aplicarPoliticaPrecioVenta = Me.VrCostoPromedio + (Me.VrCostoPromedio * (Me.Controls("PUtilidadPrecio" & Precio) / 100))
        Case 5: aplicarPoliticaPrecioVenta = Me.VrCosto1 + (Me.VrCosto1 * (Me.Controls("PUtilidadPrecio" & Precio) / 100))
        Case 6: aplicarPoliticaPrecioVenta = Me.VrCosto2 + (Me.VrCosto2 * (Me.Controls("PUtilidadPrecio" & Precio) / 100))
        Case 7: aplicarPoliticaPrecioVenta = Me.VrCosto3 + (Me.VrCosto3 * (Me.Controls("PUtilidadPrecio" & Precio) / 100))
    End Select
End Function

Private Sub CalcularVrVenta(Precio As Integer)
    Me.Controls("VentaVrBruto" & Precio) = aplicarPoliticaPrecioVenta(Precio, Me.Controls("PoliticaPrecioVenta" & Precio), Me.Controls("PUtilidadPrecio" & Precio), Me.Controls("VentaVrBruto" & Precio))
    Me.Controls("VentaVrBase" & Precio) = Me.Controls("VentaVrBruto" & Precio) - (Me.Controls("VentaVrBruto" & Precio) * Me.PDescuento1 / 100) - (Me.Controls("VentaVrBruto" & Precio) * Me.PDescuento2 / 100) - (Me.Controls("VentaVrBruto" & Precio) * Me.PDescuento3 / 100)
End Sub

Private Sub CalcularPUtilidad(Precio As Integer)
    Me.Controls("PUtilidadPrecio" & Precio) = (Me.VentaVrBruto1 - Me.VrCostoBruto) / (Me.VentaVrBruto1)
End Sub

Private Sub bloquearCamposPrecio(Precio As Integer, politica As Integer)
If (politica = 1) Then
    Me.Controls("VentaVrBase" & Precio).Enabled = True
    Me.Controls("VentaVrBruto" & Precio).Enabled = True
Else
    Me.Controls("VentaVrBase" & Precio).Enabled = False
    Me.Controls("VentaVrBruto" & Precio).Enabled = False
End If
End Sub

Private Sub actualizarPrecios()
    CalcularVrVenta 1
    CalcularVrVenta 2
    CalcularVrVenta 3
    CalcularVrVenta 4
    CalcularVrVenta 5
End Sub

Private Sub btnAnterior_Click()
On Error GoTo noRecord_Err
DoCmd.RunCommand acCmdRecordsGoToPrevious
Exit Sub
noRecord_Err:
    MsgBox "No hay más registros por mostrar"
    Exit Sub
End Sub

Private Sub btnGuardar_Click()
    If (Me.Guardado = 0 And Nz(DLookup("[IDActivo]", "ActivosDatosSucursal", "[IDActivo]=" & Me.IDActivo & " AND [IDSucursal]=" & Me.IDSucursal & " And [Indice]<>" & Me.Indice), 0) <> 0) Then
        MsgBox "Ya existen datos para este activo en esta sucursal"
        DoCmd.RunCommand acCmdRecordsGoToPrevious
        DoCmd.RunSQL "delete activosdatossucursal.* from activosdatossucursal where guardado =0" & Me.Indice
        DoCmd.RunCommand acCmdRefresh
        End
    End If
    Me.Guardado = 1
    DoCmd.RunCommand acCmdSave
    MsgBox "Guardado con éxito"
End Sub


Private Sub btnSiguiente_Click()
On Error GoTo noRecord_Err
If (Me.Guardado = 0 And Nz(DLookup("[IDActivo]", "ActivosDatosSucursal", "[IDActivo]=" & Me.IDActivo & " AND [IDSucursal]=" & Me.IDSucursal & " And [Indice]<>" & Me.Indice), 0) <> 0) Then
        MsgBox "Ya existen datos para este activo en esta sucursal"
        DoCmd.RunCommand acCmdRecordsGoToPrevious
        DoCmd.RunSQL "delete activosdatossucursal.* from activosdatossucursal where guardado =0" & Me.Indice
        DoCmd.RunCommand acCmdRefresh
        End
End If
DoCmd.RunCommand acCmdRecordsGoToNext
Exit Sub
noRecord_Err:
    MsgBox "No hay más registros por mostrar"
    Exit Sub
End Sub

Private Sub Comando1_Click()
DoCmd.Close
End Sub

Private Sub Form_Close()
DoCmd.RunSQL "delete activosdatossucursal.* from activosdatossucursal where guardado=0;", -1
End Sub

Private Sub Form_Current()
DoCmd.RunCommand acCmdRefresh
If (Me.IDActivo = 0) Then
    Me.IDActivo = getIDActivo
    DoCmd.RunCommand acCmdRefresh
    DoCmd.OpenQuery "Activos_ValoresPredeterminadosSucursal"
    DoCmd.RunCommand acCmdRefresh
End If
End Sub

Private Sub PDescuento1_AfterUpdate()
    actualizarPrecios
End Sub

Private Sub PDescuento2_AfterUpdate()
    actualizarPrecios
End Sub

Private Sub PDescuento3_AfterUpdate()
    actualizarPrecios
End Sub

Private Sub PUtilidadPrecio1_AfterUpdate()
    CalcularVrVenta 1
End Sub

Private Sub PUtilidadPrecio2_AfterUpdate()
    CalcularVrVenta 2
End Sub

Private Sub PUtilidadPrecio3_AfterUpdate()
    CalcularVrVenta 3
End Sub

Private Sub PUtilidadPrecio4_AfterUpdate()
    CalcularVrVenta 4
End Sub

Private Sub PUtilidadPrecio5_AfterUpdate()
    CalcularVrVenta 5
End Sub

Private Sub VentaVrBruto1_AfterUpdate()
    CalcularPUtilidad 1
End Sub

Private Sub VentaVrBruto2_AfterUpdate()
    CalcularPUtilidad 2
End Sub

Private Sub VentaVrBruto3_AfterUpdate()
    CalcularPUtilidad 3
End Sub

Private Sub VentaVrBruto4_AfterUpdate()
    CalcularPUtilidad 4
End Sub

Private Sub VentaVrBruto5_AfterUpdate()
    CalcularPUtilidad 5
End Sub

Private Sub VrCostoBruto_AfterUpdate()
    actualizarPrecios
End Sub

Private Sub PoliticaPrecioVenta1_AfterUpdate()
    CalcularVrVenta 1
    bloquearCamposPrecio 1, Me.PoliticaPrecioVenta1
End Sub

Private Sub PoliticaPrecioVenta2_AfterUpdate()
    CalcularVrVenta 2
    bloquearCamposPrecio 2, Me.PoliticaPrecioVenta2
End Sub

Private Sub PoliticaPrecioVenta3_AfterUpdate()
    CalcularVrVenta 3
    bloquearCamposPrecio 3, Me.PoliticaPrecioVenta3
End Sub

Private Sub PoliticaPrecioVenta4_AfterUpdate()
    CalcularVrVenta 4
    bloquearCamposPrecio 4, Me.PoliticaPrecioVenta4
End Sub

Private Sub PoliticaPrecioVenta5_AfterUpdate()
    CalcularVrVenta 5
    bloquearCamposPrecio 5, Me.PoliticaPrecioVenta2
End Sub

Private Sub VrCosto1_AfterUpdate()
    actualizarPrecios
End Sub

Private Sub VrCosto2_AfterUpdate()
    actualizarPrecios
End Sub

Private Sub VrCosto3_AfterUpdate()
    actualizarPrecios
End Sub

Private Sub VrCostoEstandar_AfterUpdate()
    actualizarPrecios
End Sub


