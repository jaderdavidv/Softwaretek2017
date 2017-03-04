Attribute VB_Name = "Modulo_Activos"
Option Compare Database

Public Function obtenerSaldoActivoSucursal(IDActivo As Integer, IDSucursal As Integer) As Double
    Dim saldo As Double, saldoconversion As Double, tarifaconversion As Double
    saldo = Nz(DLookup("[dbsaldo]", "activosdatossucursal", "[idactivo]=" & IDActivo & " And [idsucursal]=" & IDSucursal), 0)
    obtenerSaldoActivoSucursal = saldo
End Function

Public Function calcularVrBrutoActivoSegunPromedio(promedio As Double, putilidad As Double) As Double
    calcularVrBrutoActivoSegunPromedio = Round(promedio + (promedio * putilidad / 100))
End Function

Public Function calcularVrBaseActivo(VrBruto As Double, ptotaldesc As Double) As Double
    Dim valor As Double
    Dim redondearCententa As Double
    valor = Round(VrBruto - (VrBruto * ptotaldesc / 100))
    redondearCentena = (100 - Right(valor, 2))  'Vr necesario para redondear el vr a multiplo de cien para efecto de valores reales de venta
    If (redondearCentena > 0 And redondearCentena < 100) Then
        valor = valor + redondearCentena
    End If
    calcularVrBaseActivo = valor
End Function

' Actualizar datos del activo al actualizar o guardar un movimiento

Public Sub ActualizarDatosActivoSucursal(afectaCostos As Boolean, MueveInventarios As Boolean)

Dim IDSucursal As Integer
Dim IDActivo As Integer
Dim index As Integer
Dim promedio As Double
Dim saldo As Double
Dim saldoInv As Double
Dim PDescuento As Double

Dim db As DAO.Database
Dim rcs As DAO.Recordset

On Error GoTo Funcion_Err

    ' Traemos los activos del movimiento a una tabla temporal para saber cuales vamos a actualizar
    Set db = CurrentDb
    DoCmd.RunSQL "Delete Temp_ActivosActualizarDatosSucursal.* from Temp_ActivosActualizarDatosSucursal;", -1
    DoCmd.OpenQuery "Temp_ActivosActualizarDatosSucursal01"
    Set rcs = db.OpenRecordset("SELECT * FROM Temp_ActivosActualizarDatosSucursal")
     
    Do Until rcs.EOF
    
        IDSucursal = rcs!IDSucursal
        IDActivo = rcs!IDActivo
        
        If (MueveInventarios) Then
            ActualizarSaldoActivosSucursal IDActivo, IDSucursal
        End If
        
        If (afectaCostos) Then
            promedio = Nz(DLookup("[CostoPromedio]", "Activos_CostoPromedio", "[IDSucursal]=" & IDSucursal & " And [IDActivo]=" & IDActivo), 0)
            ActualizarCostoPromedioActivo IDSucursal, IDActivo, promedio
            ActualizarVrVentaActivos IDSucursal, IDActivo
        End If
        
        rcs.MoveNext
    Loop

Finalizar:
    rcs.Close
    db.Close
    Set rcs = Nothing
    Set db = Nothing
    Exit Sub

Funcion_Err:
    MsgBox Err.Description, vbCritical, "Error en la función Actualizar Costo Promedio ActivoSucursal"
    Resume Finalizar

End Sub

' Actualiza el costo promedio del activo
Private Sub ActualizarCostoPromedioActivo(IDSucursal As Integer, IDActivo As Integer, promedio As Double)

On Error GoTo Funcion_Err
DoCmd.RunSQL "update ActivosDatosSucursal set VrCostoPromedio =" & reemplazarSeparadorDecimal(promedio) & ", UPDATEFIX = UPDATEFIX+1 where IDSucursal=" & IDSucursal & " And IDActivo=" & IDActivo & ";", -1

Finalizar:
    Exit Sub

Funcion_Err:
    MsgBox "Ocurrio un error al momento de actualizar el costo promedio del activo: " & DLookup("[nombreactivo]", "activos_v", "[idactivo]=" & IDActivo), vbCritical, Err.Description
    Resume Finalizar
End Sub

' Actualizar el saldo de los activos

Private Sub ActualizarSaldoActivosSucursal(IDActivo As Integer, IDSucursal As Integer)

Dim saldo As Double
Dim saldoInv As Double

On Error GoTo Funcion_Err

saldo = Nz(DLookup("[saldo]", "activos_dbsaldo", "[IDActivo]=" & IDActivo & " And [IDSucursal]=" & IDSucursal), 0)
saldoInv = Nz(DLookup("[SALDOINV]", "activos_dbsaldo", "[IDActivo]=" & IDActivo & " And [IDSucursal]=" & IDSucursal), 0)
DoCmd.RunSQL "update ActivosDatosSucursal set dbsaldo =" & saldo & ", dbsaldoinv=" & saldoInv & ", UPDATEFIX = UPDATEFIX+1 where IDSucursal=" & IDSucursal & " And IDActivo=" & IDActivo & ";", -1

Finalizar:
    Exit Sub

Funcion_Err:
    MsgBox "Ocurrio un error al momento de actualizar el saldo del activo: " & DLookup("[nombreactivo]", "activos_v", "[idactivo]=" & IDActivo), vbCritical, Err.Description
    Resume Finalizar
End Sub

' Actualiza los valores de venta de acuerdo al valor promedio

Private Sub ActualizarVrVentaActivos(IDSucursal As Integer, IDActivo As Integer)

Dim db As DAO.Database
Dim rcs As DAO.Recordset
Dim VrBruto As Double
Dim VrBase As Double
Dim ptotaldesc As Double
Dim promedio As Double
Dim utilidad As Double

On Error GoTo Funcion_Err

Set db = CurrentDb
Set rcs = db.OpenRecordset("SELECT * FROM activosdatossucursal where IDSucursal=" & IDSucursal & " And IDActivo=" & IDActivo)

If (rcs.EOF) Then
    Resume Finalizar
End If

ptotaldesc = rcs!PDescuento1 + rcs!PDescuento2 + rcs!PDescuento3
promedio = rcs!VrCostoPromedio

Dim i As Double

For i = 1 To 5
    utilidad = rcs.Fields("PUtilidadPrecio" & i)
    VrBruto = calcularVrBrutoActivoSegunPromedio(promedio, utilidad)
    VrBase = calcularVrBaseActivo(VrBruto, ptotaldesc)
    DoCmd.RunSQL "update ActivosDatosSucursal set VentaVrBruto" & i & "=" & VrBruto & ",ventavrbase" & i & "=" & VrBase & ",UPDATEFIX = UPDATEFIX+1 where IDSucursal=" & IDSucursal & " And IDActivo=" & IDActivo & ";", -1
Next i

Finalizar:
    rcs.Close
    db.Close
    Set rcs = Nothing
    Set db = Nothing
    Exit Sub

Funcion_Err:
    MsgBox "Ocurrio un error al momento de actualizar el vrventa del activo: " & DLookup("[nombreactivo]", "activos_v", "[idactivo]=" & IDActivo), vbCritical, Err.Description
    Resume Finalizar
    
End Sub



