VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FiltroBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Function agregarFiltro(sql As String, field As String, value As String, isInteger As Boolean) As String
    If (Len(value) > 0) Then
        If (isInteger) Then
            agregarFiltro = sql & " And " & field & "=" & value
        Else
            agregarFiltro = sql & " And " & field & "='" & value & "'"
        End If
        
    Else
        agregarFiltro = sql
    End If
End Function

Private Function agregarFiltroEntreDatos(sql As String, field As String, value1 As String, value2 As String) As String
    If (Len(value1) > 0 And Len(value2) > 0) Then
        agregarFiltroEntreDatos = sql & " And " & field & " Between " & value1 & " And " & value2
    Else
        agregarFiltroEntreDatos = sql
    End If
End Function

Private Function agregarFiltroEntreFechas(sql As String, field As String, fecha1 As String, fecha2 As String) As String
    If (Len(fecha1) > 0 And Len(fecha2) > 0) Then
        agregarFiltroEntreFechas = sql & " And " & field & " Between #" & fecha1 & "# And #" & fecha2 & "#"
    Else
        agregarFiltroEntreFechas = sql
    End If
End Function

Private Function agregarFiltroLike(sql As String, field As String, value As String)
    If (Len(value) > 0) Then
        agregarFiltroLike = sql & " And " & field & " Like '*' & " & value & " & '*'"
    Else
        agregarFiltroLike = sql
    End If
End Function

Private Sub aplicarFiltro()

Dim sql As String

sql = "SELECT IDDocumentoRef, NombreComprobante, Fecha, FechaRegistro, NoDocumento, Observacion, NombreTercero," & _
"Sum(MovimientoContable_v.debito) AS Debito, Sum(MovimientoContable_v.credito) AS Credito, Sum(MovimientoContable_v.debitoniif) AS DebitoNIIF, Sum(MovimientoContable_v.creditoniif) AS CreditoNIIF," & _
"MovLocal, MovNIIF, Conciliado, Cerrado, Cancelado, IDEmpresa, Contabilizar," & _
"SubDatoExtra1, SubDatoExtra2, SubDatoExtra3, SubDatoExtra4 FROM MovimientoContable_v WHERE IDDocumentoRef<>null"

sql = agregarFiltro(sql, "Nro", Nz(Me.DocRef.value, ""), False)
sql = agregarFiltro(sql, "IDNIT", Nz(Me.NIT.value, ""), True)
sql = agregarFiltro(sql, "IDProceso", Nz(Me.PROCESO.value, ""), True)
sql = agregarFiltro(sql, "IDUsuario", Nz(Me.Usuario.value, ""), True)
sql = agregarFiltro(sql, "IDPC", Nz(Me.PC.value, ""), True)
sql = agregarFiltro(sql, "IDComprobante", Nz(Me.Comprobante.value, ""), True)
sql = agregarFiltroEntreDatos(sql, "CodigoCuentaPUC", Nz(Me.PUCInicio.value, ""), Nz(Me.PUCFin.value, ""))
sql = agregarFiltro(sql, "IDActivo", Nz(Me.Activo.value, ""), True)

sql = sql & " GROUP BY IDDocumentoRef, IDComprobante, NombreComprobante, Fecha, FechaRegistro, NoDocumento, " & _
"Observacion, NombreTercero, MovLocal, MovNIIF, Conciliado, Cerrado, Cancelado, IDEmpresa, Contabilizar, " & _
" SubDatoExtra1, SubDatoExtra2, SubDatoExtra3, SubDatoExtra4 HAVING IDDocumentoRef<>null "

sql = agregarFiltroEntreFechas(sql, "Fecha", Me.fechaInicio.value, Me.fechaFin.value)
sql = agregarFiltro(sql, "NoDocumento", Nz(Me.NoComprobante.value, ""), False)
sql = agregarFiltroLike(sql, "Observacion", Nz(Me.Observacion.value, ""))
sql = agregarFiltro(sql, "MovLocal", Nz(Me.MovLocal.value, ""), True)
sql = agregarFiltro(sql, "MovNIIF", Nz(Me.MovNIIF.value, ""), True)
sql = agregarFiltro(sql, "Conciliado", Nz(Me.Conciliado.value, ""), True)
sql = agregarFiltro(sql, "IDEmpresa", Nz(Me.empresa.value, ""), True)
sql = agregarFiltro(sql, "Contabilizar", Nz(Me.Contabilizado.value, ""), True)
sql = agregarFiltro(sql, "IDSucursal", Nz(Me.Sucursal.value, ""), True)
sql = agregarFiltro(sql, "IDCDC", Nz(Me.CDC.value, ""), True)
sql = agregarFiltro(sql, "NoCheque", Nz(Me.NoCheque.value, ""), False)
sql = agregarFiltro(sql, "Lote", Nz(Me.Lote.value, ""), True)
sql = agregarFiltro(sql, "Bodega", Nz(Me.Bodega.value, ""), True)
sql = agregarFiltro(sql, "SubDatoExtra1", Nz(Me.Dato1.value, ""), False)
sql = agregarFiltro(sql, "SubDatoExtra2", Nz(Me.Dato2.value, ""), False)
sql = agregarFiltro(sql, "SubDatoExtra3", Nz(Me.Dato3.value, ""), False)
sql = agregarFiltro(sql, "SubDatoExtra4", Nz(Me.Dato4.value, ""), False)
sql = sql & ";"

Forms!CMovimientoContable!CMovimientoContableSub.Form.RecordSource = sql
Forms!CMovimientoContable!CMovimientoContableSub.Requery
    
End Sub


Private Sub btnGuardar_Click()
    aplicarFiltro
End Sub

Private Sub Comando19_Click()
    DoCmd.Close
End Sub
