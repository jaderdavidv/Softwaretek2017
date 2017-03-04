VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Comprobante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Function generarResolucionDIAN(no As String, Fecha As Date, no1 As Integer, no2 As Integer, Prefijo As String, tipo As String) As String
    generarResolucionDIAN = "Resolución DIAN No " & no & " del " & Fecha & " " & tipo & " del " & Prefijo & no1 & " hasta el " & Prefijo & no2
End Function

Private Sub camposInventarioVisibles(val As Boolean)
    Me.ManejaIDActivo.Visible = val
    Me.ManejaEntradas.Visible = val
    Me.ManejaSalidas.Visible = val
    Me.ManejaListaPrecios.Visible = val
    Me.MueveValor.Visible = val
    Me.AfectaCostoPromedio.Visible = val
End Sub

Private Sub deshabilitarTodoMueveInventarios()

    Me.ManejaIDActivo = False
    Me.ManejaEntradas = False
    Me.ManejaSalidas = False
    Me.ManejaListaPrecios = False
    Me.ManejaCantidadReferencia = False
    Me.ManejaColor = False
    Me.ManejaTalla = False
    Me.ManejaLote = False
    Me.ManejaUnidad = False
    Me.ManejaVidaUtil = False
    Me.AfectaCostoPromedio = False
    
End Sub

Private Sub ResolucionDianVisible(mostrar As Boolean)
    Me.ResolucionDIAN.Visible = mostrar
    Me.ResolucionDocFin.Visible = mostrar
    Me.ResolucionDocInicio.Visible = mostrar
    Me.ResolucionFechaFin.Visible = mostrar
    Me.ResolucionFechaInicio.Visible = mostrar
    Me.ResolucionNo.Visible = mostrar
    Me.Prefijo.Visible = mostrar
    Me.TipoResolucion.Visible = mostrar
End Sub


Private Sub btnGuardar_Click()
    DoCmd.RunCommand acCmdRefresh
    Me.UpdateFIX = Me.UpdateFIX + 1
    DoCmd.RunCommand acCmdSave
End Sub

Private Sub btnGuardarPendiente_Click()
DoCmd.OpenForm "comprobante_configuracion", acNormal, , "[idcomprobante]=" & Me.IDComprobante
End Sub

Private Sub Comando219_Click()
DoCmd.OpenForm "Comprobante_ConfigImpuestos", acNormal, , "[idcomprobante]=" & Me.IDComprobante
End Sub

Private Sub Comando220_Click()

If (Me.ResolucionNo > 0) Then
Else
    MsgBox "Favor Ingresar Número de Resolución Válido"
    GoTo ResolucionInvalida
End If

If (Len(Me.ResolucionFechaInicio) > 0) Then
Else
    MsgBox "Favor Ingresar Fecha de Resolución Válida"
    GoTo ResolucionInvalida
End If

If (Len(Me.ResolucionFechaFin) > 0) Then
Else
    MsgBox "Favor Ingresar Fecha de Resolución Válida"
    GoTo ResolucionInvalida
End If

If (Me.ResolucionDocInicio >= 0) Then
Else
    MsgBox "Favor Ingresar Consecutivo Inicial Válido"
    GoTo ResolucionInvalida
End If

If (Me.ResolucionDocFin > Me.ResolucionDocInicio) Then
Else
    MsgBox "Favor Ingresar Consecutivo Final Válido"
    GoTo ResolucionInvalida
End If

Me.ResolucionDIAN = generarResolucionDIAN(Me.ResolucionNo, Me.ResolucionFechaInicio, Me.ResolucionDocInicio, Me.ResolucionDocFin, Nz(Me.Prefijo, ""), Me.TipoResolucion)
End

ResolucionInvalida:
    Me.ResolucionDIAN = ""

End Sub

Private Sub Form_Load()
    ResolucionDianVisible (Me.ManejaResolucion)
    camposInventarioVisibles (Me.MueveInventarios)
    Me.UpdateFIX = Me.UpdateFIX + 1
End Sub

Private Sub Grupo_AfterUpdate()
Me.SubGrupo.Requery
End Sub

Private Sub ManejaResolucion_AfterUpdate()
    ResolucionDianVisible (Me.ManejaResolucion)
End Sub

Private Sub MueveInventarios_AfterUpdate()

camposInventarioVisibles (Me.MueveInventarios)

    If (Me.MueveInventarios = False) Then
        deshabilitarTodoMueveInventarios
    End If
End Sub

