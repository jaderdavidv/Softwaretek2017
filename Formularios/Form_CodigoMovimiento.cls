VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_CodigoMovimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnGuardar_Click()
    Me.UpdateFIX = Me.UpdateFIX + 1
    DoCmd.RunCommand acCmdSave
End Sub


Private Sub CodigoMovimiento_GotFocus()
    Me.CodigoMovimiento_Etiqueta.ForeColor = obtenerColorEnfoque(True)
End Sub

Private Sub CodigoMovimiento_LostFocus()
    Me.CodigoMovimiento_Etiqueta.ForeColor = obtenerColorEnfoque(False)
End Sub

Private Sub ComprobanteAsociado_GotFocus()
    Me.ComprobanteAsociado_Etiqueta.ForeColor = obtenerColorEnfoque(True)
End Sub

Private Sub ComprobanteAsociado_LostFocus()
    Me.ComprobanteAsociado_Etiqueta.ForeColor = obtenerColorEnfoque(False)
End Sub

Private Sub Favorito_GotFocus()
    Me.Favorito_Etiqueta.ForeColor = obtenerColorEnfoque(True)
End Sub

Private Sub Favorito_LostFocus()
    Me.Favorito_Etiqueta.ForeColor = obtenerColorEnfoque(False)
End Sub

Private Sub Habilitado_GotFocus()
    Me.Habilitado_Etiqueta.ForeColor = obtenerColorEnfoque(True)
End Sub


Private Sub Habilitado_LostFocus()
    Me.Habilitado_Etiqueta.ForeColor = obtenerColorEnfoque(False)
End Sub

Private Sub IDCuentaPUC_AfterUpdate()
    If (Me.IDCuentaPUCNIIF = 0) Then
        Me.IDCuentaPUCNIIF = Me.IDCuentaPUC
    End If
End Sub

Private Sub IDCuentaPUC_GotFocus()
    Me.IDCuentaPUC_Etiqueta.ForeColor = obtenerColorEnfoque(True)
End Sub

Private Sub IDCuentaPUC_LostFocus()
    Me.IDCuentaPUC_Etiqueta.ForeColor = obtenerColorEnfoque(False)
End Sub

Private Sub IDCuentaPUCNIIF_GotFocus()
    Me.IDCuentaPUCNIIF_Etiqueta.ForeColor = obtenerColorEnfoque(True)
End Sub

Private Sub IDCuentaPUCNIIF_LostFocus()
    Me.IDCuentaPUCNIIF_Etiqueta.ForeColor = obtenerColorEnfoque(False)
End Sub

Private Sub IDGrupoMovimiento_AfterUpdate()
    Me.IDSubGrupoMovimiento = DLookup("[idsubgrupocodigomovimiento]", "codigomovimiento_subgrupo", "[idgrupocodigomovimiento]=" & Me.IDGrupoMovimiento)
    Me.IDSubGrupoMovimiento.Requery
End Sub

Private Sub IDSubGrupoMovimiento_GotFocus()
    Me.IDSubGrupoMovimiento_Etiqueta.ForeColor = obtenerColorEnfoque(True)
End Sub

Private Sub IDSubGrupoMovimiento_LostFocus()
    Me.IDSubGrupoMovimiento_Etiqueta.ForeColor = obtenerColorEnfoque(False)
End Sub

Private Sub NombreMovimiento_GotFocus()
    Me.NombreMovimiento_Etiqueta.ForeColor = obtenerColorEnfoque(True)
End Sub

Private Sub NombreMovimiento_LostFocus()
    Me.NombreMovimiento_Etiqueta.ForeColor = obtenerColorEnfoque(False)
End Sub

Private Sub Observacion_GotFocus()
    Me.Observacion_Etiqueta.ForeColor = obtenerColorEnfoque(True)
End Sub

Private Sub Observacion_LostFocus()
    Me.Observacion_Etiqueta.ForeColor = obtenerColorEnfoque(False)
End Sub

