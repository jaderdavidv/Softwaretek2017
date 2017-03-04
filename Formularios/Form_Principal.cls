VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub AsignarIDComprobanteADocumento(IDComprobante As Integer)
    Me.SubformularioDeNavegación.Form!IDComprobante = IDComprobante
    Me.SubformularioDeNavegación.Form.Requery
End Sub

Private Sub BotónDeNavegación312_Click()
    DoCmd.RunSQL "delete documentoencabezado.* from documentoencabezado;", -1
    DoCmd.RunSQL "delete MovimientoContable.* from movimientoContable", -1
    DoCmd.RunSQL "delete documentosub.* from documentosub", -1
End Sub

Private Sub BotónDeNavegaciónCompras_Click()
    AsignarIDComprobanteADocumento (2)
End Sub

Private Sub BotónDeNavegaciónEgresos_Click()
    AsignarIDComprobanteADocumento (3)
End Sub

Private Sub BotónDeNavegaciónOtrosIngresos_Click()
    AsignarIDComprobanteADocumento (4)
End Sub

Private Sub BotónDeNavegaciónVentas_Click()
    AsignarIDComprobanteADocumento (1)
End Sub

Private Sub Form_Load()
'    If (getUsuarioActual > 0) Then
'    Else
'        MsgBox "No te has autenticado, favor iniciar el programa nuevamente"
'        DoCmd.RunMacro "SalirDeAccess"
'   End If
End Sub
