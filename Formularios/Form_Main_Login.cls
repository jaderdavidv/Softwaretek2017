VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Main_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim Color As String

Public Function getFocusColor(focus As Boolean) As Long
If (focus) Then

    Color = "#F3A647"
    getFocusColor = HexToLong((Right(Color, Len(Color) - 1)))
Else
    Color = "#A29D96"
    getFocusColor = HexToLong((Right(Color, Len(Color) - 1)))
End If
End Function

Private Sub CodigoUsuario_AfterUpdate()
    If (DCount("[nombreusuario]", "main_usuarios_v", "[codigousuario]=" & Me.CodigoUsuario) > 0) Then
        Me.nombreusuario = DLookup("[nombreusuario]", "main_usuarios_v", "[codigousuario]=" & Me.CodigoUsuario)
    Else
        MsgBox "No hay algun usuario registrado con este c�digo."
        DoCmd.GoToControl "CodigoUsuario"
        Me.nombreusuario = " "
    End If
End Sub

Private Sub CodigoUsuario_GotFocus()
Me.Usuario_Etiqueta.ForeColor = getFocusColor(True)
End Sub

Private Sub CodigoUsuario_LostFocus()
Me.Usuario_Etiqueta.ForeColor = getFocusColor(False)
End Sub

Private Sub contrase�a_GotFocus()
Me.Contrase�a_Etiqueta.ForeColor = getFocusColor(True)
End Sub

Private Sub contrase�a_LostFocus()
Me.Contrase�a_Etiqueta.ForeColor = getFocusColor(False)
End Sub

Private Sub Login_Click()
    If (Len(Me.CodigoUsuario) > 0 And Len(Me.Contrase�a) > 0) Then
        Dim pass As String
        Dim auth As Integer
        pass = Encriptar(Me.Contrase�a)
        auth = VerificarUsuario(Me.CodigoUsuario, pass)
        If (auth = 1) Then
            MsgBox "AUTENTICADO CON EXITO"
            setUsuarioActual Me.CodigoUsuario
            ConfigurarSistemaUsuario
            DoCmd.Close acForm, "main_login"
            DoCmd.OpenForm "Principal"
        Else
            MsgBox "USUARIO O CONTRASE�A INCORRECTOS"
        End If
    Else
        MsgBox "Favor Ingresar Usuario y Contrase�a"
        End
    End If
End Sub
