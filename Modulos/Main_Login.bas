Attribute VB_Name = "Main_Login"
Option Compare Database
Option Explicit

Dim usuarioActual As Integer

Public Function getUsuarioActual() As Integer
    getUsuarioActual = usuarioActual
End Function

Public Function setUsuarioActual(COD As Integer)
    usuarioActual = COD
End Function

'//////////////////////////////////////////////////////////////////////////////////////////////////////////

Public Function CrearUsuario(Codigo As Integer, nombre As String, pass As String) As Integer
    Dim sql As String
    sql = "SELECT CrearUsuario(" & Codigo & ",'" & nombre & "','" & pass & "') crea"
    CrearUsuario = EjecutarSQL("main_usuarios_v", sql)
End Function


'//////////////////////////////////////////////////////////////////////////////////////////////////////////

Public Function VerificarUsuario(COD As Integer, pass As String) As Integer
    Dim sql As String
    sql = "SELECT Login(" & COD & ",'" & pass & "') verifica"
    VerificarUsuario = EjecutarSQL("main_usuarios_v", sql)
End Function

'//////////////////////////////////////////////////////////////////////////////////////

Public Function ConfigurarSistemaUsuario()

    Dim Ribbon As Integer
    Dim Shift As Integer
    
    Ribbon = DLookup("[Ribbon]", "main_usuarios_v", "[id_usuario]=" & getUsuarioActual)
    Shift = DLookup("[Shift]", "main_usuarios_v", "[id_usuario]=" & getUsuarioActual)

    If (Shift = -1) Then
            TeclaShift "AllowBypassKey", dbBoolean, True
    End If
    
    If (Shift = 0) Then
            TeclaShift "AllowBypassKey", dbBoolean, False
    End If
    
    If (Ribbon = -1) Then
            DoCmd.ShowToolbar "Ribbon", acToolbarYes
    End If
    
    If (Ribbon = 0) Then
            DoCmd.ShowToolbar "Ribbon", acToolbarNo
    End If
    

End Function

'//////////////////////////////////////////////////////////////////////////////////////

Public Function EjecutarSQL(TablaConexion As String, cadenaSQL As String) As Integer

    Dim CON As New ADODB.Connection
    Dim rec As New ADODB.Recordset
    Dim constr As String
    Dim recstr As String
    
    constr = Mid(CurrentDb(TablaConexion).Connect, 6)
    recstr = cadenaSQL
    
    CON.Open constr
    rec.Open recstr, CON
    
    If Not rec.BOF And Not rec.EOF Then
        EjecutarSQL = rec(0)
    End If
    
    rec.Close
    CON.Close
    
    Set rec = Nothing
    Set CON = Nothing
End Function
