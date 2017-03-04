Attribute VB_Name = "Main_Funciones"
Option Compare Database

Public Function ObtenerNuevoRegedit() As Long
    '
    Dim CON As New ADODB.Connection
    Dim rec As New ADODB.Recordset
    Dim constr As String
    Dim recstr As String
    '
    constr = Mid(CurrentDb("main_regedit").Connect, 6)
    recstr = "SELECT ObtenerNuevoRegedit() newId"
    '
    CON.Open constr
    rec.Open recstr, CON
    '
    If Not rec.BOF And Not rec.EOF Then
        '
        ObtenerNuevoRegedit = rec(0)
        '
    End If
    '
    rec.Close
    CON.Close
    '
    Set rec = Nothing
    Set CON = Nothing
    '
End Function


' Conversion de 0 y 1 a true y false

Public Function getBoolean(i As Integer) As Boolean
    If (i = 0) Then
        getBoolean = False
    Else
        getBoolean = True
    End If
End Function



