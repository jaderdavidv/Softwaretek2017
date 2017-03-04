Attribute VB_Name = "Modulo_Seguridad"
Option Compare Database
Option Explicit

Public Function TeclaShift(strPropName As String, _
        varPropType As Variant, varPropValue As Variant) As Integer
        Dim dbs As DAO.Database, prp As Property
    Const conPropNotFoundError = 3270
    Set dbs = CurrentDb
    On Error GoTo Change_Err
    dbs.Properties(strPropName) = varPropValue
    TeclaShift = True
Change_Bye:
    Exit Function
Change_Err:
    If Err = conPropNotFoundError Then
        Set prp = dbs.CreateProperty(strPropName, varPropType, varPropValue)
        dbs.Properties.Append prp
        Resume Next
    Else
        TeclaShift = False
        Resume Change_Bye
    End If
End Function

