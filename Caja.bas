Attribute VB_Name = "Caja"
Public Function cajaAbierta(vfecha As Date) As Boolean
On Error Resume Next
Dim vidg, vsql As String

Dim vidUlt, vidAModificar As Long

Dim vfecha2 As Date


cajaAbierta = True


Exit Function

vsql = "select  max(idhasta) as c from t_cajacierre"
vidUlt = traerDatos2(vsql, "c", pathDBMySQL)

vsql = "select  fecha as c from bancosmovimientos where idbancosmovimientos =" + Str(vidUlt)
vfecha2 = traerDatos2(vsql, "c", pathDBMySQL)

If Not vfecha >= vfecha2 Then
    cajaAbierta = False
    MsgBox "La caja está cerrada para esta fecha", vbCritical, "Caja"
End If



'vsql = "select * from t_cajacierre where fecha = '" + strfechaMySQL(vfecha) + "'"
'
'vidg = traerDatos2(vsql, "id", pathDBMySQL)
'
'If Not vidg = "" Then
'    cajaAbierta = False
'    MsgBox "La caja está cerrada para esta fecha", vbCritical, "Caja"
'Else
'    cajaAbierta = True
'End If
'
'If Err Then
'    cajaAbierta = True
'    Exit Function
'End If
End Function


