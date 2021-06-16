Attribute VB_Name = "Balance"
Public viddesde, vidhasta As Long


Public Function informeDelCierre(ByVal vidd As Long, ByVal vidH As Long) As String
On Error Resume Next
Dim vsql As String
Dim vfecha1, vfecha2, vmen As String
Dim vfechaauxi As Date

vsql = "select fecha as c from asientos where idAsientos=" + Str(vidd)
vfechaauxi = traerDatos2(vsql, "c", pathDBMySQL)
'vfecha1 = traerDatos2(vsql, "c", pathDBMySQL)

vfecha1 = Format(vfechaauxi + 1, "dd/mm/yyyy")

vsql = "select fecha as c from asientos where idAsientos=" + Str(vidH)
vfecha2 = traerDatos2(vsql, "c", pathDBMySQL)

vmen = "Se están cerrando los movimientos del balance de ejecución del siguiente período: " + _
Chr(13) + " > Fecha desde: " + vfecha1 + _
Chr(13) + " > Fecha hasta: " + vfecha2

MsgBox vmen, vbInformation

informeDelCierre = " Fecha desde: " + vfecha1 + "  /   Fecha hasta: " + vfecha2
If Err Then
    MsgBox "No se puede mostrar el balance. Consulte servicio técnico"
    Exit Function
End If

End Function


Public Function selectNrobalance(vfdesde As Date, vfhasta As Date, vnrobalance1 As Integer) As Integer
' selecciona automaticamente el nro de balance según la fecha seleccionada

Dim vsql, vcodigo  As String
Dim vperiodos, vnrobalance As Integer


vsql = "select count(activo) as c from balances where fechainicio<='" + strfechaMySQL(vfdesde) + "' and   fechafin>='" + strfechaMySQL(vfhasta) + "' group by activo"
vperiodos = Val(EsNulo(traerDatos2(vsql, "c", pathDBMySQL)))


If vperiodos = 0 Then
    MsgBox "El rango de fecha  ingresado corresponde a dos período diferente.", vbInformation, "Opción no permitida"
    vnrobalance = 0
    selectNrobalance = 0
    Exit Function
End If


vsql = "select nrobalance as c from balances where fechainicio<='" + strfechaMySQL(vfdesde) + "' and fechafin >='" + strfechaMySQL(vfhasta) + "' order by nrobalance desc"
vnrobalance = traerDatos2(vsql, "c", pathDBMySQL)

vcodigo = traerDatos2("select * from balances where nrobalance = " + Str(vnrobalance), "codigo", pathDBMySQL)

If Not vnrobalance = vnrobalance1 Then
    MsgBox "Las fechas ingresadas corresponden al período de balance con código: " + (vcodigo), vbInformation, "Advertencia.."
End If

selectNrobalance = vnrobalance

End Function


Function sqlBalanceIE(vd As Date, vh As Date, vi As Boolean, ve As Boolean) As String
Dim vsqlIE, vsqlIE2 As String


' discrimino si es ingreso o egreso
If vi Then
    vsqlIE = " and left(asientosdetalle.CodigoCuenta,2) = '01' "
    vsqlIE2 = " where left(CodigoCuenta,2) = '01' "
    
End If


If ve Then
    vsqlIE = " and left(asientosdetalle.CodigoCuenta,2) = '02' "
    vsqlIE2 = " where left(CodigoCuenta,2) = '02' "
End If

'sqlBalanceIE = " SHAPE {select distinct * from cuentas " + vsqlIE2 + " and imputable = 'S'  group by CodigoCuenta }  AS balanceIE APPEND ({select " + _



sqlBalanceIE = " SHAPE {select distinct * from cuentas " + vsqlIE2 + " and imputable = 'S'  group by CodigoCuenta }  AS balanceIE APPEND ({select " + _
" asientosdetalle.CodigoCuenta, " + _
" asientosdetalle.Debe  as Debe, " + _
" asientosdetalle.haber as Haber " + _
", asientosdetalle.LeyendaBancoCaja, asientos.fecha " + _
" From " + _
" asientos " + _
" inner join asientosdetalle " + _
" on asientos.numero = asientosdetalle.numero " + _
" Where   asientos.nrobalance = asientosdetalle.nrobalance and asientos.fecha >= '" + strfechaMySQL(vd) + "' and asientos.fecha <= '" + strfechaMySQL(vh) + "' " + _
vsqlIE + _
" }  AS biedetalle RELATE 'CodigoCuenta' TO 'CodigoCuenta') AS biedetalle "


' " Where   asientos.nrobalance = asientosdetalle.nrobalance and asientos.fecha >= '" + strfechaMySQL(vd) + "' and asientos.fecha <= '" + strfechaMySQL(vh) + "' " + _


End Function
