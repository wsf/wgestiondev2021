Attribute VB_Name = "Control_Tablero"
Function getSaldosCaja(vfecha As Date)
On Error Resume Next

Dim vsql As String


vsql = " SELECT sum(`bm`.`Debito`) - sum(`bm`.`Credito`) AS c " + _
       " From `bancos` " + _
        " INNER JOIN `bancosmovimientos` `bm` ON (`bancos`.`idBancos` = `bm`.`idBancos`) " + _
        " Where " + _
        " `bancos`.`tipodisponibilidad` = 'Disponible' "

getSaldosCaja = traerDatos2(vsql, "c", pathDBMySQL)

If Err Then Exit Function
End Function


Function getSaldosCajaDetalle(vfecha As Date)
On Error Resume Next

Dim vsql As String


vsql = "SELECT bancos.idBancos, Descripcion, format(sum(Debito) - sum(Credito),'###,###,##0.00') AS Saldo From bancos INNER JOIN bancosmovimientos bm ON (bancos.idBancos = bm.idBancos) Where tipodisponibilidad = 'Disponible' group by bancos.idBancos"

getSaldosCajaDetalle = vsql

If Err Then Exit Function
End Function



Function getSaldoProveedor(ByVal vfecha As Date) As Double
On Error Resume Next

Dim vsql As String

vsql = " SELECT  sum(`t`.`debito`) - sum(`t`.`Credito`) AS c From `pcuentascorrientes` `t` " + _
" where fecha <= '" + strfechaMySQL(vfecha) + "'"

getSaldoProveedor = traerDatos2(vsql, "c", pathDBMySQL)

If Err Then Exit Function
End Function


Function getSaldoProveedorDetalle(ByVal vfecha As Date, vb As String)
On Error Resume Next

Dim vsql As String

vsql = " SELECT  codigo, Nombre, format(sum(`t`.`debito`) - sum(`t`.`Credito`),'###,###,##0.00') AS Saldo From `pcuentascorrientes` `t` " + _
" where fecha <= '" + strfechaMySQL(vfecha) + "' and saldo > 0.1 and nombre like '%" + vb + "%' group by codigo order by nombre"

getSaldoProveedorDetalle = vsql

If Err Then Exit Function
End Function






Function GetDeudasVencidas(vfecha As Date)
On Error Resume Next

Dim vsql As String


vsql = " SELECT  sum(`t`.`debito`) - sum(`t`.`Credito`) AS c From `pcuentascorrientes` `t` " + _
" where fecha <= '" + strfechaMySQL(vfecha) + "' or credito > 0 "


GetDeudasVencidas = traerDatos2(vsql, "c", pathDBMySQL)

If Err Then Exit Function
End Function


Function GetDeudasVencidasDetalles(vfecha As Date, vb As String)
On Error Resume Next

Dim vsql As String


vsql = " SELECT  codigo, nombre, format(sum(`t`.`debito`) - sum(`t`.`Credito`),'###,###,##0.00') AS saldo  From `pcuentascorrientes` `t` " + _
" where credito > 0 or fecha <= '" + strfechaMySQL(vfecha - 90) + "' and saldo > 0.1 and nombre like '%" + vb + "%' group by codigo "


GetDeudasVencidasDetalles = vsql

If Err Then Exit Function
End Function


Function getCobrosUrbano(vfecha As Date) As Double
On Error Resume Next

Dim vsql As String


vsql = " SELECT  sum(`t`.`importe_total2`) AS c From   `recibo_resumen`  t " + _
" where not t.fecha_pago is null   and  t.fecha_vencimiento >= '" + strfechaMySQL(vfecha) + "'"



getCobrosUrbano = traerDatos2(vsql, "c", pathDBMySQLComuna)

If Err Then Exit Function
End Function

Function getCobrosRuralDetalle(vfecha As Date)
On Error Resume Next

Dim vsql As String
Dim vp2, vt2 As Double

vsql = " SELECT  count(1) As c  From   `recibo_resumen_rurales` t  where t.fecha_emision >= '" + strfechaMySQL(vfecha) + "' and  NOT `t`.`fecha_pago` IS NULL"
vp2 = traerDatos2(vsql, "c", pathDBMySQLComuna)


vsql = " SELECT  count(1) As c  From   `recibo_resumen_rurales` t  where t.fecha_emision >= '" + strfechaMySQL(vfecha) + "' and   `t`.`fecha_pago` IS NULL"
vt2 = traerDatos2(vsql, "c", pathDBMySQLComuna)

getCobrosRuralDetalle = Format(vt2 / vp2, "#0.00")

If Err Then Exit Function
End Function


Function getCobrosUrbanoDetalle(vfecha As Date)
On Error Resume Next

Dim vsql As String
Dim vp2, vt2 As Double

vsql = " SELECT  count(1) As c  From   `recibo_resumen` t  where t.fecha_emision >= '" + strfechaMySQL(vfecha) + "' and  NOT `t`.`fecha_pago` IS NULL"
vp2 = traerDatos2(vsql, "c", pathDBMySQLComuna)


vsql = " SELECT  count(1) As c  From   `recibo_resumen` t  where t.fecha_emision >= '" + strfechaMySQL(vfecha) + "' and   `t`.`fecha_pago` IS NULL"
vt2 = traerDatos2(vsql, "c", pathDBMySQLComuna)

getCobrosUrbanoDetalle = Format(vt2 / vp2, "#0.00")

If Err Then Exit Function
End Function




Function getCobrosRural(vfecha As Date) As Double
On Error Resume Next

Dim vsql As String


vsql = " SELECT  sum(`t`.`importe_total2`) AS c From   `recibo_resumen_rurales`  t " + _
" where not t.fecha_pago is null   and  t.fecha_vencimiento >= '" + strfechaMySQL(vfecha) + "'"



getCobrosRural = traerDatos2(vsql, "c", pathDBMySQLComuna)

If Err Then Exit Function
End Function

Function getCtasdetalle(vfecha As Date, vcta As String) As String
On Error Resume Next
Dim vsql As String

vsql = " SELECT  asientos.Fecha,  asientosdetalle.Debe,  asientosdetalle.Haber,  asientos.Leyenda  From  `asientos` " + _
 " INNER JOIN `asientosdetalle` ON (`asientos`.`Numero` = `asientosdetalle`.`Numero`) " + _
 " INNER JOIN `cuentas` ON (`asientosdetalle`.`CodigoCuenta` = `cuentas`.`CodigoCuenta`) " + _
 " Where (`cuentas`.`Cuenta` LIKE '" + Trim(vcta) + "' or  cuentas.CodigoCuenta  ='" + Trim(vcta) + "') and  fecha >= ' " + strfechaMySQL(vfecha - 90) + "'"
 
getCtasdetalle = vsql


If Err Then Exit Function
End Function

Function getCtas(vfecha As Date, vcta As String) As Double
On Error Resume Next
Dim vsql As String

vsql = " SELECT  avg(debe+ haber) AS `c` From  `asientos` " + _
 " INNER JOIN `asientosdetalle` ON (`asientos`.`Numero` = `asientosdetalle`.`Numero`) " + _
 " INNER JOIN `cuentas` ON (`asientosdetalle`.`CodigoCuenta` = `cuentas`.`CodigoCuenta`) " + _
 " Where (`cuentas`.`Cuenta` LIKE '" + Trim(vcta) + "' or  cuentas.CodigoCuenta  ='" + Trim(vcta) + "') and  fecha >= ' " + strfechaMySQL(vfecha - 90) + "'"
 
getCtas = traerDatos2(vsql, "c", pathDBMySQL)


If Err Then Exit Function
End Function
Function getEventuales(vfecha As Date) As Double
On Error Resume Next
Dim vsql As String

vsql = " SELECT  avg(`asientosdetalle`.`haber`) AS `saldo` From  `asientos` " + _
 " INNER JOIN `asientosdetalle` ON (`asientos`.`Numero` = `asientosdetalle`.`Numero`) " + _
 " INNER JOIN `cuentas` ON (`asientosdetalle`.`CodigoCuenta` = `cuentas`.`CodigoCuenta`) " + _
 " Where `cuentas`.`Cuenta` LIKE 'eventual' and  fecha >= ' " + strfechaMySQL(vfecha - 90) + "'"
 
getEventuales = traerDatos2(vsql, "c", pathDBMySQL)

If Err Then Exit Function
End Function

Function getCooparticipacion(vfecha As Date) As Double
On Error Resume Next
Dim vsql As String

vsql = " SELECT  avg(`asientosdetalle`.`haber`) AS `saldo` From  `asientos` " + _
 " INNER JOIN `asientosdetalle` ON (`asientos`.`Numero` = `asientosdetalle`.`Numero`) " + _
 " INNER JOIN `cuentas` ON (`asientosdetalle`.`CodigoCuenta` = `cuentas`.`CodigoCuenta`) " + _
 " Where `cuentas`.`Cuenta` LIKE '%coop%' and  fecha >= ' " + strfechaMySQL(vfecha - 90) + "'"
 
getCooparticipacion = traerDatos2(vsql, "c", pathDBMySQL)

If Err Then Exit Function
End Function


Function getfae(vfecha As Date) As Double
On Error Resume Next
Dim vsql, vsql2 As String

vsql = " select " + _
"   sum(t.Debe ) * 0.05 as c " + _
" from asientosdetalle t " + _
" inner join asientos a  on t.Numero = a.Numero   " + _
" where  " + _
" t.CodigoCuenta ='01.01.01.0001.0002'  " + _
" or " + _
" t.CodigoCuenta ='01.01.01.0001.0003' " + _
" or " + _
" t.CodigoCuenta ='01.01.01.0001.0007' " + _
" or " + _
" t.CodigoCuenta ='01.01.01.0001.0010' " + _
" or " + _
" t.CodigoCuenta ='01.01.01.0001.0028' " + _
" or " + _
" t.CodigoCuenta ='01.01.01.0001.0008' " + _
" or " + _
" t.CodigoCuenta ='01.01.02.0001.0000' " + _
" or " + _
" t.CodigoCuenta ='01.01.02.0000.0000' " + _
" or " + _
" t.CodigoCuenta ='01.03.01.0002.0000' " + _
" or "

vsql2 = " t.CodigoCuenta ='01.01.04.0001.0000' " + _
" or " + _
" t.CodigoCuenta ='01.01.04.0002.0000' " + _
" or " + _
" t.CodigoCuenta ='01.01.04.0003.0000' " + _
" or " + _
" t.CodigoCuenta ='01.01.04.0004.0000' " + _
" or " + _
" t.CodigoCuenta ='01.01.04.0005.0000' " + _
" and  " + _
" year(a.fecha) = " + Str(Year(a.fecha)) + _
" and  " + _
" month(a.fecha) = " + Str(Month(vfecha))

vsql = vsql + vsql2

getfae = traerDatos2(vsql, "c", pathDBMySQL)

If Err Then Exit Function
End Function





