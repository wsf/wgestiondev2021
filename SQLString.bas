Attribute VB_Name = "SQLString"
Public Function fSQLCheques_temp(vIdCheques As String) As String
    fSQLCheques_temp = " SELECT b.Descripcion as c FROM bancos b  INNER JOIN cheques_temp ct ON (b.idBancos=c" & _
    "t.idbancocaja)  Where ct.idCheques =" & vIdCheques
End Function


Public Function fSQLConciliaBancoCtas(vfd As Date, vfh As Date, vnot As String) As String
fSQLConciliaBancoCtas = "select " & _
 "  `bancosmovimientos`.`Fecha`,   `bancosmovimientos`.`Comentario`,   `bancos`.`D" & _
 "escripcion`,   `bancosmovimientos`.`NroCheque`,   `bancosmovimientos`.`Debito`, " & _
 "  `bancosmovimientos`.`Credito`,   `bancosmovimientos`.`NroInterno`,   `asientos" & _
 "detalle`.`Numero`,   `asientosdetalle`.`Haber`, `asientosdetalle`.`Debe`,`asientosdetalle`.`CodigoCuent" & _
 "a`,`asientos`.`Leyenda`" & _
 " from `asientos`    inner join `asientosdetalle` on      (`asientos`" & _
 ".`Numero` = `asientosdetalle`.`Numero`)    inner join `bancos` on      (`asiento" & _
 "sdetalle`.`CodigoCuenta` = `bancos`.`CuentaContableAsociada`)    inner join `ban" & _
 "cosmovimientos` on      (`bancos`.`idBancos` = `bancosmovimientos`.`idBancos`)  " & _
 "        where  " & vnot & " (bancosmovimientos.Credito + bancosmovimientos.Debito) = (asiento" & _
 "sdetalle.Debe + asientosdetalle.Haber)          and           asientos.Fecha >= '" & _
 strfechaMySQL(vfd) & "' and asientos.fecha <= '" & strfechaMySQL(vfh) & "' AND (`bancosmovimientos`.`NroInterno` = `asientos`.`NroInte" & _
 "rno`)"
End Function

Public Function FSQLSaldos(ByVal vsqlFecha As String, ByVal vtabla, Optional tipo As String, Optional vcondi, Optional vestadodocumento As String, Optional ByVal vcorto As Variant, Optional via As String) As String
Dim vcp As String
Dim vtipo, vtablaFactura As String

vtipo = ""


If vtabla = "cuentascorrientes" Then vcp = "clientes"
If vtabla = "pcuentascorrientes" Then
    vcp = "proveedores"
    vtablaFactura = "pfactura"
Else
    vtablaFactura = "factura"
End If




If Not tipo = "" And vcp = "clientes" Then vcondi = vcondi + " and tipocliente = '" + tipo + "'"


If Not tipo = "" And vcp = "proveedores" Then vcondi = vcondi + " and tipoproveedor = '" + tipo + "'"


If vtabla = "pcuentascorrientes" And tipo = "Eventual" Then vtipo = " and (tipoproveedor='Eventuales')"


If vestadodocumento = "Aceptados" Then
        FSQLSaldos = " select C.codigo,C.Localidad,C.Nombre," + _
        " (select max(fecha) as f from " + vtabla + " where debito > 0 and codigo=c.codigo) FechaD, " + _
        " (select max(fecha) as f from " + vtabla + "  where credito > 0 and codigo=c.codigo) FechaC, " + _
        "C.Telefono,sum(CCC.Debito) - sum(CCC.Credito) as SSaldo, " + _
        "(sum(ccc.debito) - sum(ccc.credito) - (select debito as du from " + vtabla + "  " + _
        " where debito > 0  and codigo = ccc.codigo  order by fecha desc limit 1)) as sr, " + _
        " (select  debito as du from " + vtabla + "  where debito > 0 and codigo = ccc.codigo order by  fecha desc limit 1) as uf " + _
        " from " + vtabla + " CCC  inner join " + vcp + " C on c.Codigo = CCC.Codigo " + _
        "where 1=1 " + vcondi + " and (not CCC.estadoadmicion = 1 or CCC.estadoadmicion is null)  " + _
        "group by C.codigo "
Else
        
        If UCase(LeerXml("Puesto")) = "PONS" Then
       
                    FSQLSaldos = " select C.codigo,C.Localidad,C.Nombre," + _
                    "C.Telefono,sum(CCC.Debito) - sum(CCC.Credito) as SSaldo, " + _
                    "(sum(ccc.debito) - sum(ccc.credito) - (select debito as du from " + vtabla + "  " + _
                    " where debito > 0  and codigo = ccc.codigo  order by fecha desc limit 1)) as sr, " + _
                    " (select  debito as du from " + vtabla + "  where debito > 0 and codigo = ccc.codigo order by  fecha desc limit 1) as uf " + _
                    " from " + vtabla + " CCC  inner join " + vcp + " C on c.Codigo = CCC.Codigo where 1=1 " + vcondi + "group by C.codigo "
                    
             
        Else
        
                Dim vcome1 As String
                vcome1 = " and (comentario like '%" + via + "%' or comentario is null)"
                
                
                
                 FSQLSaldos = " select C.codigo,C.Localidad,C.Nombre," + _
                    " (select max(fecha) as f from " + vtabla + " where debito > 0 and codigo=c.codigo " + vcome1 + ") FechaD, " + _
                    " (select max(fecha) as f from " + vtabla + "  where credito > 0 and codigo=c.codigo " + vcome1 + " ) FechaC, " + _
                    "C.Telefono,sum(CCC.Debito) - sum(CCC.Credito) as SSaldo, " + _
                    "(sum(ccc.debito) - sum(ccc.credito) - (select debito as du from " + vtabla + "  " + _
                    " where debito > 0  and codigo = ccc.codigo " + vcome1 + " order by fecha desc limit 1)) as sr, " + _
                    " (select  debito as du from " + vtabla + "  where debito > 0 and codigo = ccc.codigo " + vcome1 + " order by  fecha desc limit 1) as uf " + _
                    " from " + vtabla + " CCC  inner join " + vcp + " C on c.Codigo = CCC.Codigo where 1=1 " + vcondi + vcome1 + " group by C.codigo "
            

        End If
        

End If


If vcorto = 1 And Not UCase(LeerXml("Puesto")) = "PONS" Then

                'Dim vcome1 As String
                vcome1 = "  and  (comentario like '%" + via + "%' or comentario is null)"
                
                
                

                    FSQLSaldos = " select C.codigo,C.Localidad as Localidad,C.Nombre," + _
                    "C.Telefono,sum(CCC.Debito) - sum(CCC.Credito) as SSaldo, " + _
                    " '' as FechaD, '' as FechaC, '' as uf " + _
                    " from " + vtabla + " CCC  inner join " + vcp + " C on c.Codigo = CCC.Codigo where 1=1 " + vcondi + vcome1 + " group by C.codigo "
                    
End If



Debug.Print ("--------> Consulta de saldos : " + FSQLSaldos)


'FSQLSaldos = " select cli.codigo as Codigo, cli.nombre as Nombre , cf.uc as FechaC, c.ud as FechaD,  cli.Telefono as Telefono, cc.saldo as SSaldo   from  " + vcp + "  cli " + _
'" Inner Join " + _
'" (select  max(fecha) as ud, codigo   from " + vtabla + " where debito > 0 group by codigo) as c " + _
'" on cli.Codigo = c.codigo " + _
'" Inner Join " + _
'" (select  max(fecha) as uc, codigo   from " + vtabla + "  where credito > 0 group by codigo) as cf " + _
'" on cli.Codigo = cf.codigo " + _
'" Inner Join " + _
'" (select a.Codigo, (a.debito) -sum(a.credito) as saldo from " + vtabla + "  a group by a.codigo  ) as cc " + _
'" on cli.Codigo = cc.codigo "


'FSQLSaldos = "SELECT   C.Codigo,   C.Nombre,  (SELECT max(fecha) AS m  FROM " & vtabla & _
' " WHERE debito > 0 and codigo = ccc.codigo) FechaD,  (SELECT max(fecha) AS m  FROM " & vtabla & " WHERE" & _
' " credito > 0 and codigo = ccc.codigo) FechaC,   Telefono,    Sum(CCC.Debito) AS TotalDebito,   Sum(CCC.C" & _
' "redito) AS TotalCredito,   Sum(CCC.Debito) -Sum(CCC.Credito) AS SSaldo FROM  " & vcp & _
' " C   LEFT JOIN " & vtabla & " CCC     ON   C.Codigo = CCC.Codigo " & vsqlFecha & vTipo & _
' " GROUP BY   C.Codigo"
End Function




Public Function fSQLConciliaBancoCtas2(vfd As Date, vfh As Date, vnot As String) As String

fSQLConciliaBancoCtas2 = "SELECT  * from conciliacion2  where   not saldo = 0 and " & _
 "  Fecha >= '" & _
 strfechaMySQL(vfd) & "' and fecha <= '" & strfechaMySQL(vfh) & "'"
End Function


