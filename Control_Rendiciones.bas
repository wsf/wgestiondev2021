Attribute VB_Name = "Control_Rendiciones"
Public Sub faedetalle()
Dim vul, sqlSaldoCaja As String
On Error Resume Next
    
    
    Unload Mantenimiento
    Load Mantenimiento
    
    With Mantenimiento.rsFAEDetalle
        If .State = 1 Then .Close
        
        sqlSaldoCaja = ""
        
        frmFiltro.vcomando = "fae"
        frmFiltro.Show


        
        If Not .State = 1 Then .Open
        .Close
        .Open
        
    End With
    
    With drFaeDetalle
    .Show
    End With

If Err Then Exit Sub
End Sub



 
Public Function fsqlFaeDetalle2(vwhere As String) As String
Dim vsq As String
vsql = "select concat(year(fecha),' - ', month(fecha)) as Periodo, format((sum(t.Debe ) * 0.05), '###########') as Importe1, format((sum(t.Debe) * 0.025), '###########') as Importe2, " + _
" format(sum(t.Debe), '###########') as Total,  c.cuenta, c.codigocuenta from asientos a inner join asientosdetalle t on a.numero = t.numero  Inner Join Cuentas c " + _
" on c.codigocuenta = t.codigocuenta " + _
" Where (t.CodigoCuenta  ='01.01.01.0001.0002' or t.CodigoCuenta ='01.01.01.0001.0003' or t.CodigoCuenta ='01.01.01.0001.0007' or t.CodigoCuenta ='01.01.01.0001.0010' " + _
" or t.CodigoCuenta ='01.01.01.0001.0028' or  t.CodigoCuenta ='01.01.01.0001.0008' or t.CodigoCuenta ='01.01.02.0001.0000' or t.CodigoCuenta ='01.01.02.0000.0000' " + _
" or t.CodigoCuenta ='01.03.01.0002.0000' or t.CodigoCuenta ='01.01.04.0001.0000' or  t.CodigoCuenta ='01.01.04.0002.0000' or t.CodigoCuenta ='01.01.04.0003.0000' " + _
" or t.CodigoCuenta ='01.01.04.0004.0000' or t.CodigoCuenta ='01.01.04.0005.0000') and " + _
vwhere + _
" group by periodo, cuenta  order by fecha desc "
fsqlFaeDetalle2 = vsql
End Function


 
Public Function fsqlFaeDetalle(vwhere As String) As String
 
fsqlFaeDetalle = " SHAPE {select substring(CAST(t.fecha2  AS CHAR),1,6) as Periodo, (sum(t.Importe) * 0.05) as Total, (sum(t.Importe) * 0.025) as Total2, " & _
" sum(t.Importe) as sub from movimientoscaja t " & _
" where ( t.CodCuenta ='01.01.01.0001.0002'  or t.CodCuenta ='01.01.01.0001.0003' or t.CodCuenta ='01.01.01.0001.0007' or t.CodCuenta ='01.01.01.0001.0010' " & _
" or t.CodCuenta ='01.01.01.0001.0028' or t.CodCuenta ='01.01.01.0001.0008' or t.CodCuenta ='01.01.02.0001.0000' or t.CodCuenta ='01.01.02.0000.0000' " & _
" or t.CodCuenta ='01.03.01.0002.0000' or t.CodCuenta ='01.01.04.0001.0000' or t.CodCuenta ='01.01.04.0002.0000' " & _
" or t.CodCuenta ='01.01.04.0003.0000' or t.CodCuenta ='01.01.04.0004.0000' or t.CodCuenta ='01.01.04.0005.0000' )" & _
vwhere & _
" group by Periodo order by fecha2 desc }  AS FaeDetalle APPEND ({select " & _
"  substring(CAST(t.fecha2  AS CHAR),1,6) as Periodo,   (sum(t.Importe) * 0.05) as Total, (sum(t.Importe) * 0.025) as Total2 , sum(t.Importe) as sub, tt.CodCuenta, tt.Descripcion " & _
" from movimientoscaja t  inner join cuentacontable tt  on t.CodCuenta = tt.CodCuenta " & _
" where ( t.CodCuenta ='01.01.01.0001.0002'  or t.CodCuenta ='01.01.01.0001.0003' or t.CodCuenta ='01.01.01.0001.0007' or t.CodCuenta ='01.01.01.0001.0010' " & _
" or t.CodCuenta ='01.01.01.0001.0028' or t.CodCuenta ='01.01.01.0001.0008' or t.CodCuenta ='01.01.02.0001.0000' or t.CodCuenta ='01.01.02.0000.0000' " & _
" or t.CodCuenta ='01.03.01.0002.0000' or t.CodCuenta ='01.01.04.0001.0000' or t.CodCuenta ='01.01.04.0002.0000' " & _
" or t.CodCuenta ='01.01.04.0003.0000' or t.CodCuenta ='01.01.04.0004.0000' or t.CodCuenta ='01.01.04.0005.0000' )" & _
vwhere & _
" group by Periodo, tt.CodCuenta order by fecha2 desc }  AS Detalle RELATE 'Periodo' TO 'Periodo') AS Detalle "

End Function

