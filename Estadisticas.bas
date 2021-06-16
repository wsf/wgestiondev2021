Attribute VB_Name = "Estadisticas"
Function sqlArtFactVenta(ByVal vwhere As String, ByVal vgroup As String, ByVal vorder As String) As String
sqlArtFactVenta = "select   fdetalle.codigo as codigo,   " & _
 " ucase(fdetalle.detalle) as Detalle,    sum(fdetalle.cantidad) cantidad,    sum(fdetalle.total) as " & _
 "Total, (sum(fdetalle.`Pcosto` * fdetalle.cantidad)) as TPcosto,   (sum(fdetalle.total) -  sum(fdetalle" & _
 ".Pcosto * fdetalle.cantidad)) as Ganancia from fdetalle    left outer join articulos on      " & _
 "(fdetalle.codigo = articulos.codigo)    inner join factura t on      (factura.remi" & _
 "to = fdetalle.remito)    where   1 = 1 " + vwhere + " " + vgroup + " order by " + "Detalle"
End Function


Function sqlArtFactVenta2(ByVal vwhere As String, ByVal vgroup As String, ByVal vorder As String, Optional ByVal vcv As String) As String
' vcv es para saber si viene de la consulta de ventas o de compras


If vgroup = "group by factura.codigo" Then
             
             sqlArtFactVenta2 = "select   t.codigo as codigo,   " & _
             "ucase(t.nombre) as Detalle,    sum(fdetalle.cantidad) cantidad,    sum(fdetalle.total) as " & _
             "Total,    sum(fdetalle.`Pcosto` * fdetalle.cantidad) as TPcosto,   sum(fdetalle.total) -  sum(fdetalle." & _
             "Pcosto * fdetalle.cantidad) as Ganancia from fdetalle    left outer join articulos on      " & _
             "(fdetalle.codigo = articulos.codigo)    inner join factura t on      (t.remi" & _
             "to = fdetalle.remito)    where   1 = 1 " + vwhere + " " + vgroup + " order by " + vorder
             
             
             If UCase(vcv) = UCase("pfactura") Then  ' si es una compra
                sqlArtFactVenta2 = Replace(sqlArtFactVenta2, "from fdetalle    left", "from pfdetalle    left", , , vbTextCompare)
                sqlArtFactVenta2 = Replace(sqlArtFactVenta2, "inner join factura", "inner join pfactura", , , vbTextCompare)
                sqlArtFactVenta2 = Replace(sqlArtFactVenta2, "factura.", "pfactura.", , , vbTextCompare)
                sqlArtFactVenta2 = Replace(sqlArtFactVenta2, "fdetalle.", "pfdetalle.", , , vbTextCompare)
             End If


Else
            
            
            sqlArtFactVenta2 = "select   fdetalle.codigo as codigo,   " & _
             "ucase(fdetalle.detalle) as Detalle,    sum(fdetalle.cantidad) cantidad,    format(sum(fdetalle.total),2) as " & _
             "Total,    format(sum(fdetalle.`Pcosto` * fdetalle.cantidad),2) as TPcosto,   format((sum(fdetalle.total) -  sum(fdetalle" & _
             ".Pcosto * fdetalle.cantidad)),2) as Ganancia from fdetalle    left outer join articulos on      " & _
             "(fdetalle.codigo = articulos.codigo)    inner join factura t on      (t.remi" & _
             "to = fdetalle.remito)    where   1 = 1 " + vwhere + " " + vgroup + " order by " + "Detalle"
             
             
             
             
End If
End Function




Public Sub valoresGraficas(rs As Adodc, ByRef Chart As MSChart, vtitulo As String)
 ReDim Values(1 To rs.Recordset.RecordCount, 1 To 4)
  
         'Cargar los datos.
         rs.Recordset.MoveFirst
         For i = 1 To rs.Recordset.RecordCount
              Values(i, 1) = Str(rs.Recordset("ano")) + "-" + Str(rs.Recordset("mes"))
              Values(i, 2) = rs.Recordset("deuda")
              Values(i, 3) = rs.Recordset("pago")
              Values(i, 4) = Val(EsNulo(rs.Recordset("por")))
              rs.Recordset.MoveNext
         Next i
        
         'Dibujar gráfica
         Chart.ChartType = VtChChartType2dXY
         Chart.Plot.Axis(VtChAxisIdX).AxisTitle.Text = "Muestras [n]"
         Chart.Plot.Axis(VtChAxisIdY).AxisTitle.Text = "Cantidad"
         Chart.TitleText = vtitulo
        ' Chart.RowCount = rs.Recordset.RecordCount
        ' Chart.ColumnCount = 2
         Chart.ChartData = Values

End Sub
