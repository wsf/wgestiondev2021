Attribute VB_Name = "Control_Documentos"
Option Explicit
Public vf As vfactura


Public Sub setFacturaDatosRemito(vremito As Long)

Dim vsql, vsql2, vcodigo  As String
Dim vnrointerno As Long

vsql = "select * from factura  where   factura.remito = " + Str(vremito) + " order by factura.NroInterno desc limit 1"
vcodigo = traerDatos2(vsql, "codigo", pathDBMySQL)
vsql2 = "select * from clientes  where codigo = " + Str(vcodigo)

vnrointerno = Val(traerDatos2(vsql, "nrointerno", pathDBMySQL))
vf.vcodigo = vcodigo
vf.vfecha = CDate(traerDatos2(vsql, "fecha", pathDBMySQL))
vf.vnombre = traerDatos2(vsql, "nombre", pathDBMySQL)
vf.vdireccion = traerDatos2(vsql2, "direccion", pathDBMySQL)
vf.vlocalidad = traerDatos2(vsql2, "localidad", pathDBMySQL)
vf.vcuit = traerDatos2(vsql2, "cuit", pathDBMySQL)
vf.vtotal = Val(traerDatos2(vsql, "total", pathDBMySQL))
vf.vSubTotal = Val(traerDatos2(vsql, "subtotal", pathDBMySQL))
vf.viva = traerDatos2(vsql, "Iva", pathDBMySQL)

vf.vsaldo = getSaldoCliente2(vcodigo)

vsql = "select * from ivafacturaventa where ivafacturaventa.nrointerno =  " + Str(vnrointerno)
vf.viva150 = Val(traerDatos2(vsql, "iva105", pathDBMySQL))
vf.vIva210 = Val(traerDatos2(vsql, "iva210", pathDBMySQL))
vf.viva270 = Val(traerDatos2(vsql, "iva270", pathDBMySQL))


End Sub


Public Sub llenarDetallesRemito(vremito As Long, bdetalle)

bdetalle.ConnectionString = pathDBMySQL
bdetalle.RecordSource = "select * from fdetalle where remito = " + Str(vremito)
bdetalle.Refresh

Dim i As Integer

i = 0
With bdetalle.Recordset

    .MoveFirst
    
    Do Until .EOF
        
        i = i + 1
        
        Call llenarlinea2Remito(i, .Fields("cantidad"), .Fields("detalle"), .Fields("precio"), Val(EsNulo(.Fields("descuento"))), .Fields("total"))
        
        .MoveNext
    Loop
    

End With

End Sub


Public Sub setMarcarImpresaRemito(vremito As Long)
    Dim vsql As String
    
    vsql = "update factura set estado2= 'Impreso' where remito =" + Str(vremito)
    Call EjecutarScript(vsql, pathDBMySQL)
End Sub


Public Sub mostrar_Doc2Remito(ByVal vtipo As String, ByVal vncomprobante As Integer)

Call llenarDocumentos2Remito(vtipo, vncomprobante)


With drDoc2
    If vf.viva150 = 0 Then
.Sections("titulos").Controls("eiva105").Visible = False
.Sections("titulos").Controls("etiva10").Visible = False

.Sections("titulos").Controls("eeiva105").Visible = False
.Sections("titulos").Controls("eetiva10").Visible = False
    End If


If vf.vIva210 = 0 Then
.Sections("titulos").Controls("eiva21").Visible = False
.Sections("titulos").Controls("etiva21").Visible = False

.Sections("titulos").Controls("eeiva21").Visible = False
.Sections("titulos").Controls("eetiva21").Visible = False
End If


If vf.viva270 = 0 Then
.Sections("titulos").Controls("eiva27").Visible = False
.Sections("titulos").Controls("etiva27").Visible = False

.Sections("titulos").Controls("eetiva27").Visible = False
.Sections("titulos").Controls("eeiva27").Visible = False


End If


.Sections("titulos").Controls("esaldo").Caption = Format(vf.vsaldo, "###,###,##0.00")
.Sections("titulos").Controls("eesaldo").Caption = Format(vf.vsaldo, "###,###,##0.00")

'Set .DataSource = Nothing
'.DataMember = ""
.Show
'Call .PrintReport(False, rptRangeAllPages)
'Unload .object
End With



End Sub



Public Sub llenarlinea2Remito(vi As Integer, vCantidad As Double, vDetalle As String, vPrecio As Double, vdesc As Double, vtotal As Double)
Dim ve, vd, vp, vdes, vt As String

With drDoc2.Sections("titulos")

ve = "e" + Trim(Str(vi))
vd = "d" + Trim(Str(vi))
vp = "p" + Trim(Str(vi))
vdes = "des" + Trim(Str(vi))
vt = "t" + Trim(Str(vi))


.Controls(ve).Caption = Str(vCantidad)
.Controls(vd).Caption = vDetalle
.Controls(vp).Caption = Format(vPrecio, "###,###,##0.00")
.Controls(vdes).Caption = Format(vdesc, "###,###,##0.00")
.Controls(vt).Caption = Format(vtotal, "###,###,##0.00")

ve = "ee" + Trim(Str(vi))
vd = "dd" + Trim(Str(vi))
vp = "pp" + Trim(Str(vi))
vdes = "ddes" + Trim(Str(vi))
vt = "tt" + Trim(Str(vi))


.Controls(ve).Caption = Str(vCantidad)
.Controls(vd).Caption = vDetalle
.Controls(vp).Caption = Format(vPrecio, "###,###,##0.00")
.Controls(vdes).Caption = Format(vdesc, "###,###,##0.00")
.Controls(vt).Caption = Format(vtotal, "###,###,##0.00")



End With

End Sub


Private Sub llenarDocumentos2Remito(vtipo As String, vncomprobante As Integer)

Dim Form As DataReport


With drDoc2

                '----------- titulos -------
                .Sections("titulos").Controls("enroremito").Caption = ""
                .Sections("titulos").Controls("ecventa").Caption = "Cuentas Corrientes"
                
                .Sections("titulos").Controls("enombre").Caption = vf.vnombre
                .Sections("titulos").Controls("edomicilio").Caption = vf.vdireccion
                .Sections("titulos").Controls("elocalidad").Caption = vf.vlocalidad
                .Sections("titulos").Controls("ecuit").Caption = vf.vcuit
                .Sections("titulos").Controls("efecha").Caption = vf.vfecha
                '---------------------------
                
                .Sections("titulos").Controls("etotal").Caption = Format(vf.vtotal, "#,###,##0.00")
                .Sections("titulos").Controls("esubtotal").Caption = Format(vf.vSubTotal, "#,###,##0.00")
                
                .Sections("titulos").Controls("eiva105").Caption = Format(vf.viva150, "#,###,##0.00")
                .Sections("titulos").Controls("eiva21").Caption = Format(vf.vIva210, "#,###,##0.00")
                .Sections("titulos").Controls("eiva27").Caption = Format(vf.viva270, "#,###,##0.00")
                
                .Sections("titulos").Controls("edescuento").Caption = ""
                '.Sections("Totales").Controls("eimpuesto").Caption = Format(vgTimpuesto, "#,###,##0.00")
                
                '.Sections("Totales").Controls("etotal").Caption = Format(vgTtotal, "#,###,##0.00")
                
                .Sections("titulos").Controls("etitulo").Caption = ""
                .Sections("titulos").Controls("eetitulo").Caption = ""
                
                
                
                If vtipo = "Documento" Then
                    .Sections("titulos").Controls("etitulo").Caption = "Docuemento no válido como factura"
                    .Sections("titulos").Controls("etitulo").Caption = "Docuemento no válido como factura"
                    .Sections("titulos").Controls("encomprobante").Caption = "Nro. Comprobante : " + Str(vncomprobante)
                    .Sections("titulos").Controls("eetitulo").Caption = "Docuemento no válido como factura"
                    .Sections("titulos").Controls("eencomprobante").Caption = "Nro. Comprobante : " + Str(vncomprobante)
                    
                    
                    .Sections("titulos").Controls("ttiva").Caption = ""
                    .Sections("titulos").Controls("tttiva").Caption = ""
                End If
                
                
                If vtipo = "Exento" Then
                    .Sections("titulos").Controls("etitulo").Caption = ""
                    .Sections("titulos").Controls("ncomprobante").Caption = ""
                    .Sections("titulos").Controls("eetitulo").Caption = ""
                    .Sections("titulos").Controls("encomprobante").Caption = ""
                    
                    
                    .Sections("titulos").Controls("ttiva").Caption = ""
                    .Sections("titulos").Controls("tttiva").Caption = ""
                End If
                
                
                
                If vtipo = "Fact A" Then
                    .Sections("titulos").Controls("etitulo").Caption = ""
                    .Sections("titulos").Controls("encomprobante").Caption = ""
                    .Sections("titulos").Controls("eetitulo").Caption = ""
                    .Sections("titulos").Controls("encomprobante").Caption = " "
                 
                End If





'----------- titulos -------
.Sections("titulos").Controls("eenroremito").Caption = ""
.Sections("titulos").Controls("eecventa").Caption = "Cuentas Corrientes"

.Sections("titulos").Controls("eenombre").Caption = vf.vnombre
.Sections("titulos").Controls("eedomicilio").Caption = vf.vdireccion
.Sections("titulos").Controls("eelocalidad").Caption = vf.vlocalidad
.Sections("titulos").Controls("eecuit").Caption = vf.vcuit
.Sections("titulos").Controls("eefecha").Caption = vf.vfecha
'---------------------------

.Sections("titulos").Controls("eetotal").Caption = Format(vf.vtotal, "#,###,##0.00")
.Sections("titulos").Controls("eesubtotal").Caption = Format(vf.vSubTotal, "#,###,##0.00")

.Sections("titulos").Controls("eeiva105").Caption = Format(vf.viva150, "#,###,##0.00")
.Sections("titulos").Controls("eeiva21").Caption = Format(vf.vIva210, "#,###,##0.00")
.Sections("titulos").Controls("eeiva27").Caption = Format(vf.viva270, "#,###,##0.00")

.Sections("titulos").Controls("eedescuento").Caption = ""
'.Sections("Totales").Controls("eeimpuesto").Caption = Format(vgTimpuesto, "#,###,##0.00")

'.Sections("Totales").Controls("eetotal").Caption = Format(vgTtotal, "#,###,##0.00")

If vtipo = "Documento" Then
    .Sections("titulos").Controls("eetitulo").Caption = "Docuemento no válido como factura"
    .Sections("titulos").Controls("encomprobante").Caption = "Nro. Comprobante : "
Else
    .Sections("titulos").Controls("eetitulo").Caption = ""
    .Sections("titulos").Controls("eencomprobante").Caption = ""
End If





End With

End Sub
