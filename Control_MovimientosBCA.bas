Attribute VB_Name = "Control_MovimientosBCA"
Option Explicit

Dim vtD, vtC, vtGralD, vtGralC As Double


Private Sub denunciarNumeroNoEmitido(vnro As Long)
'-------------- poner ------------------
Dim vsql, vcampos, vValor As String

vcampos = "c1,c2,c3,c4"
vValor = "'*****','[" + Str(vnro) + "]', 'Nro de comprobante no emitido'," + "'" + Str(vnro) + "'"
vsql = "insert into movibca (" + vcampos + ") values (" + vValor + ")"

Call EjecutarScript(vsql, PathDBListados)

'------------------------------------
End Sub



Public Sub Reporte_MoviCBA(vfd As Date, vfh As Date, Optional ByVal vidd As Long, Optional ByVal vidH As Long, Optional vsolofaltantes As Boolean)
On Error Resume Next
Dim vsaldo As Double
Dim vcampos, vvalores, vsql, vvcp, vsqlFecha, vcoment2, vpersona, vsql1 As String
Dim vNroComprobanteAnterior As Long


Dim rs As New ADODB.Recordset
Dim vnrocomprobante, vnrointeno  As Long

vsql = "delete from movibca"
Call EjecutarScript(vsql, PathDBListados)  ' vacio la tabla
 
vcampos = "c1,c2,c3,c4,i1,i2"

If vidH > 0 Then
    vsqlFecha = " idbancosmovimientos > " + Str(vidd) + " and idbancosmovimientos <= " + Str(vidH)
Else
    vsqlFecha = " fecha >= '" + strfechaMySQL(vfd) + "' and fecha <= '" + strfechaMySQL(vfh) + "'"
End If

vsql = " SELECT * FROM bancosmovimientos t " + _
" inner join bancos b " + _
" on t.idBancos = b.idBancos where " + vsqlFecha

 

With rs
        
        .CursorLocation = adUseClient
        
        Call .Open(vsql, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .RecordCount > 0 Then
            MsgBox "No hay datos para mostar", vbInformation
            Exit Sub
        End If
        
        .MoveFirst

        vnrocomprobante = .Fields("nrocomprobante")
       
        
       ' vsql1 = "select * proveedo where codigo = '" + .Fields("codpersona") + "'"
      
       ' Call EjecutarScript(vsql1, PathDBListados)


         vtC = 0
         vtD = 0
        
        vnrointeno = .Fields("nrointerno")
        vcoment2 = .Fields("comentario2")
        
        frmBancoCajaDetalle.barra.Max = .RecordCount
        frmBancoCajaDetalle.barra.Value = 0
        
        
            '----------------------------------------
                  vpersona = ""
                  vsql1 = "select * from  proveedores where codigo = '" + .Fields("ClienteProveedor") + "'"
                  vpersona = traerDatos2(vsql1, "Nombre", pathDBMySQL)
            '---------------------------------------
   
Dim i As Integer
   
i = 0

initbarra (.RecordCount)
        
        
        
        
        
        Do Until .EOF = True
                     i = i + 1
                     actualizarBarra (i)
                     
                     
                   ' ------------------------------------------------------------------------
                   
                    If Not vnrocomprobante = .Fields("nrocomprobante") Then

                        vnrocomprobante = .Fields("nrocomprobante")
                                  
                                  'vpersona = ""
                                  vsql1 = "select * from  proveedores where codigo = '" + .Fields("ClienteProveedor") + "'"
                                  'vpersona = traerDatos2(vsql1, "Nombre", pathDBMySQL)
                            '---------------------------------------
                            
                            'vcoment2 = " - Persona: [" + vpersona + "]  -- " + .Fields("comentario2")
                            
                            
                           ' vnrointeno = .Fields("nrointerno")
                                             
                                             
                        If Not vsolofaltantes Then
                            Call GrabarAsiento(vnrointeno, vtD, vtC, vcoment2)
                        End If
                        
                            vpersona = traerDatos2(vsql1, "Nombre", pathDBMySQL)

                             vcoment2 = " - Persona: [" + vpersona + "]  -- " + .Fields("comentario2")
                           
                   
                            vnrointeno = .Fields("nrointerno")
                   End If
                   '-----------------------------------------------------------------------------
                     
                                          
                                          
                   vvalores = vvalores + "'" + Str(.Fields("fecha")) + "' ,"
                   vvalores = vvalores + "'" + (.Fields("Descripcion")) + " - [" + (Str(.Fields("nrocomprobante"))) + "] " + " [Tipo: " + (.Fields("TipoMovimiento")) + "]',"
                   vvalores = vvalores + "' " + (EsNulo(.Fields("Comentario"))) + "',"
                   vvalores = vvalores + "'Persona: [" + vpersona + "] -- " + ((EsNulo(.Fields("Comentario2")))) + "',"
                   vvalores = vvalores + Str(.Fields("Debito")) + ","
                   vvalores = vvalores + Str(.Fields("Credito"))
                   
                   vtD = vtD + .Fields("Debito")
                   vtC = vtC + .Fields("Credito")
                   
                   vtGralD = vtGralD + .Fields("Debito")
                   vtGralC = vtGralC + .Fields("Credito")
                   
                   
                
                   vsql = "insert into movibca (" + vcampos + ") values (" + vvalores + ")"
                   
                   Debug.Print vsql
                
                   If Not vsolofaltantes Then
                        Call EjecutarScript(vsql, PathDBListados)
                   End If
                
                   vvalores = ""
                                              
                    frmBancoCajaDetalle.barra.Value = frmBancoCajaDetalle.barra.Value + 1
                                              
                   
                   
                    
                   
                    
                   .MoveNext
                                   
        Loop
           
        .MovePrevious
        
        If Not vsolofaltantes Then
                Call GrabarAsiento(.Fields("nrointerno"), vtD, vtC, vcoment2)
        End If

    End With


   Call listar_nrocomprabantes_noimputados(vidd, vidH)


    Unload Mantenimiento
    Load Mantenimiento
    
        
    drBCMTotal.Sections("TituloEmpresa").Controls("lblFechaDesde").Caption = frmBancoCajaDetalle.dtpFecha(0).Value
    drBCMTotal.Sections("TituloEmpresa").Controls("lblFechaHasta").Caption = frmBancoCajaDetalle.dtpFecha(1).Value
        
        
    If Not vsolofaltantes Then
        drBCMTotal.Sections("PieInforme").Controls("totaldebe").Caption = Format(vtGralD, "###,###,##0.00")
        drBCMTotal.Sections("PieInforme").Controls("totalhaber").Caption = Format(vtGralC, "###,###,##0.00")
    End If
    
    drBCMTotal.Show


If Err Then Exit Sub

End Sub

Private Sub initbarra(i As Long)
    frmBancoCajaDetalle.barra.Value = 0
    frmBancoCajaDetalle.barra.Max = i + 1
End Sub

Private Sub actualizarBarra(i As Long)
    frmBancoCajaDetalle.barra.Value = i
End Sub

Private Sub GrabarAsiento(ByVal vnrointerno As Long, ByVal vtD As Double, ByVal vtC As Double, ByVal vcoment2 As String)
On Error Resume Next

Dim vsaldo As Double
Dim vcampos, vvalores, vsql As String
Dim rs As New ADODB.Recordset
Dim vnrocomprobante As Long

vcampos = "c1,c2,c3,i1,i2"

  vsql = "insert into movibca (c1) values ('_________')"
  Call EjecutarScript(vsql, PathDBListados)
  

vsql = " select * from asientos  a " + _
" inner join asientosdetalle ad " + _
" on a.Numero = ad.Numero  " + _
" inner join cuentas  " + _
" on cuentas.CodigoCuenta = ad.CodigoCuenta " + _
" where a.nrointerno = " + Str(vnrointerno)

With rs
        
        .CursorLocation = adUseClient
        
        Call .Open(vsql, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If .RecordCount > 0 Then
           
                        Do Until .EOF = True
                                        
                                       vvalores = vvalores + "'         > ',"
                                       vvalores = vvalores + "'',"
                                       vvalores = vvalores + "'" + .Fields("CodigoCuenta") + " - " + (.Fields("Cuenta")) + "',"
                                       vvalores = vvalores + Str(.Fields("Debe")) + ","
                                       vvalores = vvalores + Str(.Fields("Haber"))
                                    
                                       vsql = "insert into movibca (" + vcampos + ") values (" + vvalores + ")"
                                       Debug.Print vsql
                                    
                                       Call EjecutarScript(vsql, PathDBListados)
                                    
                                       vvalores = ""
                                                                      
                                       .MoveNext
                                       
                        Loop
                        
                                    Debug.Print (" ************************************** ")
    
        End If
 
    End With

  vcampos = "c1,c2,c3,i1,i2"



  vsql = "insert into movibca (c1,c2,i1,i2) values ('Totales: ','" + vcoment2 + "'," + Str(vtD) + "," + Str(vtC) + ")"
  Call EjecutarScript(vsql, PathDBListados)


  vsql = "insert into movibca (c2) values ('__________________________________________________________________________________________________')"
  Call EjecutarScript(vsql, PathDBListados)
  
If Error Then Exit Sub
End Sub



Public Sub ControlTransaccion(vnrointerno As Long, Optional vtipo As String)

On Error Resume Next

Dim vsql As String
Dim valor1, valor2 As Double

vsql = " select " + _
" sum(debito) - sum(credito) as c " + _
" from  " + _
" bancosmovimientos b " + _
" where b.nrointerno = " + Str(vnrointerno)

valor1 = Val(traerDatos2(vsql, "c", pathDBMySQL))

vsql = " select " + _
" sum(ad.debe) - sum(ad.haber) as c " + _
" from asientos a " + _
" inner join asientosdetalle ad " + _
" on a.numero = ad.numero  " + _
" where a.NroInterno = " + Str(vnrointerno)


valor2 = Val(traerDatos2(vsql, "c", pathDBMySQL))


If Not valor2 = valor1 And vtipo = "TR" Then
    MsgBox "Cuidado. Este movimiento no se grabó adecuadamente. Revise esta operación", vbCritical
    Call GrabarLog(" Diferencia entre Banco y Cta nro. interno: " + Str(vnrointerno), "Error", "CajaMoviIE")
End If


If Err Then Exit Sub

End Sub



Public Sub ControlAsientoBCMovimiento(vnrointerno As Long, Optional vtipo As String)

On Error Resume Next

Dim vsql As String
Dim valor1, valor2 As Double
Dim vnumero As Long

vsql = " select " + _
" sum(debito) - sum(credito) as c " + _
" from  " + _
" bancosmovimientos b " + _
" where b.nrointerno = " + Str(vnrointerno)

valor1 = Val(traerDatos2(vsql, "c", pathDBMySQL))

vsql = " select " + _
" sum(ad.debe) - sum(ad.haber) as c " + _
" from asientos a " + _
" inner join asientosdetalle ad " + _
" on a.numero = ad.numero  " + _
" where a.NroInterno = " + Str(vnrointerno)


valor2 = Val(traerDatos2(vsql, "c", pathDBMySQL))


If (Abs(valor2 - valor1) > 0.1) And Not (vtipo = "TR") And Not (vtipo = "VL") And Not (vtipo = "AD") And Not (vtipo = "AC") Then
    
    If Not Str(valor2 - valor1) = "0" Then
    
                MsgBox "Cuidado. Este movimiento no se grabó adecuadamente." _
                + Chr(13) + "El sistema se cerrará. Vuelva a ingresar el movimiento." + Chr(13) + Str(valor2) + " - " + Str(valor1), vbCritical, "Error crítico !!! "
                Call GrabarLog(" Diferencia entre Banco y Cta nro. interno: " + Str(vnrointerno), "Error", "CajaMoviIE")
                
    End If


    vsql = "delete from bancosmovimientos where nrointerno=" + Str(vnrointerno)
    Call EjecutarScript(vsql, pathDBMySQL)
    
    vsql = "select numero as c from asientos where nrointerno=" + Str(vnrointerno)
    Call EjecutarScript(vsql, pathDBMySQL)
    
    
    
    vnumero = traerDatos2(vsql, "c", pathDBMySQL)
    
    
    
    vsql = "delete from asientosdetalle where numero=" + Str(vnumero)
    Call EjecutarScript(vsql, pathDBMySQL)
    
    vsql = "delete from asientos where nrointerno=" + Str(vnrointerno)
    Call EjecutarScript(vsql, pathDBMySQL)
    
    End

End If


If Err Then Exit Sub

End Sub

