Attribute VB_Name = "FeModulo"
Option Explicit

Dim vgmodotest As Integer
Dim vcuit, vcertificado As String
Dim TipoDocumento_CUIT, tipoComprobante As Integer
Dim modoFiscal As Variant
Dim gfe As WSAFIPFEx
Public strTicket, vnombre_archivo As String
Public abortarFactura As Boolean
Public vnroempresa As Integer


Public eeEmpresa, eeDireccion, eeCuit, eeLocalidad, eeOtros, eeIngresosBrutos As String

Public Type lineVentaCBTE
    iva105 As Double
    iva21 As Double
    iva27 As Double
    total As Double
    
End Type

Public Type lineVentaAlicuotas

    iva105 As Double
    iva21 As Double
    iva27 As Double
    total As Double
    
End Type

Public Function verificacionCodigoBarra(vnro As String)
verificacionCodigoBarra = "0"
Dim i, v, t, t2 As Integer


t = 0
v = 0
t2 = 0


'1

For i = 1 To Len(vnro) Step 2

    v = Mid(vnro, i, 1)
    t = t + Val(v)

Next


'2

t = t * 3


'3


For i = 2 To Len(vnro) Step 2

    v = Mid(vnro, i, 1)
    t2 = t2 + Val(v)

Next




'4

Dim t3 As Integer
t3 = t + t2

'5

Dim resp As Integer

For i = 1 To 9
    If ((t3 + i) Mod 10) = 0 Then resp = i
Next


verificacionCodigoBarra = resp


End Function


'Const vcuit = "20249182940"
'Const vcertificado = "sartorio5.pfx"
'Const modoFiscal_Test = 0
'Const modoFiscal_Produccion = 1
'Const modoFiscal_Fiscal = 1
'Const TipoDocumento_CUIT = 80
'Const TipoComprobante_FacturaA = 1
'Const TipoComprobante_FacturaB = 6

Private Sub initConstante()
            vcuit = LeerXml("vcuit")
            vcertificado = LeerXml("vcertificado")
            modoFiscal = LeerXml("modoFiscal")
            If vgmodotest = 0 Then modoFiscal = 0
            TipoDocumento_CUIT = LeerXml("TipoDocumento_CUIT")
End Sub



Public Sub fecae(ByRef fe As WSAFIPFEx, ByVal vTipoComprobante, ByVal vnrocomprobante, ByVal vneto, ByVal vtotal, ByVal _
vcuitCliente, ByVal fecha, ByVal vPuntoDeVenta, ByVal videntificador, ByRef vcae, ByRef vcaeFecha, ByVal viva1, ByVal viva2, ByVal viva3, Optional vmodotest, Optional vnrointerno)
On Error Resume Next

Dim vcantiIVA, vultimoNroComprobante As Integer
Dim vauxi As String


Screen.MousePointer = vbHourglass

If Not (vTipoComprobante > 0) Or UCase(LeerXml("ObtieneCAE")) = "NO" Then
    vcae = ""
    vcaeFecha = ""
    Exit Sub
End If

vgmodotest = vmodotest

initConstante
  
' Documentación en: https://sites.google.com/site/facturaelectronicax/documentacion-wsfev1/wsfev1/wsfev1-metodos
  
  Dim bResultado As Boolean
  Dim cIdentificador As String
  Dim v As Variant
  
  v = Test
  
  
 ' bResultado = fe.iniciar(Trim(LeerXml("modoFiscal")), Trim(vCuit), vcertificado, "WSAFIPFE.lic")    ' Paso 1
' '--------------- comentado -------------------------------------------------------
'  If Trim(LeerXml("modoFiscal")) = "1" Then
'    bResultado = fe.iniciar(1, Trim(vCuit), Trim(vcertificado), "WSAFIPFE.lic")
'  Else
'    bResultado = fe.iniciar(1, Trim(vCuit), Trim(vcertificado), "WSAFIPFE.lic")
'    End If
'
'  If bResultado Then
'    If Not fe.f1TicketEsValido Then bResultado = fe.ObtenerTicketAcceso()
'     'bResultado = fe.ObtenerTicketAcceso()
'  End If
'
' -------------------------------------------------------------------------------
 
 'fe.f1RestaurarTicketAcceso (strTicket)
 
 bResultado = True
 bResultado = fe.f1TicketValido

 
  If bResultado Then
  
     fe.F1CabeceraCantReg = 1 ''
     fe.F1CabeceraCbteTipo = 1 ''
     fe.FECabeceraPresta_serv = 1
     
    ' fe.F1DetalleCbteDesde = 1
   '  fe.F1DetalleCbteHasta = 1
  
      
     fe.f1Indice = 0 ''
    
    ' fe.F1DetalleMonId = "PES"
     'fe.F1DetalleMonCotiz = 1
     fe.F1CabeceraPtoVta = 2 ''
     fe.FEDetalleFecha_vence_pago = "20150801"
    ' fe.F1DetalleFchServDesde = "20150701"
    ' fe.F1DetalleFchServHasta = "20150701"
     fe.FEDetalleImp_neto = Val(vneto)
     fe.FEDetalleImp_total = Val(vtotal)
     fe.FEDetalleFecha_cbte = fecha
     fe.FEDetalleNro_doc = vcuit
     'fe.F1DetalleDocNro = vCuit
     fe.FEDetalleTipo_doc = TipoDocumento_CUIT
     
     fe.F1DetalleDocTipo = TipoDocumento_CUIT
         '
    
   ' cIdentificador = Str(videntificador) ' no tiene importancia para la versión WSFEv1
     cIdentificador = "1"
     ' TipoComprobante_FacturaA = 1, TipoComprobante_FacturaB = 6
     
    '--------------------------------------------------------------------------------------------
    '------------IVA-----------------------------------------------------------------------------
    '--------------------------------------------------------------------------------------------
    
     ' Iva 10.5 - 21 - 27
     vcantiIVA = 1
    
     'fe.F1DetalleIvaItemCantidad = vcantiIVA
       
     If viva1 > 0 Then
        vcantiIVA = vcantiIVA + 1
        fe.f1IndiceItem = vcantiIVA - 1
         fe.F1DetalleIvaId = 4
          fe.F1DetalleIvaBaseImp = viva1
          fe.F1DetalleIvaImporte = (viva1 / 10.5) * 100
     
     End If
     
     If viva2 > 0 Then
        vcantiIVA = vcantiIVA + 1
        fe.f1IndiceItem = vcantiIVA - 1
         fe.F1DetalleIvaId = 5
          fe.F1DetalleIvaBaseImp = viva2
          fe.F1DetalleIvaImporte = (viva2 / 21) * 100
     End If
     
     If viva3 > 0 Then
        vcantiIVA = vcantiIVA + 1
        fe.f1IndiceItem = vcantiIVA - 1
        fe.F1DetalleIvaId = 6
        fe.F1DetalleIvaBaseImp = viva3
        fe.F1DetalleIvaImporte = (viva3 / 27) * 100
        
     End If
     
    'fe.F1DetalleIvaItemCantidad = vcantiIVA
    'fe.F1DetalleIvaItemCantidad = Trim(vcantiIVA)
    '---------------------------------------------------------------------------
    
    If vnrointerno > 0 Then cIdentificador = vnrointerno
    cIdentificador = "1"  ' revisar !
    
OtraVez:
    
     'bResultado = fe.Registrar(Val(vPuntoDeVenta), Trim(vTipoComprobante), cIdentificador)
     
     
     
    ' antes de registrar
       fe.ArchivoXMLRecibido = "c:\recibido.xml"
       fe.ArchivoXMLEnviado = "c:\enviado.xml"
        fe.FECabeceraCantReg = 1
       
       bResultado = fe.f1CAESolicitar()
    
    
    'MsgBox "CAE.: " + fe.FERespuestaDetalleCae
      
      
      
      Debug.Print ("resultado global AFIP: " + fe.F1RespuestaResultado + "   " + fe.UltimoMensajeError)
      Debug.Print ("es reproceso? " + fe.F1RespuestaReProceso)
      Debug.Print ("registros procesados por AFIP: " + Str(fe.F1RespuestaCantidadReg))
      Debug.Print ("error genérico global:" + fe.f1ErrorMsg1)
     
     
     If MsgBox("Registra", vbYesNo) = vbNo Then Exit Sub
     
    
     bResultado = fe.Registrar(Val(vPuntoDeVenta), 1, cIdentificador)
     
   '   bResultado = fe.Registrar(Val(vPuntoDeVenta), Trim(vTipoComprobante), cIdentificador)
     If bResultado Then
        vcae = fe.FERespuestaDetalleCae
        vcaeFecha = fe.FERespuestaDetalleFecha_vto
        
       ' vauxi = fe.FERespuestaDetalleFecha_vto
       ' If Not vauxi = "" Then vcaeFecha = Trim(Right(vauxi, 2) + "/" + Mid(vauxi, 5, 2) + "/" + Left(vauxi, 4))
        
        Debug.Print "CAE : " + Str(vcae) + Chr(13) + " Vto. CAE:  " + Str(vcaeFecha) + Chr(13) + "Ultimo error " + fe.UltimoMensajeError _
        + Chr(13) + "TicketHora: " + fe.TicketHora + Chr(13) + "Ticket:" + fe.UltimoMensajeError + Chr(13) + "Ultimo error " + fe.UltimoMensajeError
        MsgBox "Datos generado por AFIP " + Chr(13) + "-  CAE : " + Str(vcae) + Chr(10) + "-  Vto. CAE:  " + Str(vcaeFecha)
     Else
     
        ' hubo un error
        ' verifico si el problema es el nro
        'If vultimoNroComprobante = 0 Then
        '           vultimoNroComprobante = frmRemito.getNroCompAfip
        '
        '           If Not Val(vPuntoDeVenta) = vultimoNroComprobante + 1 Then
        '               vPuntoDeVenta = vultimoNroComprobante + 1
        '               GoTo OtraVez:
        '           End If
        'End If
     
         MsgBox " Hubo un problema con los datos generados por AFIP " + Chr(13) + "El documento no se puede generar" + Chr(13) _
         + "Motivo: " + fe.FERespuestaDetalleMotivo + Chr(13) + " Detalle: " + fe.UltimoMensajeError
         Chr (13) + "El sistema se cerrará "
        End
        
        Debug.Print "Motivo: " + fe.FERespuestaDetalleMotivo + Chr(13) + " Detalle: " + fe.UltimoMensajeError
        End
       ' MsgBox ("Motivo: " + fe.FERespuestaDetalleMotivo + Chr(10) + " Error " + fe.Permsg + "Detalle: " + fe.UltimoMensajeError)
     End If
  
  
 Else
 
        MsgBox "Error al fe.ObtenerTicketAcceso() : " + fe.UltimoMensajeError + "- " + fe.UltimoNumeroError + Chr(13) + _
        "-Modo: " + Trim(modoFiscal) + "- CUIT" + Trim(vcuit) + " Ultimo error " + fe.UltimoMensajeError
        Debug.Print fe.UltimoMensajeError
  
  End If
  
  Screen.MousePointer = vbDefault
  
If Err < 0 Then
  Screen.MousePointer = vbDefault
    MsgBox "Hubo un problema con el servicio de AFIP para obtener el CAE " + Err.Description
    Exit Sub
End If


End Sub


Public Sub fecae2(ByRef vfe As WSAFIPFEx, ByVal vTipoComprobante, ByVal vnrocomprobante, ByVal vneto, ByVal vtotal, ByVal _
vcuitCliente, ByVal fecha, ByVal vPuntoDeVenta, ByVal videntificador, ByRef vcae, ByRef vcaeFecha, ByVal viva1, ByVal viva2, ByVal viva3, Optional vmodotest, Optional vnrointerno, Optional vViene As String, Optional ByVal vConsumidorFinal As Integer)
On Error Resume Next

Dim viva0, vivas  As Double
Dim vmensaje As String


vivas = (viva1 / 10.5 * 100) + (viva2 / 21 * 100) + (viva3 / 27 * 100)
viva0 = vneto - vivas


                    If LeerXml("CAEFIJO") = "999" Then
                      vcae = 999
                      vcaeFecha = 999
                      Exit Sub
                    End If
          

Dim vcantiIVA, vultimoNroComprobante As Integer
Dim vauxi As String

Dim vsql As String

'Dim fe As WSAFIPFEx
'Set fe = vfe
abortarFactura = False
Dim fe As WSAFIPFEx

If vViene = "frmBuscarFactura" Then
    Set fe = frmBuscarFactura.vfe
Else
    Set fe = frmRemito.fe
End If


Screen.MousePointer = vbHourglass

vcae = ""
vcaeFecha = ""


'vsql = "select SucursalDocVenta as c from configuracion where id = " + Str(vnroempresa + 1)


If Not UCase(LeerXml("Puesto")) = "EMPRESAS" Then
    
    If Not Val(traerDatos2("select * from configuracion", "SucursalDocVenta", PathDBConfig)) = vPuntoDeVenta Then
        Exit Sub
    End If

End If
If Not (vTipoComprobante > 0) Or UCase(LeerXml("ObtieneCAE")) = "NO" Then Exit Sub


vgmodotest = vmodotest

initConstante
  
 
  Dim bResultado As Boolean
  Dim cIdentificador As String
  Dim v As Variant
  
  v = Test

'If fe.f1TicketValido Or Not strTicket = "" Then
If fe.f1TicketValido Then
                        ' usa un ticket válido
                         fe.f1RestaurarTicketAcceso (strTicket)

Else
                        ' nuevo ticket
                    
                         If Trim(LeerXml("modoFiscal")) = "1" Then
                           ' bResultado = fe.iniciar(1, Trim(LeerXml("vcuit")), App.Path + "\" + Trim(LeerXml("vcertificado")), App.Path + "\" + Trim(LeerXml("LicenciaWSAFIP")))
                          
                            bResultado = fe.iniciar(1, Trim(getCuitFE(vnroempresa)), App.Path + "\" + Trim(getCertificadoFE(vnroempresa)), App.Path + "\" + Trim(getLicenciaFE(vnroempresa)))
                          
                          
                          Else
                          '  bResultado = fe.iniciar(0, Trim(LeerXml("vcuit")), App.Path + "\" + Trim(LeerXml("vcertificado")), App.Path + "\" + Trim(LeerXml("LicenciaWSAFIP")))
                           ' bResultado = fe.iniciar(0, Trim(LeerXml("vcuit")), App.Path + "\" + Trim(LeerXml("vcertificado")), "")
                            ' modo prueba
                            
                            bResultado = fe.iniciar(0, "20249182940", "sartorio.pfx", " ")
                            
                            'bResultado = fe.iniciar(0, Trim(getCuitFE(vnroempresa)), App.Path + "\" + Trim(getCertificadoFE(vnroempresa)), App.Path + "\" + Trim(getLicenciaFE(vnroempresa)))
                            
                            'bResultado = fe.iniciar(1, Trim(LeerXml("vcuit")), App.Path + "\" + Trim(LeerXml("vcertificado")), App.Path + "\" + Trim(LeerXml("LicenciaWSAFIP")))
                          
                        End If
                        
                        'bResultado = fe.iniciar(1, "30707384316", App.Path + "\PoliCertificadoProduccion11.pfx", App.Path + "\WSAFIPFE.lic")
                        
                        If Not fe.f1TicketValido Then
                            bResultado = fe.f1ObtenerTicketAcceso()
                        Else
                            bResultado = True
                        End If
                        
                        If Not bResultado Then
                                MsgBox "Presione <Aceptar para continuar>"
                                If Not fe.f1ObtenerTicketAcceso() Then If Not fe.f1ObtenerTicketAcceso() Then Exit Sub
                        End If
                          
                        strTicket = fe.f1GuardarTicketAcceso()
                          
End If
  
      '---- ale 3-1-18 ----------todo sacar panic---------
     ' MsgBox "Ultimos nros de comprobantes"
      
     ' MsgBox fe.f1CompUltimoAutorizado(2, 1)
     ' MsgBox fe.f1CompUltimoAutorizado(2, 6)
      
      '----------------------------------------
        
      fe.F1CabeceraCantReg = 1
      fe.F1CabeceraPtoVta = Val(vPuntoDeVenta)
      fe.F1CabeceraCbteTipo = vTipoComprobante
      fe.f1Indice = 0
      fe.F1DetalleConcepto = 1  ' *
      
      If vConsumidorFinal = 99 Then
        fe.F1DetalleDocTipo = 99
      Else
        fe.F1DetalleDocTipo = Val(TipoDocumento_CUIT)
      
      End If
      
      fe.F1DetalleDocNro = Trim$(vcuitCliente)
    
      fe.F1DetalleCbteDesde = Val(vnrocomprobante)
      fe.F1DetalleCbteHasta = Val(Trim(vnrocomprobante))
    
      fe.F1DetalleCbteFch = fecha
      
      fe.F1DetalleImpTotal = Val(Format(vtotal, "#######.00"))
      
            
      fe.F1DetalleMonId = "PES"
      fe.F1DetalleMonCotiz = 1
    
      ' fe.F1DetalleImpNeto = Val(vneto)
   
    If vTipoComprobante >= 6 Then
        fe.F1DetalleImpOpEx = Val(Format(vtotal, "#######.00"))
    End If
    
    
    
    ' si es una nota de credito 2021
    ' le puse un periodo asociado cualquira
    '
    
    
    
    If vTipoComprobante = 3 Or vTipoComprobante = 8 Or vTipoComprobante = 13 Then
        
        fe.F1DetalleCbtesAsocTipo = vTipoComprobante
        fe.F1DetalleCbtesAsocPtoVta = Val(vPuntoDeVenta)
        
        Dim vnro_comprobante_asociado, vfecha_comp_asociado
        
        vnro_comprobante_asociado = InputBox("Ingresar nro de comprobante asociado")
        vfecha_comp_asociado = InputBox("Fecha del comprobante asociado. Ejemplo: 20210630")
        
        
        'fe.F1DetalleCbtesAsocNroS = vnro_comprobante_asociado
        
        fe.F1DetalleCbtesAsocNro = vnro_comprobante_asociado
        
        fe.F1DetalleCbtesAsocItemCantidad = 1
       
        fe.F1DetalleCbtesAsocFecha = vfecha_comp_asociado
        
        fe.F1DetallePeriodoAsocFchDesde = vfecha_comp_asociado
        fe.F1DetallePeriodoAsocFchHasta = vfecha_comp_asociado

        
    End If
                        
                        
                        
                        
                        
      If vTipoComprobante > 0 And vTipoComprobante < 6 Then
      fe.F1DetalleImpTotalConc = 0
      
        
                         fe.F1DetalleImpNeto = Val(Format(vneto, "#######.00"))
                        
                        'fe.F1DetalleImpOpEx = 0
                      
                       'fe.F1DetalleImpOpEx = 0
                       ' fe.F1DetalleImpTrib = 0
                        fe.F1DetalleImpIva = Val(Format(viva2 + viva1 + viva3, "#######.00"))
       
                        
                         vcantiIVA = 0
                      
                       '' fe.F1DetalleIvaItemCantidad = vcantiIVA + 1
                         
                       If viva1 > 0 Then
                            vcantiIVA = vcantiIVA + 1
                            fe.f1IndiceItem = vcantiIVA - 1
                            fe.F1DetalleIvaId = 4
                            fe.F1DetalleIvaBaseImp = Val(Format(viva1 / 10.5 * 100, "######0.00"))
                            fe.F1DetalleIvaImporte = Val(Format(viva1, "#####0.00"))
                       End If
                       
                       
                    
                       
                       
                       If viva2 > 0 Then
                            vcantiIVA = vcantiIVA + 1
                            fe.f1IndiceItem = vcantiIVA - 1
                            fe.F1DetalleIvaId = 5
                            fe.F1DetalleIvaBaseImp = Val(Format(((viva2 / 21) * 100), "######0.00"))
                            fe.F1DetalleIvaImporte = Val(Format(viva2, "#####0.00"))
                       End If
                       
                       If viva3 > 0 Then
                            vcantiIVA = vcantiIVA + 1
                            fe.f1IndiceItem = vcantiIVA - 1
                            fe.F1DetalleIvaId = 6
                            fe.F1DetalleIvaBaseImp = Val(Format((viva3 / 27) * 100, "######0.00"))
                            fe.F1DetalleIvaImporte = Val(Format(viva3, "#####0.00"))
                       End If
               
                     ' 2019-09-03 --
                     ' ---
                       
                       If Val(Format(viva0, "#####0.00")) > 0 Then
                            vcantiIVA = vcantiIVA + 1
                            fe.f1IndiceItem = vcantiIVA - 1
                            fe.F1DetalleIvaId = 3
                            fe.F1DetalleIvaBaseImp = Val(Format(viva0, "#####0.00"))
                            fe.F1DetalleIvaImporte = Val(Format(0, "#####0.00"))
                       End If
                     
                     ' ---
                           
                           
                       fe.F1DetalleIvaItemCantidad = vcantiIVA
                           
    End If
                           
      fe.F1DetalleCbtesAsocItemCantidad = 0
      
      fe.F1DetalleOpcionalItemCantidad = 0
    
      
     vnombre_archivo = ""
     
     vnombre_archivo = Str(Val(vPuntoDeVenta)) + "-" + Str(vTipoComprobante) + "-" + Str(Val(TipoDocumento_CUIT)) + "-" + Trim$(vcuitCliente) + "-" + Str(Val(vnrocomprobante)) + "-" + Str(Val(Trim(vnrocomprobante)))
               
     
     ' ***********************************
     fe.ArchivoXMLRecibido = App.Path + "\Log\recibido" + Trim(vnombre_archivo) + ".xml"
     fe.ArchivoXMLEnviado = App.Path + "\Log\enviado" + Trim(vnombre_archivo) + ".xml"
    '--------------------------------
    
    ' qrqr
  
    vqrnombre = Trim(vcuitCliente) + Trim(vnrocomprobante)
    fe.F1Detalleqrarchivo = App.Path + "\" + vqrnombre + ".jpg"
    fe.F1Detalleqrformato = 6
    fe.F1Detalleqrtipocodigo = "E"
    fe.F1Detalleqrtolerancia = 1
    fe.F1Detalleqrresolucion = 2
     
     Debug.Print (vqrnombre)
     

     
     If fe.f1CAESolicitar() Then
           
                    vcae = fe.F1RespuestaDetalleCae
                    vcaeFecha = fe.F1RespuestaDetalleCAEFchVto
                     
                     
                     If Val(vcae) = 0 Then
                        
                        MsgBox "No se pudo generar el CAE. " + Chr(13) + " Mensaje de error de AFIP: " + fe.UltimoMensajeError
                        vcae = ""
                        vcaeFecha = ""
                        
                        If vViene = "frmBuscarFactura" Then
                            MsgBox "No se puede seguir con el proceso de impresión automático" + Chr(13) + _
                            "El sistema se cerrará"
                            End
                        
                        End If
                        
                     End If
                     
                    '----------------------------------------------------------------------------------------------
                     
                     If vViene = "frmBuscarFactura" Then
                            
                            frmBuscarFactura.log2.Visible = True
                            frmBuscarFactura.log2.AddItem "C.A.E :" + vcae _
                            + vbTab + "Vto C.A.E :" + Str(vcaeFecha) + vbTab + _
                            Str(vcuitCliente) + vbTab + _
                            Str(vnrocomprobante) + vbTab + _
                            Str(vTipoComprobante) + vbTab + _
                            Format(fecha, "dd/mm/yyyy")
                            
                           Call log2("C.A.E :" + vcae _
                            + vbTab + "Vto C.A.E :" + Str(vcaeFecha) + vbTab + _
                            Str(vcuitCliente) + vbTab + _
                            Str(vnrocomprobante) + vbTab + _
                            Str(vTipoComprobante) + vbTab + _
                            Format(fecha, "dd/mm/yyyy"))
                                                  
     
                     Else
                     
                            MsgBox "C.A.E :" + vcae _
                            + Chr(13) + "Vto C.A.E :" + Str(vcaeFecha) _
                            + Chr(13) + "Presione <Aceptar> para continuar", vbMsgBoxRtlReading, "Certificado por AFIP"
                            
                             Call log2("C.A.E :" + vcae _
                            + vbTab + "Vto C.A.E :" + Str(vcaeFecha) + vbTab + _
                            Str(vcuitCliente) + vbTab + _
                            Str(vnrocomprobante) + vbTab + _
                            Str(vTipoComprobante) + vbTab + _
                            Format(fecha, "dd/mm/yyyy"))
                            
                    End If
                   
                     
     Else
                                    
                                    
                    ' todo sis estoy acá es porque no se puedo obtener el CAE
                    ' deberia ver que pasó
                    '
                    
                    For i = 1 To 10
                    
                                        MsgBox "AFIP no responde. " + Chr(13) + "Presione <Aceptar> para volver a intentar" _
                                        + Chr(13) + "Esta acción puede ser ejecutada hasta 10 veces"
                                                                                
                                        vcae = fe.F1RespuestaDetalleCae
                                        vcaeFecha = fe.F1RespuestaDetalleCAEFchVto
                                        
                                        If Val(vcae) > 0 Then
                                            Exit For
                                        End If
                    
                    Next
                    
                    
                    vcae = 0
                    vcaeFecha = 0
                    
                    If LeerXml("CAEFIJO") = 1 Then
                      vcae = 999
                      vcaeFecha = 999
                    End If
          
                If vViene = "frmBuscarFactura" Then
                        
                            frmBuscarFactura.log2.Visible = True
                            frmBuscarFactura.log2.AddItem "Error :" + fe.UltimoMensajeError + vbTab + _
                            Str(fe.f1CompUltimoAutorizado(2, 1)) + vbTab + _
                            Str(vcuitCliente) + vbTab + _
                            Str(vnrocomprobante) + vbTab + _
                            Str(vTipoComprobante) + vbTab + _
                            Format(vTipoComprobante, "dd/mm/yyyy")
                                            
                            Call log2("Mensaje AFIP FE :" + fe.UltimoMensajeError + vbTab + _
                            Str(fe.f1CompUltimoAutorizado(2, 1)) + vbTab + _
                            Str(vcuitCliente) + vbTab + _
                            Str(vnrocomprobante) + vbTab + _
                            Str(vTipoComprobante) + vbTab + _
                            Format(fecha, "dd/mm/yyyy"))
                            
                            
                            If vViene = "frmBuscarFactura" Then
                            MsgBox "No se puede seguir con el proceso de impresión automático" + Chr(13) + _
                            "Vuelva a retomarlo"
                            
                             abortarFactura = True
                             Exit Sub
                        
                            End If

                            
                    
                    Else
                        'MsgBox "No se pudo obtener el CAE. Intente nuevamente"
            
                           MsgBox "No se pudo obtener el CAE. Intente nuevamente " + Chr(13) + _
                           "Error :" + Trim(fe.UltimoMensajeError)
                            vcae = 0
                            vcaeFecha = 0
                           
                           
                           abortarFactura = True
                           
                            Call log2("Mensaje AFIP FE :" + fe.UltimoMensajeError + vbTab + _
                            Str(vcuitCliente) + vbTab + _
                            Str(vnrocomprobante) + vbTab + _
                            Str(vTipoComprobante) + vbTab + _
                            Format(fecha, "dd/mm/yyyy"))
                            
                End If
     
     
     End If
    

     
  Screen.MousePointer = vbDefault
  
If Err < 0 Then
    Screen.MousePointer = vbDefault
    MsgBox "Hubo un problema con el servicio de AFIP para obtener el CAE " + Err.Description
    Exit Sub
End If


End Sub



Public Sub getStatusAfip2(ByRef vnfa, ByRef vnfb)
On Error Resume Next

Dim vcantiIVA, vultimoNroComprobante, vPtoVta2 As Integer
Dim bResultado As Boolean
Dim cIdentificador, vsql, vultimoMensajeError  As String
Dim v As Variant

Screen.MousePointer = vbHourglass

' Documentación en: https://sites.google.com/site/facturaelectronicax/documentacion-wsfev1/wsfev1/wsfev1-metodos
v = Test
  
If Trim(LeerXml("modoFiscal")) = "1" Then
    bResultado = frmPrincipal.fe.iniciar(1, Trim(LeerXml("vcuit")), App.Path + "\" + Trim(LeerXml("vcertificado")), App.Path + "\" + Trim(LeerXml("LicenciaWSAFIP")))
 Else
    bResultado = frmPrincipal.fe.iniciar(0, "20249182940", "sartorio.pfx", " ")
End If


vsql = "select SucursalDocVenta as c from configuracion limit 1"
vPtoVta2 = traerDatos2("select * from configuracion", "SucursalDocVenta", PathDBConfig)
bResultado = frmPrincipal.fe.f1ObtenerTicketAcceso()

vultimoMensajeError = frmPrincipal.fe.UltimoMensajeError
vnfa = frmPrincipal.fe.f1CompUltimoAutorizado(vPtoVta2, 1) + 1
vnfb = frmPrincipal.fe.f1CompUltimoAutorizado(vPtoVta2, 6) + 1

'vnfa = vnroFactA
'vnfb = vnroFactB

'logStatus.Clear

'Me.logStatus.AddItem "Ultimo mensaje WS AFIP: " + vultimoMensajeError + Chr(13)
'Me.logStatus.AddItem "Ultima Factura A: " + Str(vnroFactA)
'Me.logStatus.AddItem "Ultima Factura B: " + Str(vnroFactB)

Screen.MousePointer = vbDefault

If Err < 0 Then
    'getNroCompAfip = 0
    Exit Sub
End If
End Sub

Public Function fn(valor As Variant, espacios As Integer, esdecimal As String) As String
Dim vceros As String
Dim i As Integer

valor = Replace$(valor, "-", "")


If valor = "" Then valor = 0
valor = Abs(valor)

For i = 1 To espacios
    vceros = vceros + "0"
Next

If (UCase(esdecimal) = "S") Then
    valor = valor * 100
Else
    valor = valor
End If

fn = Format(valor, vceros)

End Function
Function str_ancho(valor As String, espacios As Integer) As String
Dim vceros, vcomilla As String
Dim i As Integer
Dim vlen As Integer

vceros = ""
vlen = Len(valor)

If vlen > espacios Then
        valor = Left(valor, espacios)
Else
        For i = 1 To espacios - vlen
            vceros = vceros + " "
        Next
End If

str_ancho = valor + vceros

End Function

Public Function fc(ByVal valor As String, ByVal espacios As Integer) As String
Dim vceros, vcomilla As String
Dim i As Integer
Dim vlen As Integer

valor = Replace(valor, "-", "")

valor = Replace(valor, Chr(34), " ")

vceros = ""
vlen = Len(valor)

If vlen > espacios Then
        valor = Left(valor, espacios)
Else
        For i = 1 To espacios - vlen
            vceros = vceros + " "
        Next
End If

fc = valor + vceros

End Function


Public Function FF(ByVal vfecha As Date) As String
    FF = Format(vfecha, "yyyymmdd")
End Function


Function getTipoDoc(vtipo As Variant) As String

Select Case vtipo

    Case "Fact A"
        getTipoDoc = "1"
    Case "Fact B"
        getTipoDoc = "6"
    Case "Nota C"
        getTipoDoc = "3"
        Case "NotaCB"
        getTipoDoc = "8"
    Case "Nota D"
        getTipoDoc = "2"
    Case "Fact C"
        getTipoDoc = "11"
End Select

End Function
