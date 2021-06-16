Attribute VB_Name = "Pago_Link"
'-----------------------------------------------------------------------------------------------------------------------
'   BRL- COMUNA DE WHEELWRIGHT TGI
'   El nombre de archivo debe respetar la siguiente nomenclatura: PEEEVMDD.
'   Donde
'   P:      Fijo
'   EEE:    Código del ente
'   V:      Identificación numérica consecutiva de volumen
'   M:      Identificación del mes al que corresponde la información ( a partir del mes 10 se usan letras desde la A )
'   DD.:    Identificación del día al que corresponde la información.
'-----------------------------------------------------------------------------------------------------------------------
' Variables gloobales '

Public vfecha_actual, vfechas, vfvencimiento1, vfvencimiento2, vfecha_pago, vfecha_ultimo_vencimiento As Date


Public Canal%, i%


'Public _
'vid_deuda, _
'vid_usuario, _
'vimporte1, _
'vimporte2, _
'vdiscrecional, _
'vid_documento, _
'vcant_registro, _
'vtotal1, _
'vtotal2, _
'vtotal3, _
'vcantidad_registros_extract, _
'vcantidad_total_registros_refresh, _
'vcantidad_registros_lotes_refresg _
'As Long



Public _
vid_deuda, _
vid_usuario, _
vimporte1, _
vimporte2, _
vdiscrecional, _
vid_documento, _
vcant_registro, _
vtotal1, _
vtotal2, _
vtotal3, _
vcantidad_registros_extract, _
vcantidad_total_registros_refresh, _
vcantidad_registros_lotes _
As Long

Public vdiscrecional_extract As String   ' ver
Public vnombre_archivo_control, vcodigo_ente, vnombre_archivo_refresh As String
Public vnombre_archivo_extract As String





'****************************************
' REFRESH
'***************************************

Public Sub init() ' inicializa todos los valores para la pueba

        vid_deuda = 1
        vid_usuario = 1
        vimporte1 = 11.11
        vimporte2 = 22.22
        vdiscrecional = "vdiscrecional"
        vid_documento = 1
        vcant_registro = 1
        vtotal1 = 11.11
        vtotal2 = 22.22
        vtotal3 = 33.33
        vcantidad_registros_extract = 1
        vcantidad_total_registros_refresh = 1
        vcantidad_registros_lotes = 1
        
        
        vfecha_actual = Date
        vfechas = Date
        vfvencimiento1 = Date
        vfvencimiento2 = Date
        vfecha_pago = Date
        vfecha_ultimo_vencimiento = Date
        
        vcodigo_ente = "AAA"
        
        vnombre_archivo_refresh = Pago_Link.Refresh_nombre(vcodigo_ente, 1)
        
        vnombre_archivo_extract = Pago_Link.Extract_nombre(vcodigo_ente)
        
        vnombre_archivo_control = Pago_Link.Control_nombre(vcodigo_ente, "C")

        
End Sub


Public Sub prueba1()
    Open App.Path + "\PagoLink\" + "A" For Output As 1
    Write #1, vl;
    Close 1
End Sub


Public Sub init_archivos()
Pago_Link.abrir_archivo (vnombre_archivo_refresh)
End Sub


Public Function Refresh_nombre(ByVal veee As String, Optional vvolumen As Integer) As String

        Dim vmes As Integer

        Dim eee, v, m, dd As String

        If Val(vvolumen) > 1 Then v = "0"
                
        ' --------------dd------------------'
        dd = Format(Day(Date), "00")
        ' ------------ mes -----------------

                vmes = Month(Date)

                m = Format(vmes, "00")

                If vmes = 10 Then m = "A"

                If vmes = 11 Then m = "B"

                If vmes = 12 Then m = "C"


        '-------- juntar los elementos ----------'
        Refresh_nombre = veee + v + b + dd
        '--------------------v--------------------'
        
        Debug.Print "- Nombre: " + Refresh_nombre
        
        
End Function


Public Function Control_nombre(ByVal veee As String, Optional vfijo As String) As String

        Dim vmes As Integer

        Dim eee, v, m, dd As String

        If vfijo = "" Then vfijo = "C"
                
        ' --------------dd------------------'
        dd = Format(Day(Date), "00")
        ' ------------ mes -----------------

                vmes = Month(Date)

                m = Format(vmes, "00")

                If vmes = 10 Then m = "A"

                If vmes = 11 Then m = "B"

                If vmes = 12 Then m = "C"


        '-------- juntar los elementos ----------'
        Control_nombre = Trim(veee) + Trim(vfijo) + b + dd
        '----------------------------------------'
        
        Debug.Print "- Nombre: " + hacer_nombre_Refresh
        
        
End Function



Public Function Extract_nombre(ByVal veee As String) As String

        Dim vmm, vdd As String
        
       ' Const vdia = Date
        
        vmm = Format(Month(Date), "00")
        
        vdd = Day(Date)
        
        Extract_nombre = "0" + Trim(veee) + vmm + vdd
        
        Debug.Print "- Extract_nombre: " + Extract_nombre

        
End Function


Public Sub Refresh_Inicial()

Dim vline As String
vline = fc("HRFACTURACION", 13) + _
    fc(veee, 3) + _
    FF(vfecha) + _
    fn(vnro_lote, 5, "N") + _
    fc(" ", 104)


    Debug.Print "- Refresh_Inicial : " + vline
    
    Call Pago_Link.grabar_linea(vline)
    
    
End Sub


Public Sub Refresh_Datos()

Dim vline As String
        

vline = fn(vid_deuda, 5, "N") + _
        fn(vtipo_tgi, 3, "N") + _
        fc(vid_usuario, 19) + _
        FF(vfvencimiento1) + _
        fn(vimporte1, 12, "S") + _
        FF(vfvencimiento2) + _
        fn(vimporte2, 12, "S") + _
        fc(vdiscrecional, 50) + _
        fn(vid_documento, 13, "N") + _
        fn(vcant_registro, 8, "N") + _
        fn(vtotal1, 18, "S") + _
        fn(vtotal2, 18, "S") + _
        fn(vtotal3, 18, "S") + _
        fc(" ", 50)


        Debug.Print "- Refresh_Datos : " + vline

        Call Pago_Link.grabar_linea(vline)

End Sub

Public Sub Refresh_Final()

        Dim vline As String

        vline = fc("TRFACTURACION", 13) + _
        fn(vcantidad_registros_refresh, 8, "N") + _
        fn(vtotal1, 18, "S") + _
        fn(vtotal2, 18, "S") + _
        fn(vtotal3, 18, "S") + _
        fc("", 55)

        Debug.Print "- Refresh_Final : " + vline
        
        Call Pago_Link.grabar_linea(vline)

End Sub


'
'Public Sub Extract_Inicial()
'
'Dim vline As String
'
'vline = fn(0, 1, "N") + _
'        fc("BRL", 3) + _
'        FF(vfecha) + _
'        fc("", 86)
'End Sub



'****************************************
' EXTRACT
'***************************************


Public Sub Extract_Datos()

        Dim vline As String
        
        vline = fn(0, 1, "N") + _
        fn(vid_deuda, 5, "N") + _
        fn(vtipo_tgi, 3, "N") + _
        fc(vid_usuario, 19) + _
        fn(vimporte1, 12, "S") + _
        FF(vfecha_pago) + _
        fc(vdiscrecional_extract, 50)
        
        Debug.Print "- Extract_Datos : " + vline
        
        Call Pago_Link.grabar_linea(vline)
        

End Sub

Public Sub Extract_Final()

        Dim vline As String

        vline = fn(2, 1, "N") + _
        fn(vcantidad_registros_extract, 6, "N") + _
        fn(vtotal, 16, "N") + _
        fn(0, 50, "N")

        Debug.Print "- Extract_Final : " + vline
        
        Call Pago_Link.grabar_linea(vline)

End Sub



Public Sub Extract_Inicial()

        Dim vline As String

        vline = fc("HRPASCTRL", 9) + _
        Format(vfecha_actual, "yyyymmdd") + _
        fc(vcodigo_ente, 3) + _
        fc(vnombre_archivo_refresh, 8)
        
        
        Debug.Print "- Extract_Inicial : " + vline
        
        Call Pago_Link.grabar_linea(vline)

End Sub


'****************************************
' Control
'***************************************

Public Sub Control_Inicial()

        Dim vline As String
        vline = fc("HRPASCTRL", 9) + _
        Format(vfecha_actual, "yyyymmdd") + _
        fc(vcodigo_ente, 3) + _
        fc(vnombre_archivo_control, 8) + _
        fn(vlongitud_refresh, 10, "N") + _
        fc("", 37)
        
        
        Debug.Print "- Control_Inicial : " + vline
       
        Call Pago_Link.grabar_linea(vline)
       
End Sub

Public Sub Control_Datos()
        Dim vline As String
        vline = fc("LOTES", 5) + _
        fn(vnro_lote_refresh, 5, "N") + _
        fn(vcantidad_registros_lotes, 8, "N") + _
        fc(vnombre_archivo_control, 8) + _
        fn(vimporte1, 18, "S") + _
        fn(vimporte2, 18, "S") + _
        fn(vimporte3, 18, "S") + _
        fc("", 3)
        
        Debug.Print "- Control_Datos : " + vline
       Call Pago_Link.grabar_linea(vline)
        
End Sub


Public Sub Control_Final()

        Dim vline As String
        vline = fc("FINAL", 5) + _
        fn(vcantidad_total_registros_refresh, 8, "N") + _
        fn(vimporte1, 18, "S") + _
        fn(vimporte2, 18, "S") + _
        fn(vimporte3, 18, "S") + _
        Format(vfecha_ultimo_vencimiento, "yymmdd")
        
        Debug.Print "- Control_Final : " + vline
        
        Call Pago_Link.grabar_linea(vline)
        
      '  Call Pago_Link.grabar_linea(vline)
        
        
        
End Sub

Public Sub abrir_archivo(v As String)
On Error Resume Next
    If v = "" Then v = "Vacio"
    Open App.Path + "\PagoLink\" + v For Output As 1
If Err Then Exit Sub
End Sub

Public Sub cerrar_archivo()
   Close 1
End Sub

Public Sub grabar_linea(vl As String)
On Error Resume Next
   ' Write #1, vl;
    Print #1, vl;
    'Call log.AddItem vl
If Err Then Exit Sub
End Sub
