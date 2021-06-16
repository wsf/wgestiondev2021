Attribute VB_Name = "Control_CierreCaja"
Public Function valComparaSalCajaAsiento(ByVal idAdesde As Long, ByVal idAhasta As Long, ByVal idBdesde As Long, ByVal idBhasta As Long, Optional ByVal vidanomes) As Boolean
Dim vsaldoCaja, vsaldoCta, vsaldo As Double

Dim vsql1, vsql2 As String

vsql1 = " select " + _
" sum(debe) - sum(haber) as c " + _
" From " + _
" asientos a " + _
" inner join asientosdetalle ad " + _
" on a.Numero = ad.Numero " + _
" Where a.idAsientos > " + Str(idAdesde) + " and a.idAsientos <= " + Str(idAhasta)


vsql2 = " select " + _
" sum(t.Debito) - sum(t.Credito) as c " + _
" From " + _
" bancosmovimientos t " + _
" inner join bancos tt on t.idBancos = tt.idBancos  " + _
" Where " + _
" not t.idBancosMovimientos = '098' and " + _
" not tt.EsCaja = 'B' and " + _
" t.idBancosMovimientos > " + Str(idBdesde) + " and  t.idBancosMovimientos  <= " + Str(idBhasta)


vsaldoCaja = traerDatos2(vsql2, "c", pathDBMySQL)
vsaldoCta = traerDatos2(vsql1, "c", pathDBMySQL)
vsaldo = Val(vsaldoCaja) - Val(vsaldoCta)


If Not vsaldo = 0 And vidanomes = 0 Then
    MsgBox "No es posible cerrar la caja. " + Chr(13) + _
    " > El saldo de cierre de Ejecución Contable    es de: " + Format(vsaldoCta, "###,###,##0.00") + Chr(13) + _
    " > El saldo de cierre de Composición de Caja   es de: " + Format(vsaldoCaja, "###,###,##0.00") + Chr(13) + _
    " > Hay una diferencia de: " + Format(vsaldo, "###,###,##0.00") + Chr(13) + _
    " Revise los movimientos cargado en el día "
    
    frmBancoCajaDetalle.tabbc.TabIndex = 0
    frmBancoCajaDetalle.Refresh
    valComparaSalCajaAsiento = True
    
End If


End Function


Function valCajaCerradaMovCaja(vfecha2 As Date)
On Error Resume Next
Dim vsql As String
Dim vfecha As Date
Dim vid As Long

valCajaCerradaMovCaja = True

vsql = "select max(t.idhasta) as c  from t_cajacierre t"
vid = Val(traerDatos2(vsql, "c", pathDBMySQL))

vsql = "select * from bancosmovimientos where idbancosmovimientos = " + Str(vid)
vfecha = CDate(traerDatos2(vsql, "fecha", pathDBMySQL))

If vfecha2 < vfecha Then
    MsgBox "No se puede realizar la operación.  La caja está cerrada", vbInformation, "Caja Cerrada."
    valCajaCerradaMovCaja = False
End If


If vfecha2 > Date Then
    MsgBox "Está ingresando un movimiento con fecha superior a la del día"
    valCajaCerradaMovCaja = False
End If


End Function

Public Sub listar_nrocomprabantes_noimputados(vidd As Long, vidH As Long, Optional ByVal vfd As Date, Optional ByVal vfh As Date)
On Error Resume Next

Dim vsql, vwhere As String
Dim rs22 As New ADODB.Recordset

Dim vn1, vn2, vcontador As Long

Dim aaa(4000) As String
    
Dim vflag As String

vcontador = 0

If vidd = 0 And vidH = 0 Then

    vwhere = " fecha >= '" + strfechaMySQL(vfd) + "' and fecha  <= '" + strfechaMySQL(vfh) + "' and   nrocomprobante > 1"
   Else
    vwhere = " idbancosmovimientos >= " + Str(vidd) + " and idbancosmovimientos <= " + Str(vidH)
End If

vsql = "select nrocomprobante  as c from bancosmovimientos where " + vwhere + " group by nrocomprobante order by nrocomprobante asc"

With rs22
        .CursorLocation = adUseClient
        Call .Open(vsql, ConnDDBB, adOpenStatic, adLockReadOnly)
End With

rs22.MoveFirst
vn1 = rs22.Fields("c")

rs22.MoveLast
vn2 = rs22.Fields("c")

MsgBox " Desde : " + Str(vn1) + "  Hasta: " + Str(vn2)

Dim j As Integer

j = 0


Debug.Print "------------------------------------------------------------------"

Debug.Print "Nro inicial :" + Str(vn1)

Debug.Print "Nro final :" + Str(vn2)

Debug.Print "------------------------------------------------------------------"

Dim i As Long
i = vn1

For i = vn1 To vn2
    
    vsql = "select nrocomprobante as c from bancosmovimientos where nrocomprobante = " + Str(i)
    vflag = traerDatos2(vsql, "c", pathDBMySQL)
    
    If vflag = "" Then
        ' no encontrado
        
        aaa(j - 1) = Trim(Str(i)) ' pene los nros de comprobantes no encontrados
        Debug.Print " ****  No encontrado : " + aaa(j - 1) + " *********************** "
        j = j + 1
       ' Debug.Print " ****  No encontrado : " + aaa(j - 1) + " *********************** "
        
        Debug.Print " ****  No encontrado : " + Str(i) + " *********************** "
    End If
    
    Debug.Print " > Nros correlativos analizado: " + Str(i)

Next


' insertar los nrodecomprobantes inexistentes
insertarNroComprobantes aaa


If Err < 0 Then
    MsgBox "Presione tecla para continuar"
    Exit Sub
End If


End Sub

Private Sub insertarNroComprobantes(a As Variant)
On Error Resume Next
Dim i As Integer


Dim vcampo, vvalores, vsql   As String


vcampo = "c1,c2"

Debug.Print " =============================== "

frmBancoCajaDetalle.log.Clear

'vsql = "delete from movibca "
'Call EjecutarScript(vsql, PathDBListados)

frmBancoCajaDetalle.vvbarra.Max = UBound(a) + 1

frmBancoCajaDetalle.vvbarra.Value = 0

frmBancoCajaDetalle.gnonro.ColWidth(0) = 1000
frmBancoCajaDetalle.gnonro.ColWidth(1) = 1000


For i = LBound(a) To UBound(a)
    vvalores = "'**', 'Comprobante Nro: [" + (a(i)) + "] no fue emitido' "
    vsql = "insert into movibca (" + vcampo + ") values (" + vvalores + ")"
    
    If Val(a(i)) > 0 Then
        Call EjecutarScript(vsql, PathDBListados)
        frmBancoCajaDetalle.log.AddItem (vvalores)
        frmBancoCajaDetalle.gnonro.AddItem " - " + vbTab + (a(i))
    End If
    Debug.Print " !!!!! Arreglo : " + Str(a(i))
    
    frmBancoCajaDetalle.vvbarra.Value = i
Next

Debug.Print " =============================== "


If Err Then Exit Sub
End Sub
