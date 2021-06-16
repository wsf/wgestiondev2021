Attribute VB_Name = "Asiento"
' para llamar a esta función se debe usar esto vNroAsiento = Val(GenerarDato("SELECT MAX(Numero) as NroAsiento FROM Asientos;", "NroAsiento")) + 1

'Public Sub ModuloToAsiento()
'Dim rsAsiento As New ADODB.Recordset, sqlAsiento As String
'
'sqlAsiento = "SELECT * FROM moduloToAsiento"
'
'With rsAsiento
'        .CursorLocation = adUseServer
'        Call .Open(sqlAsiento, ConnDDBB, adOpenStatic, adLockOptimistic)
'
''---------------------------------------------------------------------------------------
'Dim vbandera As Boolean
'Dim vmonto As Double
'Dim vcvalor, vcnombre As String
'Dim vcontrol As Control
'Do Until Not .EOF
'
'
'vbandera = True
'vmonto = 0
'vcnombre1 = EsNulo(TraerDato("moduloToAsiento", "cnombre1", pathDBMySQL))
'vcvalor1 = EsNulo(TraerDato("moduloToAsiento", "cvalor1", pathDBMySQL))
'Set vcontrol = FindControl(vcnombre1)
'
'
'
'' controlo si se cumple condiciones para el cammpo1
'If Trim(vcnombre1) = Trim(vcontrol.Name) Then
'
'    ' veo el valor que tiene
'        ' si es numerico
'            If Val(vcontrol.Text) > 0 Then
'                vmonto = vmonto + Val(vvcvalor1)
'            Else
'         ' si no es numerico verifico que elvalor coincida con el valor de la componente de interfaz asociada
'             If Not Trim(vvalor1) = Trim(vcontrol.Text) Then
'                vbandera = False
'             End If
'
'End If
'
'
'
'' controlo si se cumplen condiciones para el campo 2
'
'If Trim(vcnombre2) = Trim(vcontro2.Name) Then
'
'    ' veo el valor que tiene
'        ' si es numerico
'            If Val(vcontrol.Text) > 0 Then
'                vmonto = vmonto + Val(vvcvalor2)
'            Else
'         ' si no es numerico verifico que elvalor coincida con el valor de la componente de interfaz asociada
'             If Not Trim(vvalor2) = Trim(vcontro2.Text) Then
'                vbandera = False
'             End If
'
'End If
'
'
'' veo si se cumplieron las dos condiciones
'
'If vbandera Then
'' guardo el renglon del asiento con el monto calculo
'
''nuevoRenglonAsiento()
'
'End If
'
'
'.MoveNext
'Loop
''---------------------------------------------------------------------------------------
'End With
'End Function
Private Function FindControl(ByVal ControlName As String, vform As Form) As control
Dim ctr As control

For Each ctr In vform.Controls
If ctr.Name = ControlName Then
    MsgBox ctr.Name
    Set FindControl = ctr
End If
Next ctr

End Function

Public Sub nuevoRenglonAsiento(vnumero As Long, vfecha As Date, vnrobalance As Integer, vnrointerno As Integer, vleyenda As String, vlinea As Integer, vCodigoCuenta As String, vdebe As Double, vhaber As Double, vCodigoProveedor As String, vcp As String)
        Dim sqlCampos, sqlValores As String
        
       ' abrirAsiento (vNumero)
        
        sqlCampos = "Numero,Linea,CodigoCuenta,Debe,Haber, LeyendaBancoCaja,codpersona,cp"
        sqlValores = vnumero + "," + vlinea + "," + vCodigoCuenta + "," + vdebe + "," + vhaber + "," + vleyenda + "," + vCodigoProveedor + "," + vcp
        
        Call EjecutarScript("INSERT INTO AsientosDetalle (" + sqlCampos + ") VALUES (" + sqlValores + ")")
End Sub
Public Function bancoToCuenta(vidbanco As String) As String
bancoToCuenta = TraerDato(bancos, "idBancos='" + Str(vidbanco) + "'", pathDBMySQL)
End Function

Public Sub abrirAsiento(vnumero As Integer, vfecha As Date, vnroasiento As Integer, vtipomovimiento As String, vleyenda As String, vnrointerno As Integer)
On Error Resume Next
    
    Dim rsAsiento As New ADODB.Recordset, sqlAsiento As String
    Dim rsAsientoDetalle As New ADODB.Recordset, sqlAsientoDetalle As String
    
    sqlAsiento = "SELECT * FROM Asientos WHERE numero=" + Str(vnumero)
    
    'vNroAsiento = Val(GenerarDato("SELECT MAX(Numero) as NroAsiento FROM Asientos;", "NroAsiento")) + 1
    
    With rsAsiento
        .CursorLocation = adUseServer
        Call .Open(sqlAsiento, ConnDDBB, adOpenStatic, adLockOptimistic)
        
        If .EOF = True Then
            .AddNew ' crea el asiento vacio
        Else
            Exit Sub ' esto ocurre porque el asiento está abierto
        End If
        .Fields("Fecha").Value = vfecha
        .Fields("Numero").Value = vnroasiento
        .Fields("Leyenda").Value = vleyenda
        .Fields("TipoMovimiento").Value = EsNulo(vtipomovimiento)
        .Fields("NroBalance").Value = TraerDato("balances", " Activo='S' order by NroBalance Desc", "NroBalance", pathDBMySQL) + 1
        .Fields("NroInterno").Value = vnrointerno
    
        .Update
        
    End With
    
    sqlAsiento = ""

    If rsAsiento.State = 1 Then
        rsAsiento.Close
        Set rsAsiento = Nothing
    End If

If Err Then GrabarLog "GuardarAsiento", Err.Number & " " & Err.Description, "AbrirAsiento"
End Sub


'getNroAsiento

'getBalanceActual


