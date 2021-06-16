Attribute VB_Name = "Procedimientos"
Option Explicit
Public vTotalD, vTotalH, vsaldo As Double      'Variables Temporales para calcular saldos
Public vtimestampAjuste As String

Public Sub mensaje(v As String)
If LeerXml("Debug") = "True" Or LeerXml("Debug") = "TRUE" Or LeerXml("Debug") = "true" Then
    MsgBox v
Else
    Exit Sub
End If
End Sub

Function utltimoFactura() As Long
On Error Resume Next

Dim vsql As String
    vsql = "select max(idfactura) as c from factura"
    utltimoFactura = traerDatos2(vsql, "c", pathDBMySQL)

If Err Then
    utltimoFactura = 0
    Exit Function
End If

End Function

Public Sub mostrargrilla(ByRef grilla As MSHFlexGrid, ByVal vsql As Variant, Optional vw As String)
On Error Resume Next

Dim arr() As String

arr = Split(vw, ",")

Call fijarAnchoGrilla(grilla, arr)


Dim rs4 As New ADODB.Recordset



Call rs4.Open(vsql, ConnComunaDB, adOpenStatic, adLockPessimistic)
Set grilla.DataSource = rs4.DataSource
    

rs4.Close


If Err Then Exit Sub
End Sub

Public Sub fijarAnchoGrilla(ByRef grilla As MSHFlexGrid, ByRef arr() As String)
Dim i As Integer

For i = 1 To UBound(arr)
    grilla.ColWidth(i) = Val(arr(i - 1))
Next

End Sub


Public Sub GuardarRel(ByVal vIdFactura As Long, ByVal vIdEmpresa As Long, ByVal vidVendedor As Long, ByVal vnrointerno As Long)
Dim vsql As String
Dim vcampos, vValor As String

vcampos = "(idFactura, idEmpresa, idVendedor, nrointerno)"
vValor = "(" + Str(vIdFactura) + "," + Str(vIdEmpresa) + "," + Str(vidVendedor) + "," + Str(vnrointerno) + ")"
vsql = "insert into t_rel " + vcampos + " values " + vValor

Call EjecutarScript(vsql, pathDBMySQL)

End Sub

Public Function UltimaFactura() As Long
    UltimaFactura = traerDatos2("select max(idfactura) as c", "c", pathDBMySQL)
End Function

Public Function codigo2id(vcodigo As String) As Long
On Error Resume Next
Dim vsql As String

vsql = "select idProveedores as c from proveedores where codigo = '" + vcodigo + "'"
codigo2id = traerDatos2(vsql, "c", pathDBMySQL)

If Err Then
    codigo2id = 0
    Exit Function
End If

End Function


Public Function existeRegistro(vnrointerno As Long) As Boolean
Dim vvsql, vlog As String

vlog = ""


' ---------------------------------------------------------------------------------

If vnrointerno = 0 Then
    existeRegistro = False
    Exit Function
End If

' controlo asiento ---------------------------------------------------------------

vvsql = "select * from asientos where nrointerno=" + Str(vnrointerno)

If Val(EsNulo(traerDatos2(vvsql, "nrointerno", pathDBMySQL))) > 0 Then
    vlog = vlog + Chr(13) + " - Asiento "
End If


' controlo bancosmovimientos -----------------------------------------------------
vvsql = "select * from bancosmovimientos where nrointerno=" + Str(vnrointerno)

If Val(EsNulo(traerDatos2(vvsql, "nrointerno", pathDBMySQL))) > 0 Then
    vlog = vlog + Chr(13) + " - Movi Cajas / Bancos "
End If



' controlo ivacompra  -----------------------------------------------------
vvsql = "select * from ivafacturacompra where nrointerno=" + Str(vnrointerno)

If Val(EsNulo(traerDatos2(vvsql, "nrointerno", pathDBMySQL))) > 0 Then
    vlog = vlog + Chr(13) + " - Iva Compra "
End If



' controlo ivaventa -----------------------------------------------------
vvsql = "select * from ivafacturaventa where nrointerno=" + Str(vnrointerno)

If Val(EsNulo(traerDatos2(vvsql, "nrointerno", pathDBMySQL))) > 0 Then
    vlog = vlog + Chr(13) + " - Iva Venta "
End If



If Not vlog = "" Then
    MsgBox "No es posible grabar el movimiento. Este número interno pertenece a los siguientes movimientos " + Str(vnrointerno) + vlog, vbCritical
    existeRegistro = True
Else
    existeRegistro = False
End If


End Function
Public Function existeRegistroAsientos(vnrointerno As Long) As Boolean
Dim vvsql, vlog As String

vlog = ""


' ---------------------------------------------------------------------------------

If vnrointerno = 0 Then
    existeRegistroAsientos = False
    Exit Function
End If

' controlo asiento ---------------------------------------------------------------

vvsql = "select * from asientos where nrointerno=" + Str(vnrointerno)

If Val(EsNulo(traerDatos2(vvsql, "nrointerno", pathDBMySQL))) > 0 Then
    vlog = vlog + Chr(13) + " - Asiento "
End If


If Not vlog = "" Then
    MsgBox "No es posible grabar el movimiento. Este número interno pertenece a los siguientes movimientos " + Str(vnrointerno) + vlog, vbCritical
    existeRegistroAsientos = True
Else
    existeRegistroAsientos = False
End If


End Function

Public Sub grabarCheque(ByVal vcampos As String, ByVal vValor As String)  ' ema:
On Error Resume Next

EjecutarScript ("insert into cheques ( " + vcampos + ") values (" + vValor + " )")

Dim i As Long

i = traerDatos2("select max(idCheques) as id from cheques order by idcheques", "id", pathDBMySQL)

Call EjecutarScript("INSERT INTO HistoricoEstadosCheques (idCheques, idEstadoAnterior, idEstadoActual, FechaCambio) VALUES (" & _
Str(i) & ",0,2,'" & strfechaMySQL(Date) & ")')")

If Err Then
    MsgBox "Disculpe. No se pudo guardar el cheque correctamente en el módulo de CHEQUES." + Chr(13) + "Verifique!", vbCritical, "Error módulo de CHEQUES"
    Exit Sub
Else
'Call alerta("El cheque de deposito ha sido guardado en el módulo CHEQUE", 1000)
End If
End Sub
Public Sub UltimaVenta(vcliente As String, vfecha As Date)
On Error Resume Next

    Dim rsUVenta As New ADODB.Recordset, sqlUVenta As String
    
    sqlUVenta = "SELECT * FROM clientes WHERE (codigo = '" & vcliente & "')"
    
    With rsUVenta
        Call .Open(sqlUVenta, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then
            .Fields("U_Venta").Value = vfecha
        Else
            '.Fields("U_Venta").Value = Null
        End If
        
        .Update
    
    End With
    
    sqlUVenta = ""
    
    rsUVenta.Close
    Set rsUVenta = Nothing

If Err Then GrabarLog "UltimaVenta", Err.Number & " " & Err.Description, "Procedimientos"
End Sub
Public Function ControlarUpdate(vClienteControl) As Boolean
On Error Resume Next

    Dim rsControl As New ADODB.Recordset, sqlControl As String
    
    sqlControl = "SELECT * FROM ErrorVista2 WHERE (codigo = '" & vClienteControl & "')"
    
    With rsControl
        Call .Open(sqlControl, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        ControlarUpdate = Not .EOF
        
    End With
    
    sqlControl = ""
    
    rsControl.Close
    Set rsControl = Nothing

If Err Then GrabarLog "ControlarUpdate", Err.Number & " " & Err.Description, "Procedimientos"
End Function
Public Function ControlarClientes() As Boolean
On Error Resume Next

    Dim rsClientes As New ADODB.Recordset, sqlClientes As String

    sqlClientes = "SELECT * FROM clientes WHERE (codigo is null) OR (codigo = '')"
    
    With rsClientes
        Call .Open(sqlClientes, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        ControlarClientes = Not .EOF
        
    End With
    
    sqlClientes = ""
    
    rsClientes.Close
    Set rsClientes = Nothing

If Err Then GrabarLog "ControlarUpdate", Err.Number & " " & Err.Description, "Procedimientos"
End Function

Public Function ControlarDatosTemp(ByRef vtipo, ByRef vIdTemp) As Boolean
On Error Resume Next

    Dim rsControlTemp As New ADODB.Recordset, sqlControlTemp As String

    Select Case vtipo
    
        Case "C"
            sqlControlTemp = "SELECT * FROM Factura_Temp WHERE (id = " & vIdTemp & ")"
        Case "F"
            sqlControlTemp = "SELECT * FROM Factura_Temp WHERE (id = " & vIdTemp & ")"
        Case "D"
            sqlControlTemp = "SELECT * FROM Fdetalle_Temp WHERE (id = " & vIdTemp & ")"
        
    End Select
    
    With rsControlTemp
        Call .Open(sqlControlTemp, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        ControlarDatosTemp = Not .EOF
        
    End With
    
    sqlControlTemp = ""
    
    rsControlTemp.Close
    Set rsControlTemp = Nothing

If Err Then GrabarLog "ControlarDatosMigrados", Err.Number & " " & Err.Description, "Procedimientos"
End Function
Public Function ControlarDatosWGestion(ByRef vtipo, ByRef vIdWGestion) As Boolean
On Error Resume Next

    Dim rsControlWGestion As New ADODB.Recordset, sqlControlWGestion As String
    
    Select Case vtipo
    
        Case "C"
            'sqlControlWGestion = "SELECT * FROM cuentascorrientes WHERE (id = " & vIdWGestion & ")"
        Case "F"
            sqlControlWGestion = "SELECT * FROM Factura WHERE (id = " & vIdWGestion & ")"
        Case "D"
            sqlControlWGestion = "SELECT * FROM Fdetalle WHERE (id = " & vIdWGestion & ")"
        
    End Select
    
    With rsControlWGestion
        Call .Open(sqlControlWGestion, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        ControlarDatosWGestion = Not .EOF
        
    End With
    
    sqlControlWGestion = ""
    
    rsControlWGestion.Close
    Set rsControlWGestion = Nothing


If Err Then GrabarLog "ControlarDatosWGestion", Err.Number & " " & Err.Description, "Procedimientos"
End Function
Public Sub ActualizarRubros()
On Error Resume Next
    
    Dim rsArticulos As New ADODB.Recordset
    Dim sqlArticulos As String
    
    sqlArticulos = "SELECT rubro FROM articulos"
    
    With rsArticulos
        Call .Open(sqlArticulos, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not .EOF = True Then .MoveFirst
        
        Do Until .EOF = True
            .Fields(0).Value = 0
            .MoveNext
        Loop
    
    End With
        
    sqlArticulos = ""
    
    rsArticulos.Close
    Set rsArticulos = Nothing

If Err Then GrabarLog "ActualizarRubros", Err.Number & " " & Err.Description, "Procedimientos"
End Sub
Public Sub Agregar99999()
On Error Resume Next
    
    Dim rsArticulos As New ADODB.Recordset, sqlArticulos As String
    
    sqlArticulos = "SELECT * FROM articulos"
    
    With rsArticulos
        Call .Open(sqlArticulos, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        .AddNew
    
        .Fields("Codigo").Value = "99999"
        .Fields("Codigo_Num").Value = 99999
        .Fields("Descrip").Value = "Mov. por agrupacion"
        .Fields("Rubro").Value = 0
        .Fields("Stock").Value = 0
        .Update
    
    End With
        
    sqlArticulos = ""
    
    rsArticulos.Close
    Set rsArticulos = Nothing

If Err Then GrabarLog "ActualizarRubros", Err.Number & " " & Err.Description, "Procedimientos"
End Sub
Public Sub AcomodarArticulos()
On Error Resume Next

    Dim rsArticulos As New ADODB.Recordset, sqlArticulos As String
    
    sqlArticulos = "SELECT * FROM Articulos"
    
    With rsArticulos
        Call .Open(sqlArticulos, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        Do Until .EOF = True
            .Fields("Codigo").Value = Trim(.Fields("Codigo").Value)
            .Fields("Descrip").Value = Trim(.Fields("Descrip").Value)
            .MoveNext
        Loop
        
    End With
    
    sqlArticulos = ""
    
    rsArticulos.Close
    Set rsArticulos = Nothing
    
    MsgBox Err.Description

If Err Then GrabarLog "AcomodarArticulos", Err.Number & " " & Err.Description, "Procedimientos"
End Sub
Public Sub ModificarStock(vSumaResta, vCantArticulo, vcodigo, Optional vCosto As Double)
On Error Resume Next

    Dim rsStockArticulo As New ADODB.Recordset, sqlStockArticulo As String

    sqlStockArticulo = "SELECT * FROM Articulos WHERE (codigo = '" & vcodigo & "')"

    With rsStockArticulo
        Call .Open(sqlStockArticulo, ConnDDBB, adOpenDynamic, adLockPessimistic)
        
        If Not .EOF = True Then
            If Not IsNull(.Fields("Peso_T").Value) = True And Not .Fields("Peso_T").Value = 0 Then
                .Fields("peso_t").Value = .Fields("peso_t").Value + .Fields("peso_u").Value * vCantArticulo * vSumaResta
                .Fields("Stock").Value = .Fields("Stock").Value + vCantArticulo * vSumaResta
            Else
                .Fields("stock").Value = .Fields("stock").Value + (Val(vCantArticulo) * vSumaResta)
            End If

            If Not Val(vCosto) = 0 Then .Fields("PCosto").Value = vCosto
            
            
            .Update
        End If
        
    End With
    
    sqlStockArticulo = ""

    If rsStockArticulo.State = 1 Then
        rsStockArticulo.Close
        Set rsStockArticulo = Nothing
    End If

If Err Then GrabarLog "ModificarStock", Err.Number & " " & Err.Description, "Global"
End Sub
Public Function GuardarCuentaContable() As Boolean
On Error Resume Next

    Dim rsCuentasGestion As New ADODB.Recordset, sqlCuentasGestion As String

    sqlCuentasGestion = "SELECT * FROM Articulos WHERE (codigo = '" & 1 & "')"

    With rsCuentasGestion
        Call .Open(sqlCuentasGestion, ConnDDBB, adOpenDynamic, adLockPessimistic)
        'CodigoCuenta

        GuardarCuentaContable = False
    End With
    
    sqlCuentasGestion = ""

    If rsCuentasGestion.State = 1 Then
        rsCuentasGestion.Close
        Set rsCuentasGestion = Nothing
    End If

If Err Then GrabarLog "GuardarCuentaContable", Err.Number & " " & Err.Description, "Global"
End Function
Public Function MostrarCodigoCuenta(vCodigoCuenta As String) As String
On Error Resume Next


MostrarCodigoCuenta = Replace(vCodigoCuenta, "0", "")

Exit Function

    Dim vPartes() As String
    
    ReDim vPartes(4)
    
    vPartes(0) = "0"
    vPartes(1) = "0"
    vPartes(2) = "00"
    vPartes(3) = "00"
    vPartes(4) = "00"
            
    Select Case Len(vCodigoCuenta)
    
        Case 1
            vPartes(0) = Mid(vCodigoCuenta, 1, 1)
            
        Case 2
            vPartes(0) = Mid(vCodigoCuenta, 1, 1)
            vPartes(1) = Mid(vCodigoCuenta, 2, 1)
            
            
        Case 4
            vPartes(0) = Mid(vCodigoCuenta, 1, 1)
            vPartes(1) = Mid(vCodigoCuenta, 2, 1)
            vPartes(2) = Mid(vCodigoCuenta, 3, 2)

        Case 6
            vPartes(0) = Mid(vCodigoCuenta, 1, 1)
            vPartes(1) = Mid(vCodigoCuenta, 2, 1)
            vPartes(2) = Mid(vCodigoCuenta, 3, 2)
            vPartes(3) = Mid(vCodigoCuenta, 5, 2)
            
        Case 8
            vPartes(0) = Mid(vCodigoCuenta, 1, 1)
            vPartes(1) = Mid(vCodigoCuenta, 2, 1)
            vPartes(2) = Mid(vCodigoCuenta, 3, 2)
            vPartes(3) = Mid(vCodigoCuenta, 5, 2)
            vPartes(4) = Mid(vCodigoCuenta, 7, 2)
    
        Case Else
            MostrarCodigoCuenta = ""
            Exit Function
    End Select
    
    MostrarCodigoCuenta = vPartes(0) & "-" & vPartes(1) & "-" & vPartes(2) & "-" & vPartes(3) & "-" & vPartes(4)
    
If Err Then GrabarLog "MostrarCodigoCuenta", Err.Number & " " & Err.Description, "Procedimientos"
End Function
Public Function VerNivelCuenta(vCodigoCuenta As String) As Integer
On Error Resume Next
            
    Select Case Len(vCodigoCuenta)
    
        Case 1
            VerNivelCuenta = 1
            
        Case 2
            VerNivelCuenta = 2
        
        Case 4
            VerNivelCuenta = 3
        
        Case 6
            VerNivelCuenta = 4
            
        Case 8
            VerNivelCuenta = 5
        
        Case Else
            VerNivelCuenta = 0
            
    End Select
    
If Err Then GrabarLog "VerNivelCuenta", Err.Number & " " & Err.Description, "Procedimientos"
End Function
Public Function VerCodigoPadre(vCodigoCuenta As String, vNivelCuenta As Integer) As String
On Error Resume Next
            
    Dim rsPadre As New ADODB.Recordset, sqlPadre As String
    Err.Clear
    Select Case vNivelCuenta
        
        Case 1
            sqlPadre = "SELECT * FROM TempSaldosCuentas WHERE 1=2"
        
        Case 2
            sqlPadre = "SELECT * FROM TempSaldosCuentas WHERE (CodigoCuenta = '" & Mid(vCodigoCuenta, 1, 1) & "')"
        
        Case 3
            sqlPadre = "SELECT * FROM TempSaldosCuentas WHERE (CodigoCuenta = '" & Mid(vCodigoCuenta, 1, 2) & "')"
        
        Case 4
            sqlPadre = "SELECT * FROM TempSaldosCuentas WHERE (CodigoCuenta = '" & Mid(vCodigoCuenta, 1, 4) & "')"
        
        Case 5
            sqlPadre = "SELECT * FROM TempSaldosCuentas WHERE (CodigoCuenta = '" & Mid(vCodigoCuenta, 1, 6) & "')"
    
    End Select
    
    With rsPadre
        .CursorLocation = adUseClient
        
        Call .Open(sqlPadre, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If .State = 1 Then
            If Not .EOF = True Then
                VerCodigoPadre = .Fields("CodigoCuenta").Value
            Else
                VerCodigoPadre = "-1"
            End If
        End If
    
    End With
    
    sqlPadre = ""
    
    If rsPadre.State = 1 Then
        rsPadre.Close
        Set rsPadre = Nothing
    End If
    
If Err Then GrabarLog "VerNivelCuenta", Err.Number & " " & Err.Description, "Procedimientos"
End Function
Public Sub MigrarCodigoCliente()
On Error Resume Next

    Dim rsClientes As New ADODB.Recordset, sqlClientes As String
    Dim vCodigoAnterior As String, vCodigoNuevo As String
    
    sqlClientes = "SELECT * FROM Clientes WHERE (Codigo > '0151') ORDER BY Codigo ASC"
    
    With rsClientes
        Call .Open(sqlClientes, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If Not .EOF = True Then .MoveFirst
        
        vCodigoNuevo = "152"
        
        Do Until .EOF = True
            
            vCodigoAnterior = .Fields("Codigo").Value
            
            vCodigoNuevo = String(4 - Len(vCodigoNuevo), "0") & vCodigoNuevo
            
            .Fields("Codigo").Value = vCodigoNuevo
            
            Call EjecutarScript("UPDATE Factura SET Codigo = '" & vCodigoNuevo & "' WHERE (Codigo = '" & vCodigoAnterior & "')")
            Call EjecutarScript("UPDATE CuentasCorrientes SET Codigo = '" & vCodigoNuevo & "' WHERE (Codigo = '" & vCodigoAnterior & "')")
        
            .MoveNext
            vCodigoNuevo = vCodigoNuevo + 1
        Loop
    
    End With

    sqlClientes = ""

    If rsClientes.State = 1 Then
        rsClientes.Close
        Set rsClientes = Nothing
    End If
    
If Err Then GrabarLog "MigrarCodigoCliente", Err.Number & " " & Err.Description, "Procedimientos"
End Sub

Public Sub GenerarBalance(vmarca As String, ByRef Index As Integer, barra As ProgressBar, vfdesde As Date, vfhasta As Date, ByVal vfbdesde As Date, ByVal vfbhasta As Date, ByVal vcb As String, vnb As Integer, _
Optional vCuentaDesde As String, Optional vCuentaHasta As String, Optional vtipo As String, Optional ByVal viddesde As Long, Optional ByVal vidhasta As Long)
On Error Resume Next
Dim bcuentas As New ADODB.Recordset, sqlCuentas As String
    
    'vlinea = 0
    Call BorrarBase("Temp2", pathDBMySQL)
    
    If IsMissing(vCuentaDesde) = True Or EsNulo(vCuentaDesde) = "" Then
        sqlCuentas = "SELECT * FROM cuentas ORDER BY CodigoCuenta ASC"
    Else
        sqlCuentas = "SELECT * FROM cuentas WHERE ((CodigoCuenta >= '" & vCuentaDesde & "') AND (CodigoCuenta <= '" & vCuentaHasta & "')) ORDER BY CodigoCuenta ASC"
    End If
    
    With bcuentas
        Call .Open(sqlCuentas, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then
            barra.Value = 0
            barra.Max = .RecordCount
            .MoveFirst
        End If
        
        Do Until .EOF = True
            DoEvents
            Call CalcularMovimientos(vmarca, Index, .Fields("CodigoCuenta").Value, .Fields("Cuenta").Value, vfdesde, vfhasta, vfbdesde, vfbhasta, vcb, vnb, vtipo, viddesde, vidhasta)
            barra.Value = barra.Value + 1
            .MoveNext
        Loop
    
    End With

    sqlCuentas = ""
    
    bcuentas.Close
    Set bcuentas = Nothing
    
If Err Then GrabarLog "GenerarBalance", Err.Number & " " & Err.Description, "Procedimientos"
End Sub



'Public Sub GenerarBalanceOld(ByRef Index As Integer, barra As ProgressBar, vfdesde As Date, vfhasta As Date, Optional vCuentaDesde As String, Optional vCuentaHasta As String)
'On Error Resume Next
'
'    Dim bcuentas As New ADODB.Recordset, sqlCuentas As String
'
'    'vlinea = 0
'    Call BorrarBase("Temp2", pathDBMySQL)
'
'    If IsMissing(vCuentaDesde) = True Or EsNulo(vCuentaDesde) = "" Then
'        sqlCuentas = "SELECT * FROM cuentas ORDER BY CodigoCuenta ASC"
'    Else
'        sqlCuentas = "SELECT * FROM cuentas WHERE ((CodigoCuenta >= '" & vCuentaDesde & "') AND (CodigoCuenta <= '" & vCuentaHasta & "')) ORDER BY CodigoCuenta ASC"
'    End If
'
'    With bcuentas
'        Call .Open(sqlCuentas, ConnDDBB, adOpenStatic, adLockReadOnly)
'
'        If Not .EOF = True Then
'            barra.Value = 0
'            barra.Max = .RecordCount
'            .MoveFirst
'        End If
'
'        Do Until .EOF = True
'            DoEvents
'
'            CalcularMovimientos Index, .Fields("CodigoCuenta").Value, .Fields("Cuenta").Value, vfdesde, vfhasta
'
'            barra.Value = barra.Value + 1
'            .MoveNext
'
'        Loop
'
'    End With
'
'    sqlCuentas = ""
'
'    bcuentas.Close
'    Set bcuentas = Nothing
'
'If Err Then GrabarLog "GenerarBalance", Err.Number & " " & Err.Description, "Procedimientos"
'End Sub
'
'Public Sub CalcularMovimientos2(vmarca As String, Index As Integer, vCodCuenta As Long, vCuenta As String, vfdesde As Date, vfhasta As Date, vfDesdeBalance As Date, vfHastaBalance As Date, vbc As String, ByVal vnrobalance As Integer, Optional vTipo As String)
'On Error Resume Next
'
'    Dim vsqlTimeStamp As String
'
'    Dim basientos As New ADODB.Recordset, sqlAsientos, sqlMarca As String
'    Dim vSaldoInicial, vSaldoPeriodo, vSaldoFinal As Double
'   ' Dim vnrobalance As Integer
'    'Dim vfHasta, vfDesde As Date
'    'Dim vcodigoBalance As String
'
'
'
'    'vnrobalance = TraerDato("balances", " Activo='S' order by NroBalance Desc", "NroBalance", pathDBMySQL)
'    'vnrobalance = TraerDato("FechaInicio", " Activo='S' order by NroBalance Desc", "NroBalance", pathDBMySQL)
'    'vnrobalance = TraerDato("FechaFin", " Activo='S' order by NroBalance Desc", "NroBalance", pathDBMySQL)
'
'
'    'vfDesde = traerDatos2("select * from balance where Activo='S' order by idBalance desc", "FechaInicio", pathDBMySQL)
'    'vfHasta = traerDatos2("select * from balance where Activo='S' order by idBalance desc", "FechaFin", pathDBMySQL)
'    'vcodigoBalance = traerDatos2("select * from balance where Activo='S' order by idBalance desc", "codigo", pathDBMySQL)
'
'   ' Me.vcbalance.Text = vbn
'
'
'  '  vsqlTimeStamp = "and (asientos.`TimeStamp`<=asientosdetalle.TIMESTAMP)"
'    'vsqlTimeStamp = ""
'   'If vnrobalance = 15 Then vsqlTimeStamp = "and ((date(asientos.TIMESTAMP)=date(asientosdetalle.TIMESTAMP)) or (asientos.TIMESTAMP<'" + vtimestampAjuste + "')) and not (asientos.nrointerno >= 150193 and asientos.nrointerno <= 150197)"
'
'    'If vbc = "2011-2012" Then vsqlTimeStamp = "and (((asientos.TIMESTAMP)<=(asientosdetalle.TIMESTAMP)) or (asientos.TIMESTAMP<'" + vtimestampAjuste + "')) and not (asientos.nrointerno >= 150192 and asientos.nrointerno <= 150197)"
'
'    '(2012-11-28) If vbc = "2011-2012" Then vsqlTimeStamp = "and (((asientos.TIMESTAMP)<=(asientosdetalle.TIMESTAMP)) or (asientos.TIMESTAMP<'" + vtimestampAjuste + "'))"
'
'    If vbc = "2011-2012" Then vsqlTimeStamp = "and ((asientosdetalle.idasientosdetalle >= 87246) or (asientos.TIMESTAMP<'" + vtimestampAjuste + "'))"
'
'
'
''    If vnrobalance = 15 Then vsqlTimeStamp = "  and not (asientos.nrointerno >= 150193 and asientos.nrointerno <= 150197)"
'
'
'    vTotalD = 0
'    vTotalH = 0
'    vsaldo = 0
'
'
'    If vmarca = "TODOS" Then
'
'    sqlMarca = ""
'
'    Else
'
'            If vmarca = "NORMAL" Then
'                sqlMarca = " and (marca='" + vmarca + "' or marca is null)"
'            Else
'                sqlMarca = " and (marca='" + vmarca + "')"
'            End If
'
'    End If
'
'
'    If vbc = "2011-2012" Then
'    'If vnrobalance = 15 Then
'      'sqlAsientos = "SELECT * FROM Asientos INNER JOIN AsientosDetalle ON Asientos.Numero = AsientosDetalle.Numero WHERE ((fecha >= '" + strfechaMySQL(vfdesde) + "') AND (fecha <= '" + strfechaMySQL(vfhasta) + "')) AND (CodigoCuenta = " & vCodCuenta & ") and (Asientos.NroBalance=" + Str(vnrobalance) + ") " + vsqlTimeStamp + " ORDER BY Asientos.Numero ASC"
'
'
'                sqlAsientos = "SELECT * FROM Asientos INNER JOIN AsientosDetalle ON Asientos.Numero = AsientosDetalle.Numero WHERE ((fecha >= '" + strfechaMySQL(vfdesde) + "') AND (fecha <= '" + strfechaMySQL(vfhasta) + "')) AND (CodigoCuenta = " & vCodCuenta & ") " + vsqlTimeStamp + " ORDER BY Asientos.Numero ASC"
'
'
'    Else
'
'        If frmBalance.chkvarios = 1 Then
'
'                sqlAsientos = "SELECT * FROM Asientos INNER JOIN AsientosDetalle ON Asientos.Numero = AsientosDetalle.Numero WHERE ((fecha >= '" + strfechaMySQL(vfdesde) + "') AND (fecha <= '" + strfechaMySQL(vfhasta) + "')) AND (CodigoCuenta = " & vCodCuenta & ") and (Asientos.NroBalance=AsientosDetalle.NroBalance)" + vsqlTimeStamp + sqlMarca + " ORDER BY Asientos.Numero ASC"
'
'
'            'sqlAsientos = "SELECT * FROM Asientos INNER JOIN AsientosDetalle ON Asientos.Numero = AsientosDetalle.Numero WHERE ((fecha >= '" + strfechaMySQL(vfdesde) + "') AND (fecha <= '" + strfechaMySQL(vfhasta) + "')) AND (CodigoCuenta = " & vCodCuenta & ") and (Asientos.NroBalance=" + Str(vnrobalance) + ") and (Asientos.NroBalance=AsientosDetalle.NroBalance)" + vsqlTimeStamp + sqlMarca + " ORDER BY Asientos.Numero ASC"
'        Else
'            sqlAsientos = "SELECT * FROM Asientos INNER JOIN AsientosDetalle ON Asientos.Numero = AsientosDetalle.Numero WHERE ((fecha >= '" + strfechaMySQL(vfdesde) + "') AND (fecha <= '" + strfechaMySQL(vfhasta) + "')) AND (CodigoCuenta = " & vCodCuenta & ") and (Asientos.NroBalance=" + Str(vnrobalance) + ") and (Asientos.NroBalance=AsientosDetalle.NroBalance)" + vsqlTimeStamp + sqlMarca + " ORDER BY Asientos.Numero ASC"
'            'sqlAsientos = "SELECT * FROM Asientos INNER JOIN AsientosDetalle ON Asientos.Numero = AsientosDetalle.Numero WHERE ((fecha >= '" + strfechaMySQL(vfdesde) + "') AND (fecha <= '" + strfechaMySQL(vfhasta) + "')) AND (CodigoCuenta = " & vCodCuenta & ") and (Asientos.NroBalance=AsientosDetalle.NroBalance)" + vsqlTimeStamp + sqlMarca + " ORDER BY Asientos.Numero ASC"
'        End If
'
'
'    End If
'
'
'    frmBalance.log.AddItem ("------------------------------------------------------------------")
'    frmBalance.log.AddItem (vCodCuenta)
'    frmBalance.log.AddItem ("------------------------------------------------------------------")
'
'
'    'Dim vida, vidad As Long
'
'
'    With basientos
'        Call .Open(sqlAsientos, ConnDDBB, adOpenStatic, adLockReadOnly)
'
'        If Not .RecordCount = 0 Then .MoveFirst
'
'        Do Until .EOF = True
'
'            frmBalance.log.AddItem ("S. Acumulado: " + Str(vsaldo) + "            F: " + Str(.Fields("Fecha").Value) + "            D/H : " + Str(Val(Format(.Fields("Debe").Value, "#######0.000")) - Val(Format(.Fields("Haber").Value, "#######0.000"))))
'
'            'vidad = .Fields("idAsientosDetalle")
'            'vida = .Fields("idAsientos")
'
'           ' Call cambiarNroBalance(vidad, vida)
'
'            vsaldo = vsaldo + Val(Format(.Fields("Debe").Value, "#######0.000")) - Val(Format(.Fields("Haber").Value, "#######0.000"))
'            vTotalD = vTotalD + Val(Format(.Fields("Debe").Value, "#######0.000"))
'            vTotalH = vTotalH + Val(Format(.Fields("Haber").Value, "#######0.000"))
'
'            .MoveNext
'
'
'        Loop
'
'    End With
'
'
'    ' ---------  calculo de los valores indirectos de cada renglón del balance ----------------------
'    vSaldoInicial = CalSaldoAnteriorCtaContable(vmarca, vCodCuenta, vnrobalance, vfdesde, vbc, vfDesdeBalance)
'    vSaldoPeriodo = vTotalD - vTotalH
'    vSaldoFinal = vSaldoInicial + vSaldoPeriodo
'    ' -----------------------------------------------------------------------------------------------
'
'
'    'GuardarTemp Index, vCodCuenta, vCuenta, vfdesde, vfhasta, vSaldoInicial, vSaldoPeriodo, vSaldoFinal
'    GuardarTemp Index, vCodCuenta, vCuenta, vfdesde, vfhasta, vSaldoInicial, vSaldoPeriodo, vSaldoFinal
'
'
'If Err Then GrabarLog "CalcularMovimientos", Err.Number & " " & Err.Description, "Procedimientos"
'End Sub




'Public Sub CalcularMovimientosOld(Index As Integer, vCodCuenta As Long, vCuenta As String, vfdesde As Date, vfhasta As Date)
'On Error Resume Next
'
'    Dim basientos As New ADODB.Recordset, sqlAsientos As String
'    Dim vSaldoInicial, vSaldoPeriodo, vSaldoFinal As Double
'    Dim vnrobalance As Integer
'
'
'    vnrobalance = TraerDato("balances", " Activo='S' order by NroBalance Desc", "NroBalance", pathDBMySQL)
'
'
'
'    vTotalD = 0
'    vTotalH = 0
'    vsaldo = 0
'
'    sqlAsientos = "SELECT * FROM Asientos INNER JOIN AsientosDetalle ON Asientos.Numero = AsientosDetalle.Numero WHERE ((fecha >= '" + strfechaMySQL(vfdesde) + "') AND (fecha <= '" + strfechaMySQL(vfhasta) + "')) AND (CodigoCuenta = " & vCodCuenta & ") and (Asientos.NroBalance=" + Str(vnrobalance) + ") ORDER BY Asientos.Numero ASC"
'
'
'    frmBalance.log.AddItem ("------------------------------------------------------------------")
'    frmBalance.log.AddItem (vCodCuenta)
'    frmBalance.log.AddItem ("------------------------------------------------------------------")
'
'
'    With basientos
'        Call .Open(sqlAsientos, ConnDDBB, adOpenStatic, adLockReadOnly)
'
'        If Not .RecordCount = 0 Then .MoveFirst
'
'        Do Until .EOF = True
'
'            frmBalance.log.AddItem ("S. Acumulado: " + Str(vsaldo) + "            F: " + Str(.Fields("Fecha").Value) + "            D/H : " + Str(Val(Format(.Fields("Debe").Value, "#######0.000")) - Val(Format(.Fields("Haber").Value, "#######0.000"))))
'
'            vsaldo = vsaldo + Val(Format(.Fields("Debe").Value, "#######0.000")) - Val(Format(.Fields("Haber").Value, "#######0.000"))
'            vTotalD = vTotalD + Val(Format(.Fields("Debe").Value, "#######0.000"))
'            vTotalH = vTotalH + Val(Format(.Fields("Haber").Value, "#######0.000"))
'
'            .MoveNext
'        Loop
'
'    End With
'
'
'    ' ---------  calculo de los valores indirectos de cada renglón del balance ----------------------
'    vSaldoInicial = CalSaldoAnteriorCtaContable(vCodCuenta, vnrobalance, vfdesde)
'    vSaldoPeriodo = vTotalD - vTotalH
'    vSaldoFinal = vSaldoInicial + vSaldoPeriodo
'    ' -----------------------------------------------------------------------------------------------
'
'
'    'GuardarTemp Index, vCodCuenta, vCuenta, vfdesde, vfhasta, vSaldoInicial, vSaldoPeriodo, vSaldoFinal
'    GuardarTemp Index, vCodCuenta, vCuenta, vfdesde, vfhasta, vSaldoInicial, vSaldoPeriodo, vSaldoFinal
'
'
'If Err Then GrabarLog "CalcularMovimientos", Err.Number & " " & Err.Description, "Procedimientos"
'End Sub
'
'Public Function CalSaldoPersona(ByVal vcodigo As String, vtabla As String) As Double
'Dim vsql As String
'
'vsql = "SELECT   SUM(Credito) AS c,   sum(Debito) as d, (sum(Debito)-SUM(Credito)) as saldo From " + vtabla + "  where codigo='" + vcodigo + "'"
'CalSaldoPersona = Val(EsNulo(traerDatos2(vsql, "saldo", pathDBMySQL)))
'
'End Function


Public Function CalSaldoAnteriorCtaContable(ByVal vmarca As String, ByVal vCodCuenta As String, vnrobalance As Integer, ByVal vfechaHasta As Date, ByVal vbc As String, ByVal vfbdesde As Date, Optional vtipo As String, Optional ByVal viddesde As Long, Optional ByVal vidhasta As Long) As Double
On Error Resume Next

'vmensaje.Caption = "Calculando saldo de la cuenta: " + vCodCuenta

Dim vsqlanterior, vwhereAnterior, vwhereActual, vnrointernos, vCorrector As String
Dim rsSaldo As New ADODB.Recordset, sqlSaldo As String
Dim vsqlFechBalance As String
Dim sqlMarca As String
Dim sqlFecha As String


sqlFecha = ""

If viddesde > 0 And vidhasta Then
   ' sqlFecha = "idAsientos >" + Str(viddesde) + " and idAsientos <= " + Str(vidhasta)
    sqlFecha = "idAsientos <=" + Str(viddesde)
Else
    vsqlFechBalance = "(asientos.Fecha >= '" & strfechaMySQL(vfbdesde) & "') and  (asientos.Fecha < '" & strfechaMySQL(vfechaHasta) & "') "
End If


'vnrointernos = " and not (asientos.nrointerno>=150193 and asientos.nrointerno<=150197)"

' condicionales para el saldo anterior
vCorrector = "" ' corrector para el balance 15


'If vnrobalance = 15 Then
If vbc = "2011-2012" Then
   ' vCorrector = " (asientos.NroBalance =" & vnrobalance & ")"
   ' vCorrector = vCorrector + " and ((date(asientos.TIMESTAMP)=date(asientosdetalle.TIMESTAMP)) or (asientos.TIMESTAMP<'" + vtimestampAjuste + "')) "
   ' corregido 21:27
   ' vCorrector = " (asientos.NroBalance =" & vnrobalance & ") and (((asientos.TIMESTAMP)<=(asientosdetalle.TIMESTAMP)) or (asientos.TIMESTAMP<'" + vtimestampAjuste + "')) and not (asientos.nrointerno >= 150192 and asientos.nrointerno <= 150197)"
     
    vCorrector = " (asientos.NroBalance =15) and ((asientosdetalle.idasientosdetalle >= 87246) or (asientos.TIMESTAMP<'" + vtimestampAjuste + "'))"
     
     
     ' (28-11-12) vCorrector = " (asientos.NroBalance =15) and (((asientos.TIMESTAMP)<=(asientosdetalle.TIMESTAMP)) or (asientos.TIMESTAMP<'" + vtimestampAjuste + "'))"
    ' vCorrector = "(((asientos.TIMESTAMP)<=(asientosdetalle.TIMESTAMP)) or (asientos.TIMESTAMP<'" + vtimestampAjuste + "'))"
    ' vCorrector = "asientosdetalle.`idAsientosDetalle` >= 87246"
    
    vsqlFechBalance = " (asientos.Fecha < '" & strfechaMySQL(vfechaHasta) & "') "
Else
    vCorrector = " (asientos.NroBalance =" & vnrobalance & ") and (asientos.NroBalance =asientosdetalle.NroBalance) "
    ' la fecha de los saldos del balance commienza en la fecha de inicio del balance
    vsqlFechBalance = "(asientos.Fecha >= '" & strfechaMySQL(vfbdesde) & "') and  (asientos.Fecha < '" & strfechaMySQL(vfechaHasta) & "') "
End If
    

    If vmarca = "TODOS" Then
    
    sqlMarca = ""
    
    Else
    
            If vmarca = "and (marca = 'NORMAL')" Then
                sqlMarca = " and (marca='" + vmarca + "' or marca is null)"
            Else
                sqlMarca = " and (marca='" + vmarca + "')"
            End If
   
    End If


If Not sqlFecha = "" Then ' pararlo
        vsqlFechBalance = sqlFecha
End If


vwhereAnterior = "(asientosdetalle.CodigoCuenta ='" + vCodCuenta & "') AND " & _
                vsqlFechBalance & " and " & _
                " " + vCorrector

  '' esta fecha or (asientos.TIMESTAMP<'2012-08-10'))  es porque dejo que los movimientos  con fechas anteriores a esta se calculen el saldo con movimientos superpuestos

  sqlSaldo = "SELECT  sum(asientosdetalle.Debe) as d, sum(asientosdetalle.Haber) As H, (sum(asientosdetalle.Debe) - sum(asientosdetalle.Haber)) as Saldo " & _
                 " From  asientos " & _
                 " INNER JOIN asientosdetalle ON (asientos.Numero=asientosdetalle.Numero) " & _
                 " INNER JOIN cuentas ON (asientosdetalle.CodigoCuenta=cuentas.CodigoCuenta) " & _
                 " Where " & vwhereAnterior & sqlMarca
    
    
     sqlSaldo = "SELECT  sum(asientosdetalle.Debe) as d, sum(asientosdetalle.Haber) As H, (sum(asientosdetalle.Debe) - sum(asientosdetalle.Haber)) as Saldo " & _
                 " From  asientos " & _
                 " INNER JOIN asientosdetalle ON (asientos.Numero=asientosdetalle.Numero) " & _
                 " " & _
                 " Where " & vwhereAnterior & sqlMarca
    
       
        'sqlSaldo = "SELECT (Temp2.C10) AS Codigo, Sum(Temp2.C03) AS SumaDeDebe, Sum(Temp2.C04) AS SumaDeHaber, Sum(C03-C04) AS Saldo FROM Temp2 GROUP BY Temp2.C10 HAVING (((Temp2.C10) = '" & vCodCuenta & "'))"
    'End If

    With rsSaldo
        Call .Open(sqlSaldo, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then
            CalSaldoAnteriorCtaContable = Val(Format(.Fields("Saldo").Value, "######0.00"))
        Else
           ' MsgBox "No se registraron movimientos para ser procesado.", vbCritical, "Error ..."
            CalSaldoAnteriorCtaContable = 0
        End If
    
    End With
        
    sqlSaldo = ""
    
    If rsSaldo.State = 1 Then
        rsSaldo.Close
        Set rsSaldo = Nothing
    End If
    
If Err Then GrabarLog "CalcularSaldoCtaContable", Err.Number & " " & Err.Description, "función gral"
End Function



'Public Function CalSaldoAnteriorCtaContableOld(ByVal vCodCuenta As String, vnrobalance As Integer, vfechaHasta As Date) As Double
'On Error Resume Next
'
''vmensaje.Caption = "Calculando saldo de la cuenta: " + vCodCuenta
'
'Dim vsqlanterior, vwhereAnterior, vwhereActual As String
'Dim rsSaldo As New ADODB.Recordset, sqlSaldo As String
'
'
'' condicionales para el saldo anterior
'vwhereAnterior = "(asientosdetalle.CodigoCuenta ='" + vCodCuenta & "' AND " & _
'                " (asientos.Fecha < '" & strfechaMySQL(vfechaHasta) & "') and " & _
'                 " asientos.NroBalance =" & vnrobalance & " )"
'
'
'  'vmensaje.Caption = "Calculando saldo anterior de la cuenta: " + vCodCuenta
'
'
'
'  sqlSaldo = "SELECT  sum(asientosdetalle.Debe) as d, sum(asientosdetalle.Haber) As H, (sum(asientosdetalle.Debe) - sum(asientosdetalle.Haber)) as Saldo " & _
'                 " From  asientos " & _
'                 " INNER JOIN asientosdetalle ON (asientos.Numero=asientosdetalle.Numero) " & _
'                 " INNER JOIN cuentas ON (asientosdetalle.CodigoCuenta=cuentas.CodigoCuenta) " & _
'                 " Where " & vwhereAnterior
'
'
'        'sqlSaldo = "SELECT (Temp2.C10) AS Codigo, Sum(Temp2.C03) AS SumaDeDebe, Sum(Temp2.C04) AS SumaDeHaber, Sum(C03-C04) AS Saldo FROM Temp2 GROUP BY Temp2.C10 HAVING (((Temp2.C10) = '" & vCodCuenta & "'))"
'    'End If
'
'    With rsSaldo
'        Call .Open(sqlSaldo, ConnDDBB, adOpenStatic, adLockReadOnly)
'
'        If Not .EOF = True Then
'            CalSaldoAnteriorCtaContable = Val(Format(.Fields("Saldo").Value, "######0.00"))
'        Else
'           ' MsgBox "No se registraron movimientos para ser procesado.", vbCritical, "Error ..."
'            CalSaldoAnteriorCtaContable = 0
'        End If
'
'    End With
'
'    sqlSaldo = ""
'
'    If rsSaldo.State = 1 Then
'        rsSaldo.Close
'        Set rsSaldo = Nothing
'    End If
'
'If Err Then GrabarLog "CalcularSaldoCtaContable", Err.Number & " " & Err.Description, "función gral"
'End Function

Public Function CalSaldoPersona(ByVal vcodigo As String, vtabla As String) As Double
Dim vsql As String

vsql = "SELECT   SUM(Credito) AS c,   sum(Debito) as d, (sum(Debito)-SUM(Credito)) as saldo From " + vtabla + "  where codigo='" + vcodigo + "'"
CalSaldoPersona = Val(EsNulo(traerDatos2(vsql, "saldo", pathDBMySQL)))

End Function
Public Sub GuardarTemp(Index, vCodCuenta, vcuenta, vfhasta, vfdesde, vSaldoInicial, vSaldoPeriodo, vSaldoFinal, vpresupuestado)
On Error Resume Next
Dim vimputable, vsql, vnivel, vespacio  As String
Dim i As Integer


    Dim btemp As New ADODB.Recordset, sqlGuardar As String
    
    If Index = 0 Then
        'Asiento de Apertura o Cierre
        sqlGuardar = "SELECT * FROM Asientos ORDER BY numero ASC, id_asientos ASC"
                                       
        With btemp
            Call .Open(sqlGuardar, ConnDDBB, adOpenDynamic, adLockOptimistic)
            
            If Not vsaldo = 0 Then
                'vlinea = vlinea + 1
                
                .AddNew
                '.Fields("Numero").Value = vNumero
                .Fields("Codigo").Value = vCodCuenta
                
                If vsaldo > 0 Then
                    .Fields("Debe").Value = 0
                    .Fields("Haber").Value = Val(Format(vsaldo, "######0.00"))
                Else
                    .Fields("Debe").Value = Val(Format(vsaldo, "######0.00")) * (-1)
                    .Fields("Haber").Value = 0
                End If
                
                .Fields("Fecha").Value = vfdesde
                .Fields("NCuenta").Value = vcuenta
                '.Fields("Leyenda").Value = vLeyenda
                '.Fields("Linea").Value = vlinea
                         
                .Update
            End If
        End With
        
    Else
    
    
    
    vsql = "select  t.imputable as c   from cuentas t where t.codigocuenta = '" + vCodCuenta + "'"
    vimputable = traerDatos2(vsql, "c", pathDBMySQL)
    
    
    
    vsql = "select  niveles as c   from cuentas t where t.codigocuenta = '" + vCodCuenta + "'"
    vnivel = traerDatos2(vsql, "c", pathDBMySQL)
    
    ' ----- guarda en el temporal los renglosnes del balance -------
    
        sqlGuardar = "SELECT * FROM Temp2"
 
        With btemp
            Call .Open(sqlGuardar, ConnDDBB, adOpenDynamic, adLockOptimistic)
            
            .AddNew
            
            .Fields("C09").Value = Trim(vimputable)
            
            Debug.Print "-------" + vimputable
            
            .Fields("C02").Value = vCodCuenta
            
            '.Fields("C10").Value = BuscarRubros(0, Str(vCodCuenta))
            '.Fields("C12").Value = BuscarRubros(1, Str(vCodCuenta))
            
            
            
            If Trim(vimputable) = "N" Then vcuenta = "+[" + UCase(vcuenta) + "]"
            
            
                For i = 1 To Len(vnivel)
                    vespacio = vespacio + "   "
                Next
                
                
            If Trim(vimputable) = "S" Then
                vcuenta = vespacio + "-   " + vcuenta
            Else
                vcuenta = vespacio + vcuenta
            End If
            
            .Fields("C05").Value = vcuenta
            
            
            
            .Fields("C03").Value = vTotalD
            .Fields("C04").Value = vTotalH
            
            
            .Fields("C06").Value = vnivel
            
            
        
        
' ------------------------------------  balance nuevo  -------------------



.Fields("C07").Value = vSaldoInicial


.Fields("C11").Value = vSaldoPeriodo


.Fields("C13").Value = vSaldoFinal ' acumulado

    
   

.Fields("C14").Value = vpresupuestado ' presupuestado


.Fields("C15").Value = Abs(vSaldoFinal) - Abs(vpresupuestado)    ' diferencia presupuesto





'--------------------------------------------------------------------------


            
' -----------------------------------  balance viejo ---------------------
        
'            Select Case Val(vsaldo)
'
'                Case Is > 0
'                    .Fields("C07").Value = Val(Format(vsaldo, "#######0.000"))
'                    .Fields("C11").Value = 0
'
'                Case Is < 0
'                    .Fields("C07").Value = 0
'                    .Fields("C11").Value = Val(Format(vsaldo, "#######0.000")) * (-1)
'
'                Case 0
'                    .Fields("C07").Value = 0
'                    .Fields("C11").Value = 0
'            End Select
'----------------------------------------------------------------------------
            .Update

        End With
    
    End If
    
    sqlGuardar = ""

    btemp.Close
    Set btemp = Nothing
    
If Err Then GrabarLog "GuardarTemp", Err.Number & " " & Err.Description, "Procedimientos"
End Sub
Private Sub MigrarLocalidades()
    On Error Resume Next

    Dim connDDBBLocalidad As New ADODB.Connection, rsLocalidades As New ADODB.Recordset, sqlLocalidades As String
    
    With connDDBBLocalidad
        .ConnectionString = pathDBLocalidad
        .Open
        
        If Not .State = 1 Then
            MsgBox Err.Description
            Exit Sub
        End If
    End With
    
    Call BorrarBase("Localidades", pathDBMySQL)
    sqlLocalidades = "SELECT * FROM LocalidadesPorProvincia"
        
    With rsLocalidades
        Call .Open(sqlLocalidades, connDDBBLocalidad, adOpenStatic, adLockPessimistic)
        
        If Not .State = 0 Then
            
            Do Until .EOF = True
                DoEvents
                If Not IsNull(.Fields(0).Value) = True Then
                    Call EjecutarScript("INSERT INTO Localidades (CodigoPostal, Localidad, Provincia) VALUES ('" & .Fields(0).Value & "','" & .Fields(1).Value & "','" & .Fields(2).Value & "') ")
                End If
                .MoveNext
            Loop
        
            
        End If
        
    End With

    sqlLocalidades = ""
    
    If rsLocalidades.State = 1 Then
        rsLocalidades.Close
        Set rsLocalidades = Nothing
    End If
    
    If connDDBBLocalidad.State = 1 Then
        connDDBBLocalidad.Close
        Set connDDBBLocalidad = Nothing
    End If
    
    If Err Then GrabarLog "MigrarLocalidades", Err.Number & " " & Err.Description, "Procedimientos"
End Sub
Function EsBancoOCaja(vCodigoCuenta As String) As String
On Error Resume Next

    Dim rsBancos As New ADODB.Recordset, sqlBancos As String
    
    sqlBancos = "SELECT * FROM Bancos WHERE (CuentaContableAsociada = '" & vCodigoCuenta & "')"
    
    With rsBancos
        .CursorLocation = adUseClient
        Call .Open(sqlBancos, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If (.State = 0) Or (.EOF = True) Then
            EsBancoOCaja = "N"
        Else
            If .Fields("EsCaja").Value = "S" Then
                EsBancoOCaja = "C"
            Else
                EsBancoOCaja = "B"
            End If
        End If
    
    End With
    
    sqlBancos = ""
    
    If rsBancos.State = 1 Then
        rsBancos.Close
        Set rsBancos = Nothing
    End If
    
If Err Then GrabarLog "EsBancoOCaja", Err.Number & " " & Err.Description, "Procedimientos"
End Function
Public Sub GuardarBancosMovimientos2(idBancos As String, idBancosCuentas As Integer, debito As Double, credito As Double, comentario As String, idCuponTarjeta As Integer, NroInterno As Long, Optional vfechaDeposito As Date, Optional vnrocheque, Optional vtipomovimiento, Optional vtipovalor)
    On Error Resume Next
    
    Dim rsBancoMovimientos As New ADODB.Recordset, sqlBancoMovimientos As String
    
    sqlBancoMovimientos = "SELECT * FROM BancosMovimientos WHERE 1=2"
    
    With rsBancoMovimientos
        If .State = 1 Then .Close
       .CursorLocation = adUseClient
       
        Call .Open(sqlBancoMovimientos, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            .AddNew
            .Fields("idBancos").Value = idBancos
            .Fields("idBancosCuentas").Value = idBancosCuentas
            .Fields("debito").Value = Val(debito)
            .Fields("credito").Value = Val(credito)
            .Fields("fecha").Value = CDate(vfechaDeposito)
            .Fields("comentario").Value = Left(comentario, 255)
            .Fields("NroCheque") = Val(EsNulo(vnrocheque))
            .Fields("idCuponTarjeta") = idCuponTarjeta
                        
            .Fields("NroInterno").Value = NroInterno
            .Fields("NroAsiento").Value = Val(GenerarDato("SELECT MAX(Numero) as NroAsiento FROM Asientos;", "NroAsiento")) + 1 ' pone el nro de asiento que luego se va a crear
            .Fields("TipoMovimiento").Value = EsNulo(vtipomovimiento)
            .Fields("TipoMovimiento").Value = EsNulo(vtipovalor)
               
            .Update
            
        End If
        
    End With
    
    sqlBancoMovimientos = ""
    
    If rsBancoMovimientos.State = 1 Then
        rsBancoMovimientos.Close
        Set rsBancoMovimientos = Nothing
   End If
   
If Err Then GrabarLog "GuardarBancosMovimientos", Left(Err.Number & " " & Err.Description, 99), "Procedimientos"
End Sub

Public Sub GuardarBancosMovimientos(vnrorecibo As Long, idBancos As String, idBancosCuentas As Integer, debito As Double, credito As Double, comentario As String, idCuponTarjeta As Integer, NroInterno As Long, Optional vfechaDeposito As Date, Optional vnrocheque As String, Optional vtipomovimiento As String, Optional vtipovalor As String, Optional vIdCheques As Long, Optional vClienteProveedor As String)
On Error Resume Next
Dim vnrobalance As Long

vnrobalance = TraerDato("balances", " Activo='S' order by NroBalance Desc", "NroBalance", pathDBMySQL)
'Me.Caption = Me.Caption + "      [Nro. de Balance: " + Str(vnrobalance) + "]"
    
    Dim rsBancoMovimientos As New ADODB.Recordset, sqlBancoMovimientos As String
    
    sqlBancoMovimientos = "SELECT * FROM BancosMovimientos WHERE 1=2"
    
    With rsBancoMovimientos
        If .State = 1 Then .Close
       .CursorLocation = adUseClient
       
        Call .Open(sqlBancoMovimientos, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            .AddNew
            .Fields("nrocomprobante").Value = vnrorecibo
            .Fields("idBancos").Value = idBancos
            .Fields("idBancosCuentas").Value = idBancosCuentas
            .Fields("debito").Value = Val(debito)
            .Fields("credito").Value = Val(credito)
            .Fields("fecha").Value = CDate(vfechaDeposito)
            .Fields("comentario").Value = Left(comentario, 255)
            .Fields("NroCheque") = Val(EsNulo(vnrocheque))
            .Fields("idCuponTarjeta") = idCuponTarjeta
                        
            .Fields("NroInterno").Value = NroInterno
            .Fields("NroAsiento").Value = Val(GenerarDato("SELECT MAX(Numero) as NroAsiento FROM Asientos where nrobalance=" + Str(vnrobalance) + ";", "NroAsiento")) + 1 ' pone el nro de asiento que luego se va a crear
            .Fields("TipoMovimiento").Value = EsNulo(vtipomovimiento)
            '.Fields("TipoMovimiento").Value = EsNulo(vtipoValor)
            .Fields("idCheques").Value = vIdCheques
            '.Fields("idCheques").Value = vidCheques
            .Fields("ClienteProveedor") = vClienteProveedor
            
            
            .Update
            
            'Call alerta("Graba movi en Caja " + idBancos, 5000)
        End If
        
    End With
    
    sqlBancoMovimientos = ""
    
    If rsBancoMovimientos.State = 1 Then
        rsBancoMovimientos.Close
        Set rsBancoMovimientos = Nothing
   End If
   
If Err Then GrabarLog "GuardarBancosMovimientos", Left(Err.Number & " " & Err.Description, 99), "Procedimientos"
End Sub
Public Function GuardarEnStock(vViene As String, vCodigoArticulo As String, vfecha As Date, vCantidadStock As Double, vcomentario As String, vIDFDetalle As Long, vIDPFDetalle As Long) As Double
On Error Resume Next
    
    Dim rsStock As New ADODB.Recordset, sqlStock As String
    
    Dim vSaldoStockInicial, vdiferencia, vsaldo  As Double
    
    Dim vsql As String
    
    
    
    vsql = "select * from articulos where codigo = '" + Trim(vCodigoArticulo) + "'"
    vSaldoStockInicial = traerDatos2(vsql, "stock", pathDBMySQL)
    
    
    
    Select Case vViene
    
        Case "Articulo-Nuevo" Or "Devolucion" Or "Automatico"
            sqlStock = "SELECT * FROM Stock WHERE 1=2"
        
        Case "Articulo-Modificar"
           ' sqlStock = "SELECT * FROM Stock WHERE (CodigoArticulo = '" & vCodigoArticulo & "') AND (idFDetalle = " & vIDFDetalle & ") AND (idPFDetalle = " & vIDPFDetalle & ") ORDER BY 1"
            
           ' vSaldoStockInicial = Val(Format(GenerarDato("SELECT Sum(Entrada), Sum(Salida), Sum(Entrada-Salida) AS SaldoActual FROM Stock WHERE not(idPFDetalle=" + Str(vIDPFDetalle) + ") and CodigoArticulo = '" & vCodigoArticulo & "'", "SaldoActual"), "#####0.00"))
           ' vdiferencia = vSaldoStockInicial - vCantidadStock
           ' vCantidadStock = vdiferencia
           
          ' Call actualizastockEnArticulo(EsNulo(vCodigoArticulo), vSaldoStockInicial)
            
            Exit Function ' si está modificando el articulo no tiene que modificar stock
            
        Case "Remito-Nuevo"
            sqlStock = "SELECT * FROM Stock WHERE (idFDetalle = " & vIDFDetalle & ") ORDER BY 1"
            
        Case "Remito-Modificar"
            sqlStock = "SELECT * FROM Stock WHERE (idFDetalle = " & vIDFDetalle & ") ORDER BY 1"
            
        Case "Compras-Nuevo"
            sqlStock = "SELECT * FROM Stock WHERE (idPFDetalle = " & vIDPFDetalle & ") ORDER BY 1"
            
        Case "Compras-Modificar"
            sqlStock = "SELECT * FROM Stock WHERE (idPFDetalle = " & vIDPFDetalle & ") ORDER BY 1"
        
        End Select
    
    
    
    '------------------- con lo siguiente anulo lo anterior ----------------
    
    With rsStock
        Call .Open(sqlStock, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If .State = 1 Then If .EOF = True Then .AddNew
        
        .Fields("Fecha").Value = vfecha
        .Fields("CodigoArticulo").Value = EsNulo(vCodigoArticulo)
        
        If vIDFDetalle = 0 Or vViene = "Devolucion" Then  ' estoy comprando
            .Fields("Entrada").Value = Val(vCantidadStock)
           ' .Fields("Saldo").Value = Val(Format(GenerarDato("SELECT Sum(Entrada), Sum(Salida), Sum(Entrada-Salida) AS SaldoActual FROM Stock WHERE not(idPFDetalle=" + Str(vIDPFDetalle) + ") and CodigoArticulo = '" & vCodigoArticulo & "'", "SaldoActual"), "#####0.00")) + vCantidadStock
           ' .Fields("Saldo").Value = Str(vSaldoStockInicial + vCantidadStock)
        
            vsaldo = vSaldoStockInicial + vCantidadStock
            
            .Fields("Saldo").Value = vsaldo
            
            
        Else ' estoy vendiendo
            .Fields("Salida").Value = Val(vCantidadStock)
            '.Fields("Saldo").Value = Val(Format(GenerarDato("SELECT Sum(Entrada), Sum(Salida), Sum(Entrada-Salida) AS SaldoActual FROM Stock WHERE not(idFDetalle=" + Str(vIDFDetalle) + ") and CodigoArticulo = '" & vCodigoArticulo & "'", "SaldoActual"), "#####0.00")) - vCantidadStock
            ' Fields("Saldo").Value = Str(vSaldoStockInicial - vCantidadStock)
            
            vsaldo = vSaldoStockInicial - vCantidadStock
            
            .Fields("Saldo").Value = vsaldo
            
        
            Debug.Print "------------------------------------------------------"
            
            If vCodigoArticulo = "526" Or vCodigoArticulo = "213" Or vCodigoArticulo = "224" Then
        
            Debug.Print "- Codigo: " + EsNulo(vCodigoArticulo)
            Debug.Print "- Stock : " + Str(vSaldoStockInicial)
            Debug.Print "- Venta : " + Str(vCantidadStock)
            Debug.Print "- Saldo : " + Str(.Fields("Saldo").Value)
            
            Debug.Print "------------------------------------------------------"
        
            End If
            
        End If
        
        .Fields("Comentario").Value = Left(vcomentario, 255)
        .Fields("idFDetalle").Value = Val(vIDFDetalle)
        .Fields("idPFDetalle").Value = Val(vIDPFDetalle)

        .Update
    
        GuardarEnStock = Val(Format(.Fields("Saldo").Value, "######0.00"))
    
        Call actualizastockEnArticulo(EsNulo(vCodigoArticulo), vsaldo)
        
        
    
    
    End With

    sqlStock = ""

   

    If rsStock.State = 1 Then
        rsStock.Close
        Set rsStock = Nothing
    End If
    
If Err Then
    GrabarLog "GuardarEnStock", Err.Number & " " & Err.Description, "Procedimientos"
End If
End Function



Public Sub actualizarPreciosArticulo(ByVal vidArticulo As Long, ByVal vpcostoNuevo As Double)
On Error Resume Next

Dim pcosto, pv1, pv2, pv3, pv4, pv5 As Double
Dim p1, p2, p3, p4, p5 As Double
Dim vsql, vcampos, vvalores, vcodigo, vdescripcion   As String

p1 = 0
p2 = 0
p3 = 0
p4 = 0
p5 = 0


pcosto = traerDatos2("select pcosto from articulos where idArticulos=" + Str(vidArticulo), "pcosto", pathDBMySQL)
pv1 = traerDatos2("select Pventa1 from articulos where idArticulos=" + Str(vidArticulo), "Pventa1", pathDBMySQL)
pv2 = traerDatos2("select Pventa2 from articulos where idArticulos=" + Str(vidArticulo), "Pventa2", pathDBMySQL)
pv3 = traerDatos2("select Pventa3 from articulos where idArticulos=" + Str(vidArticulo), "Pventa3", pathDBMySQL)
pv4 = traerDatos2("select Pventa4 from articulos where idArticulos=" + Str(vidArticulo), "Pventa4", pathDBMySQL)
pv5 = traerDatos2("select Pventa5 from articulos where idArticulos=" + Str(vidArticulo), "Pventa5", pathDBMySQL)


vcodigo = traerDatos2("select Codigo from articulos where idArticulos=" + Str(vidArticulo), "Codigo", pathDBMySQL)
vdescripcion = traerDatos2("select Descrip from articulos where idArticulos=" + Str(vidArticulo), "Descrip", pathDBMySQL)


'---- calculo los % de cada uno de lo precios
'If pv1 > 0 Then p1 = pcosto / pv1
'If pv1 > 0 Then p2 = pcosto / pv2
'If pv1 > 0 Then p3 = pcosto / pv3
'If pv1 > 0 Then p4 = pcosto / pv4
'If pv1 > 0 Then p5 = pcosto / pv5



If pv1 > 0 Then p1 = ((pv1 / pcosto) - 1)
If pv2 > 0 Then p2 = ((pv2 / pcosto) - 1)
If pv3 > 0 Then p3 = ((pv3 / pcosto) - 1)
If pv4 > 0 Then p4 = ((pv4 / pcosto) - 1)
If pv5 > 0 Then p5 = ((pv5 / pcosto) - 1)



' calculo los nuevos precios de ventas
pv1 = (1 + p1) * vpcostoNuevo
pv2 = (1 + p2) * vpcostoNuevo
pv3 = (1 + p3) * vpcostoNuevo
pv4 = (1 + p4) * vpcostoNuevo
pv5 = (1 + p5) * vpcostoNuevo

' ---------------------------------
vcampos = "pcosto=" + Str(vpcostoNuevo) + ",Pventa1=" + Str(pv1) + ",Pventa2=" + Str(pv2) + ",Pventa3=" + Str(pv3) + ",Pventa4=" + Str(pv4) + ",Pventa5=" + Str(pv5)
'vvalores = Str(vpcostoNuevo) + "," + Str(pv1) + "," + Str(pv2) + "," + Str(pv3) + "," + Str(pv4) + "," + Str(pv5)

vsql = "update  articulos  set " + vcampos + " where idArticulos=" + Str(vidArticulo)
Call EjecutarScript(vsql, pathDBMySQL)
'----------------------------------------------

vidArticulo = 0

If Err Then
    MsgBox "Debe verificar el estado del artículo desde el módulo de mantenimiento", vbCritical
    Exit Sub
End If

End Sub



Public Sub actualizastockEnArticulo(ByVal vcodigo As String, ByVal vstock As Double)
Dim vsql As String
vsql = "update articulos set stock=" + Str(vstock) + " where codigo='" + vcodigo + "'"
Call EjecutarScript(vsql, pathDBMySQL)
End Sub
Public Function FormatoUltimoCodigo(vEspacios As Integer, vValor) As String
On Error Resume Next

    FormatoUltimoCodigo = String(vEspacios - Len(vValor), "0") & Val(vValor)
    
If Err Then GrabarLog "FormatoUltimoCodigo", Err.Number & " " & Err.Description, "Procedimientos"
End Function
Public Function ControlarEjecuciones(vFechaInicio As Date, vtipo As String) As Date
On Error Resume Next

    'Panic: Controlar Sabados y Domingos y Usar la Tabla Feriados
    
    vtipo = TraerDato("Intervalos", "idIntervalos = " & Val(vtipo) & "", "Intervalo")
    
    Select Case vtipo
    
        Case "Una vez"
            ControlarEjecuciones = vFechaInicio
        
        Case "Diario"
            ControlarEjecuciones = vFechaInicio + 1
            
        Case "Semanal"
            ControlarEjecuciones = vFechaInicio + 7
        
        Case "Mensual"
            ControlarEjecuciones = "01/" & MesSiguiente(vFechaInicio, 1)
        
        Case "Bimestral"
            ControlarEjecuciones = "01/" & MesSiguiente(vFechaInicio, 2)
        
        Case "Trimestral"
            ControlarEjecuciones = "01/" & MesSiguiente(vFechaInicio, 3)
        
        Case "Cuatrimestral"
            ControlarEjecuciones = "01/" & MesSiguiente(vFechaInicio, 4)
        
        Case "Semestral"
            ControlarEjecuciones = "01/" & MesSiguiente(vFechaInicio, 6)
            
        Case "Anual"
            ControlarEjecuciones = "01/" & MesSiguiente(vFechaInicio, 12)
         
    End Select

If Err Then GrabarLog "ControlarEjecuciones", Err.Number & " " & Err.Description, "Procedimientos"
End Function
Public Function MesSiguiente(vfecha As Date, vCantidadDeMeses As Integer) As String
On Error Resume Next

    Dim vMesActual As String, vMesSiguiente As String, vMesFinal As String
    
    vMesActual = Month(vfecha)
    
    vMesSiguiente = vMesActual + vCantidadDeMeses
    
    If Val(vMesSiguiente) > 12 Then
        vMesSiguiente = "0" & (vMesSiguiente - 12) & "/" & Year(vfecha) + 1
    Else
        If vMesSiguiente < 9 Then
            vMesSiguiente = "0" & vMesSiguiente & "/" & Year(vfecha)
        Else
            vMesSiguiente = vMesSiguiente & "/" & Year(vfecha)
        End If
    End If
    
    MesSiguiente = vMesSiguiente

If Err Then GrabarLog "MesSiguiente", Err.Number & " " & Err.Description, "Procedimientos"
End Function
Public Function FacturaAutomatica(vIDFacturaAutomatica As Integer, vCodigoCliente As String, vfecha As Date, vremito As Long) As Long
On Error Resume Next

    Dim rsFacturaAutomaticaDetalle As New ADODB.Recordset, sqlFacturaAutomaticaDetalle As String
    
    sqlFacturaAutomaticaDetalle = "SELECT * FROM FacturaAutomaticaDetalle WHERE (idFacturaAutomatica = " & vIDFacturaAutomatica & ")"

    'Saco el Total
    'Y hago 2 Insert en FDetalle y Factura

    
If Err Then GrabarLog "FacturaAutomatica", Err.Number & " " & Err.Description, "Procedimientos"
End Function


Public Sub BorrarFDetalle(vnrointerno As Long, vid As String)

Dim vremito As Long
Dim vsql, vCodigoArticulo, vtabla, vsqlBorrar As String
Dim t As New ADODB.Recordset
Dim GuardarEnStock As Double

If vid = "idPFDetalle" Then
        vtabla = "PFactura"
        
Else
        vtabla = "Factura"
End If


vremito = Val(EsNulo(traerDatos2("select * from " + vtabla + " where nrointerno=" + Str(vnrointerno), "remito", pathDBMySQL)))


If vid = "idPFDetalle" Then
        vsql = "select " + vid + ", codigo  from pfdetalle where remito=" + Str(vremito)
        vsqlBorrar = "delete from pfdetalle where remito=" + Str(vremito)
Else
        vsql = "select  " + vid + ", codigo from fdetalle where remito=" + Str(vremito)
        vsqlBorrar = "delete from fdetalle where remito=" + Str(vremito)
        
End If


'vremito = Val(EsNulo(traerDatos2("select * from " + vtabla + " where nrointerno=" + Str(vnrointerno), "remito", pathDBMySQL)))


    With t
        Call .Open(vsql, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        
        
        If .EOF = True Then Exit Sub
        
        Do While Not .EOF
            
            If vid = "idPFDetalle" Then
                borrarLineaStockP (.Fields(0))  ' según el idpfdetalle,idfdetalle
            Else
                borrarLineaStockC (.Fields(0))  ' según el idpfdetalle,idfdetalle
            End If
            
            vCodigoArticulo = EsNulo(.Fields("Codigo"))
            
            GuardarEnStock = Val(Format(GenerarDato("SELECT Sum(Entrada), Sum(Salida), Sum(Entrada-Salida) AS SaldoActual FROM Stock WHERE CodigoArticulo = '" & vCodigoArticulo & "'", "SaldoActual"), "#####0.00"))
            
            Call actualizastockEnArticulo(vCodigoArticulo, GuardarEnStock)

            
            .MoveNext
        Loop
        
        
    
    End With

    Call EjecutarScript(vsqlBorrar)

End Sub


Private Sub borrarLineaStockP(vid) ' borra la line teneindo en cuenta el id del pfdetalle
Dim vsql As String

vsql = "delete from stock where idPFdetalle=" + Str(vid)
Call EjecutarScript(vsql)
End Sub


Private Sub borrarLineaStockC(vid As Long) ' borra la line teneindo en cuenta el id del pfdetalle
Dim vsql As String

vsql = "delete from stock where idFdetalle=" + Str(vid)
Call EjecutarScript(vsql)
End Sub

Public Function ChequesFiltros(ByVal vfiltro As String, vorden As String, Optional vop As String)
Dim vvistacheques, vsql As String



If vop = "resumen" Then
 
 vvistacheques = "" + _
  " SELECT " + _
  "cheques.idCheques, cheques.marcainterna, 1 as ClienteProveedor, " + _
  " cheques.Fecha," + _
  " cheques.Codigo as CodCli," + _
  " cheques.Nombre as NomCli, cheques.endoso," + _
  "b.`idBancos` as CodBanco," + _
  "b.`Descripcion` as Banco," + _
  " cheques.Ncheque," + _
  " cheques.Firmante," + _
  " cheques.FechaDeposito," + _
  " cheques.Monto," + _
  " cheques.Observaciones, " + _
  " t.`Descripcion`  as EnCustodia, " + _
  " estadocheque.descripcion as Estado, " + _
  "  1, " + _
  "  1, " + _
  "  1, " + _
  "  1, 1,  1, cheques.nrointerno,cheques.sucursal, cheques.comentarios " + _
 " From " + _
 " cheques " + _
 " left JOIN bancos b ON (cheques.idBancos=b.idbancos) " + _
 " left JOIN estadocheque ON (estadocheque.idEstadoCheque=cheques.idEstadoCheque) " + _
 " left OUTER JOIN bancos t on (cheques.idCustodia = t.idbancos )"
'' cambié id por idbancos PANIC !!!

'" inner join bancos t on (cheques.idCustodia = t.idbancos )"

' " RIGHT OUTER JOIN cheques ON (bancosmovimientos.idCheques=cheques.idCheques) " + _

 ' vsql = vvistacheques + " where 1=1 and not cheques.marcainterna is null " + vFiltro + " group by cheques.idCheques order by  cheques." + vorden + " asc"
   
   vsql = vvistacheques + " where 1=1  " + vfiltro + " group by cheques.idCheques order by  cheques." + vorden + " asc"
   
   
End If



If vop = "historial" Then
 
 vvistacheques = "" + _
  " SELECT " + _
  " cheques.idCheques, " + _
  " cheques.Fecha," + _
  " cheques.Codigo as CodCli," + _
  " cheques.Nombre as NomCli, cheques.endoso, cheques.sucursal, cheques.marcainterna, " + _
  "b.`idBancos` as CodBanco," + _
  "b.`Descripcion` as Banco," + _
  " cheques.Ncheque," + _
  " cheques.Firmante," + _
  " cheques.FechaDeposito," + _
  " cheques.Monto," + _
  " cheques.Observaciones," + _
  " t.`Descripcion`  as EnCustodia, " + _
  " estadocheque.descripcion as Estado, " + _
  "  bancosmovimientos.Debito, " + _
  "  bancosmovimientos.Credito, " + _
  "  bancosmovimientos.Credito, " + _
  "  bancosmovimientos.ClienteProveedor, bancosmovimientos.fecha, cheques.comentarios " + _
 " From " + _
 " bancosmovimientos" + _
 " left OUTER JOIN cheques ON (bancosmovimientos.idCheques=cheques.idCheques) " + _
 " INNER JOIN bancos b ON (cheques.idBancos=b.idBancos) " + _
 " INNER JOIN estadocheque ON (estadocheque.idEstadoCheque=cheques.idEstadoCheque) " + _
 " inner join bancos t on (cheques.idCustodia = t.idbancos or cheques.idCustodia is null or cheques.idCustodia='') "

   
  vsql = vvistacheques + " where 1=1 " + vfiltro + " order by  cheques.fecha desc"
   
End If

ChequesFiltros = vsql

End Function




Public Function ChequesFiltrosRespaldo(ByVal vfiltro As String, vorden As String, Optional vop As String)
Dim vvistacheques, vsql As String



If vop = "resumen" Then
 
 vvistacheques = "" + _
  " SELECT " + _
  "cheques.idCheques, cheques.marcainterna, " + _
  " cheques.Fecha," + _
  " cheques.Codigo as CodCli," + _
  " cheques.Nombre as NomCli, cheques.endoso," + _
  "b.`idBancos` as CodBanco," + _
  "b.`Descripcion` as Banco," + _
  " cheques.Ncheque," + _
  " cheques.Firmante," + _
  " cheques.FechaDeposito," + _
  " cheques.Monto," + _
  " cheques.Observaciones, " + _
  " t.`Descripcion`  as EnCustodia, " + _
  " estadocheque.descripcion as Estado, " + _
  "  bancosmovimientos.Debito, " + _
  "  bancosmovimientos.Credito, " + _
  "  bancosmovimientos.Credito, " + _
  "  bancosmovimientos.ClienteProveedor, bancosmovimientos.fecha,  max( bancosmovimientos.`idBancosMovimientos`) as m, cheques.nrointerno,cheques.sucursal, cheques.comentarios " + _
 " From " + _
 " bancosmovimientos" + _
 " right join cheques ON (bancosmovimientos.idCheques=cheques.idCheques) " + _
 " left JOIN bancos b ON (cheques.idBancos=b.idbancos) " + _
 " left JOIN estadocheque ON (estadocheque.idEstadoCheque=cheques.idEstadoCheque) " + _
 " left OUTER JOIN bancos t on (cheques.idCustodia = t.idbancos )"
'' cambié id por idbancos PANIC !!!

'" inner join bancos t on (cheques.idCustodia = t.idbancos )"

' " RIGHT OUTER JOIN cheques ON (bancosmovimientos.idCheques=cheques.idCheques) " + _

 ' vsql = vvistacheques + " where 1=1 and not cheques.marcainterna is null " + vFiltro + " group by cheques.idCheques order by  cheques." + vorden + " asc"
   
   vsql = vvistacheques + " where 1=1  " + vfiltro + " group by cheques.idCheques order by  cheques." + vorden + " asc"
   
   
End If



If vop = "historial" Then
 
 vvistacheques = "" + _
  " SELECT " + _
  " cheques.idCheques, " + _
  " cheques.Fecha," + _
  " cheques.Codigo as CodCli," + _
  " cheques.Nombre as NomCli, cheques.endoso, cheques.sucursal, cheques.marcainterna, " + _
  "b.`idBancos` as CodBanco," + _
  "b.`Descripcion` as Banco," + _
  " cheques.Ncheque," + _
  " cheques.Firmante," + _
  " cheques.FechaDeposito," + _
  " cheques.Monto," + _
  " cheques.Observaciones," + _
  " t.`Descripcion`  as EnCustodia, " + _
  " estadocheque.descripcion as Estado, " + _
  "  bancosmovimientos.Debito, " + _
  "  bancosmovimientos.Credito, " + _
  "  bancosmovimientos.Credito, " + _
  "  bancosmovimientos.ClienteProveedor, bancosmovimientos.fecha, cheques.comentarios " + _
 " From " + _
 " bancosmovimientos" + _
 " left OUTER JOIN cheques ON (bancosmovimientos.idCheques=cheques.idCheques) " + _
 " INNER JOIN bancos b ON (cheques.idBancos=b.idBancos) " + _
 " INNER JOIN estadocheque ON (estadocheque.idEstadoCheque=cheques.idEstadoCheque) " + _
 " inner join bancos t on (cheques.idCustodia = t.idbancos or cheques.idCustodia is null or cheques.idCustodia='') "

   
  vsql = vvistacheques + " where 1=1 " + vfiltro + " order by  cheques.fecha desc"
   
End If

'ChequesFiltros = vsql

End Function




Public Sub limpiarChequesSeleccionados()

With gbldsCheques
        .vid = 0
End With
End Sub


Public Sub verTransacciones(vnrointerno As Long)
    frmTransaccionMantenimiento.vnrointerno = vnrointerno
    frmTransaccionMantenimiento.Show
End Sub


Public Sub stockAnular(ByVal vremito As Long, ByVal vfdetalle As String, ByVal vid As String)


Dim rsFDetalle As New ADODB.Recordset
Dim vsql  As String

vsql = "select * from " + vfdetalle + " where remito = " + Str(vremito)


    With rsFDetalle
        
        
        Call .Open(vsql, ConnDDBB, adOpenStatic, adLockReadOnly)
  
    Do Until .EOF
        
        vsql = "delete from stock where " + vid + " = " + Str(.Fields(vid))
        Call EjecutarScript(vsql, pathDBMySQL)
        
        Call stockActualizar2(.Fields("codigo"))
        
        .MoveNext
    Loop
    
    End With
    
  '  sqlDato = ""
    
End Sub



Public Sub stockActualizar2(vcodarticulo As String)
On Error Resume Next
           
           vsaldo = Val(Format(GenerarDato("SELECT Sum(Entrada), Sum(Salida), Sum(Entrada-Salida) AS SaldoActual FROM Stock WHERE  CodigoArticulo = '" & vcodarticulo & "'", "SaldoActual"), "#####0.00"))
           
           Call actualizastockEnArticulo(EsNulo(vcodarticulo), vsaldo)


If Err Then GrabarLog "ActualizarRubros", Err.Number & " " & Err.Description, "Procedimientos"

End Sub


Public Sub fijarparametro(vValor As String, vcampo)
On Error Resume Next
Dim vsql As String

    vsql = "update  configuracion set " + Trim(vcampo) + " = " + Trim(vValor)
    Call EjecutarScript(vsql, PathDBConfig)

If Err Then Exit Sub
End Sub

Public Function validarTransaccion(ByVal vnrointerno As Long, ByVal vIdUsuario As Integer) As Boolean
Dim vsql As String
Dim vvnrointerno As Long
Dim vidUsuario2 As Integer



vsql = "select * from transacciones where nrointerno=" + Str(vnrointerno)

vidUsuario2 = Val(traerDatos2(vsql, "idUsuario", pathDBMySQL))


If (vidUsuario2 = vIdUsuario) Or (vidUsuario2 = 0) Then
    validarTransaccion = True
Else

MsgBox "Esta transacción no pertenece al usuario: " + vConfigGral.vUser, vbInformation, "Mensaje"
    validarTransaccion = False
End If

End Function

Public Function wTransaccion(ByVal vnrointerno As Long, ByVal vIdUsuario As Integer) As Boolean
On Error Resume Next
Dim vsql As String

vsql = "insert into transacciones (nrointerno,idusuario) values (" + Str(vnrointerno) + "," + Str(vIdUsuario) + ")"
Call EjecutarScript(vsql, pathDBMySQL)

If Err Then Exit Function
End Function


Public Function dTransaccion(ByVal vnrointerno As Long) As Boolean
On Error Resume Next
Dim vsql As String
vsql = "delete from transacciones where nrointerno=" + Str(vnrointerno)

Call EjecutarScript(vsql, pathDBMySQL)

If Err Then Exit Function
End Function

Public Function TraerTipoMovimiento(vnrointerno As Long) As String
Dim valor, vsql As String

valor = ""

    vsql = "select * from factura where nrointerno=" + Str(vnrointerno)
    valor = traerDatos2(vsql, "TipoMovimiento", pathDBMySQL)

If valor = "" Then

    vsql = "select * from pfactura where nrointerno=" + Str(vnrointerno)
    valor = valor + traerDatos2(vsql, "TipoMovimiento", pathDBMySQL)

End If

If valor = "" Then

    vsql = "select * from asientos where nrointerno=" + Str(vnrointerno)
    valor = valor + traerDatos2(vsql, "TipoMovimiento", pathDBMySQL)

End If

TraerTipoMovimiento = valor


End Function

Public Sub CalcularMovimientos(ByVal vmarca As String, ByVal Index As Integer, ByVal vCodCuenta As String, _
vcuenta As String, vfdesde As Date, vfhasta As Date, vfDesdeBalance As Date, vfHastaBalance As Date, _
vbc As String, ByVal vnrobalance As Integer, Optional vtipo As String, Optional ByVal viddesde As Long, Optional ByVal vidhasta As Long)
'Public Sub CalcularMovimientos(ByVal vmarca As String, ByVal Index As Integer, ByVal vCodCuenta As Long, vcuenta As String, vfdesde As Date, vfhasta As Date, vfDesdeBalance As Date, vfHastaBalance As Date, vbc As String, ByVal vnrobalance As Integer, Optional vTipo As String)
On Error Resume Next

Dim vsqlFecha As String

If viddesde > 0 And vidhasta > 0 Then
        vsqlFecha = "idAsientos > " + Str(viddesde) + " and idAsientos <= " + Str(vidhasta)
Else
        vsqlFecha = "((fecha >= '" + strfechaMySQL(vfdesde) + "') AND (fecha <= '" + strfechaMySQL(vfhasta) + "'))"
End If


If Left(vCodCuenta, 1) = "9" Then
    Exit Sub
End If

    Dim vsqlTimeStamp As String
    Dim vpresupuestado As Double
    
    Dim basientos As New ADODB.Recordset, sqlAsientos, sqlMarca As String
    Dim vSaldoInicial, vSaldoPeriodo, vSaldoFinal As Double
  
    vTotalD = 0
    vTotalH = 0
    vsaldo = 0

    
    If vmarca = "TODOS" Then
    
    sqlMarca = ""
    
    Else
    
            If vmarca = "NORMAL" Then
                sqlMarca = " and (marca='" + vmarca + "' or marca is null)"
            Else
                sqlMarca = " and (marca='" + vmarca + "')"
            End If
    
    End If
    
    If "11102001" = vCodCuenta Then
     '   MsgBox ""
    End If
    
    
    
    If vbc = "2011-2012" Then
                sqlAsientos = "SELECT * FROM Asientos INNER JOIN AsientosDetalle ON Asientos.Numero = AsientosDetalle.Numero WHERE " + vsqlFecha + " AND (CodigoCuenta = " & vCodCuenta & ") " + vsqlTimeStamp + " ORDER BY Asientos.Numero ASC"
    
    Else
        
        If frmBalance.chkvarios = 1 Then
         
            ' sqlAsientos = "SELECT * FROM Asientos INNER JOIN AsientosDetalle ON Asientos.Numero = AsientosDetalle.Numero WHERE ((fecha >= '" + strfechaMySQL(vfdesde) + "') AND (fecha <= '" + strfechaMySQL(vfhasta) + "')) AND (CodigoCuenta = '" & vCodCuenta & "') and (Asientos.NroBalance=AsientosDetalle.NroBalance)" + vsqlTimeStamp + sqlMarca + " ORDER BY Asientos.Numero ASC"
     
             sqlAsientos = "SELECT * FROM Asientos INNER JOIN AsientosDetalle ON Asientos.Numero = AsientosDetalle.Numero WHERE " + vsqlFecha + " AND (fecha <= '" + strfechaMySQL(vfhasta) + "')) AND (CodigoCuenta = '" & vCodCuenta & "') and (Asientos.NroBalance=AsientosDetalle.NroBalance)" + vsqlTimeStamp + sqlMarca + " ORDER BY Asientos.Numero ASC"
     
            'sqlAsientos = "SELECT * FROM Asientos INNER JOIN AsientosDetalle ON Asientos.Numero = AsientosDetalle.Numero WHERE ((fecha >= '" + strfechaMySQL(vfdesde) + "') AND (fecha <= '" + strfechaMySQL(vfhasta) + "')) AND (CodigoCuenta = " & vCodCuenta & ") and (Asientos.NroBalance=" + Str(vnrobalance) + ") and (Asientos.NroBalance=AsientosDetalle.NroBalance)" + vsqlTimeStamp + sqlMarca + " ORDER BY Asientos.Numero ASC"
        Else
            sqlAsientos = "SELECT * FROM Asientos INNER JOIN AsientosDetalle ON Asientos.Numero = AsientosDetalle.Numero WHERE " + vsqlFecha + " AND (CodigoCuenta = '" & vCodCuenta & "') and (Asientos.NroBalance=" + Str(vnrobalance) + ") and (Asientos.NroBalance=AsientosDetalle.NroBalance)" + vsqlTimeStamp + sqlMarca + " ORDER BY Asientos.Numero ASC"
            'sqlAsientos = "SELECT * FROM Asientos INNER JOIN AsientosDetalle ON Asientos.Numero = AsientosDetalle.Numero WHERE ((fecha >= '" + strfechaMySQL(vfdesde) + "') AND (fecha <= '" + strfechaMySQL(vfhasta) + "')) AND (CodigoCuenta = " & vCodCuenta & ") and (Asientos.NroBalance=AsientosDetalle.NroBalance)" + vsqlTimeStamp + sqlMarca + " ORDER BY Asientos.Numero ASC"
           ' sqlAsientos = "SELECT * FROM Asientos INNER JOIN AsientosDetalle ON Asientos.Numero = AsientosDetalle.Numero WHERE ((fecha >= '" + strfechaMySQL(vfdesde) + "') AND (fecha <= '" + strfechaMySQL(vfhasta) + "')) AND (CodigoCuenta = " & vCodCuenta & ") and (Asientos.NroBalance=" + Str(vnrobalance) + ") and (Asientos.NroBalance=AsientosDetalle.NroBalance)" + vsqlTimeStamp + sqlMarca + " ORDER BY Asientos.Numero ASC"
   
        
        End If
    
    
    End If
    
    
    frmBalance.log.AddItem ("------------------------------------------------------------------")
    frmBalance.log.AddItem (vCodCuenta)
    frmBalance.log.AddItem ("------------------------------------------------------------------")
    
    
    'Dim vida, vidad As Long
    
    
    With basientos
        Call .Open(sqlAsientos, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .RecordCount = 0 Then .MoveFirst
        
        Do Until .EOF = True
            
            frmBalance.log.AddItem ("S. Acumulado: " + Str(vsaldo) + "            F: " + Str(.Fields("Fecha").Value) + "            D/H : " + Str(Val(Format(.Fields("Debe").Value, "#######0.000")) - Val(Format(.Fields("Haber").Value, "#######0.000"))))
            
            'vidad = .Fields("idAsientosDetalle")
            'vida = .Fields("idAsientos")
            
           ' Call cambiarNroBalance(vidad, vida)
      
            
            vsaldo = vsaldo + Val(Format(.Fields("Debe").Value, "#######0.000")) - Val(Format(.Fields("Haber").Value, "#######0.000"))
            vTotalD = vTotalD + Val(Format(.Fields("Debe").Value, "#######0.000"))
            vTotalH = vTotalH + Val(Format(.Fields("Haber").Value, "#######0.000"))
            
            vsaldo = CDbl(vsaldo)
            .MoveNext
            Debug.Print Str(Val(Format(.Fields("Debe").Value, "#######0.000"))) + "  -- " + Str(.Fields("Debe").Value)
            
        Loop
        
    End With
    
    
    ' ---------  calculo de los valores indirectos de cada renglón del balance ----------------------
    vSaldoInicial = CalSaldoAnteriorCtaContable(vmarca, vCodCuenta, vnrobalance, vfdesde, vbc, vfDesdeBalance, , viddesde, vidhasta)
    vSaldoPeriodo = vTotalD - vTotalH
    vSaldoFinal = vSaldoInicial + vSaldoPeriodo
    ' -----------------------------------------------------------------------------------------------
    
    
    'GuardarTemp Index, vCodCuenta, vCuenta, vfdesde, vfhasta, vSaldoInicial, vSaldoPeriodo, vSaldoFinal
  ''  GuardarTemp Index, vCodCuenta, vcuenta, vfdesde, vfhasta, vSaldoInicial, vSaldoPeriodo, vSaldoFinal
    
    vpresupuestado = calPresupuesto(vnrobalance, vCodCuenta)
    
    Debug.Print ">>>> " + Str(vSaldoPeriodo)
    
If frmBalance.rdPresupuestado Then
    GuardarTemp Index, vCodCuenta, vcuenta, vfdesde, vfhasta, Abs(vSaldoInicial), Abs(vSaldoPeriodo), Abs(vSaldoFinal), Abs(vpresupuestado)
Else
    GuardarTemp Index, vCodCuenta, vcuenta, vfdesde, vfhasta, vSaldoInicial, vSaldoPeriodo, vSaldoFinal, vpresupuestado
End If
If Err Then GrabarLog "CalcularMovimientos", Err.Number & " " & Err.Description, "Procedimientos"
End Sub


Function calPresupuesto(ByVal vnrobalance As Integer, ByVal vCodCuenta As String) As Double
On Error Resume Next
Dim vsql, vperiodo As String
Dim vidCuentas As Long


'If vCodCuenta = "01.01.01.0001.0033" Then MsgBox "."

vsql = "select codigo  from balances where NroBalance = " + Str(vnrobalance)
vperiodo = traerDatos2(vsql, "codigo", pathDBMySQL)


vsql = "select idCuentas  from Cuentas where CodigoCuenta = '" + (vCodCuenta) + "'"
vidCuentas = Val(traerDatos2(vsql, "idCuentas", pathDBMySQL))



vsql = "select importe from presupuesto where periodo = '" + vperiodo + "' and idcuentas = " + Str(vidCuentas)
calPresupuesto = traerDatos2(vsql, "importe", pathDBMySQL)



If Err Then
    calPresupuesto = 0
    Exit Function
End If

End Function

Public Sub ImprimirGrid1(g As MSHFlexGrid)
Dim iFila, iCol As Integer

Screen.MousePointer = vbHourglass

For iFila = 1 To g.Rows - 1
    For iCol = 1 To g.Cols - 1
        Printer.Print g.TextMatrix(iFila, iCol)
    Next iCol
Next iFila

Screen.MousePointer = vbDefault
Printer.EndDoc

End Sub


Public Sub ImprimirGrid2(g As MSHFlexGrid)

    Dim Ancho As Integer
    Ancho = g.Width
    Printer.Orientation = vbPRORLandscape
    
    g.Width = Printer.Width
    Printer.PaintPicture g.Picture, 0, 0
    Printer.EndDoc
    Printer.Orientation = vbPRORPortrait
    g.Width = Ancho

End Sub



Sub imprimirGrilla(grid As MSHFlexGrid, Optional vcol As Integer)
       Dim tppx As Integer
       Dim tppy As Integer
       tppx = Printer.TwipsPerPixelX
       tppy = Printer.TwipsPerPixelY
       Dim Col As Integer
       Dim Row As Integer
       Dim x0 As Single
       Dim y0 As Single
       Dim X1 As Single
       Dim Y1  As Single
       Dim X2  As Single
       Dim Y2  As Single
 
       x0 = Printer.CurrentX
       y0 = Printer.CurrentY
 
       If grid.BorderStyle <> 0 Then
          Printer.Line -Step(grid.Width - tppx, grid.Height - tppy), , B
          x0 = x0 + tppx
          y0 = y0 + tppy
       End If
       X1 = x0
       
       If vcol = 0 Or grid.Cols - 2 < vcol Then vcol = grid.Cols - 2
       
       
       For Col = 0 To vcol
          If Col >= grid.FixedCols And Col < grid.LeftCol Then
             Col = grid.LeftCol
          End If
          If X1 + grid.ColWidth(Col) >= grid.Width Then Exit For
          Y1 = y0
          For Row = 0 To grid.Rows - 1
             If Row >= grid.FixedRows And Row < grid.TopRow Then
                Row = grid.TopRow
            End If
             If Y1 + grid.RowHeight(Row) >= grid.Height Then Exit For
             Printer.CurrentX = X1 + tppx * 2
             Printer.CurrentY = Y1 + tppy
             grid.Col = Col
             grid.Row = Row
             Printer.Print grid.Text
             Y1 = Y1 + grid.RowHeight(Row)
             If grid.Gridlines Then
                Y1 = Y1 + tppy
             End If
          Next
          X1 = X1 + grid.ColWidth(Col)
          If grid.Gridlines Then
             X1 = X1 + tppx
          End If
       Next
       If grid.Gridlines Then
          X2 = x0
          Y2 = y0
          For Col = 0 To grid.Cols - 1
             If Col >= grid.FixedCols And Col < grid.LeftCol Then
                Col = grid.LeftCol
             End If
             X2 = X2 + grid.ColWidth(Col)
             If X2 >= grid.Width Then Exit For
             Printer.Line (X2, y0)-Step(0, Y1 - tppy)
             X2 = X2 + tppx
          Next
          For Row = 0 To grid.Rows - 1
             If Row >= grid.FixedRows And Row < grid.TopRow Then
                Row = grid.TopRow
             End If
             Y2 = Y2 + grid.RowHeight(Row)
             If Y2 >= grid.Height Then Exit For
             Printer.Line (x0, Y2)-Step(X1 - tppx, 0)
             Y2 = Y2 + tppy
          Next
       End If
       
        Printer.EndDoc
    End Sub
 
Public Sub initRollbk(vnrointerno As Long)
 Dim vsql1 As String
 vrollbk_nrointerno = 0
 vrollbk_nroasiento = 0
 vrollbk = False
 
 vsql1 = "insert into t_rollback (nrointerno) values (" + Trim(vnrointerno) + ")"

 Call EjecutarScript(vsql1, pathDBMySQL)

End Sub

Public Sub endRollbk(vnrointerno)
On Error Resume Next
Dim vsql1 As String

vsql1 = "truncate t_rollback"

Call EjecutarScript(vsql1, pathDBMySQL)

If Err Then Exit Sub
End Sub


Public Sub endRollbk22(vnrointerno, vnroasiento)
Dim vsql As String
Dim v1, v2, v3 As Long

v1 = 0
v2 = 0
v3 = 0

vsql = "select nrointerno as c from bancosmovimientos where nrointerno= " + Str(vnrointerno) + " limit 1"
v1 = Val(TraerDato2(vsql, "c", pathDBMySQL))


If vnroasiento > 0 Then
    vsql = "select numero as c from asientos where numero= " + Str(vnroasiento) + " limit 1"
    v2 = Val(TraerDato2(vsql, "c", pathDBMySQL))
Else
    v2 = 1
End If


If vnroasiento > 0 Then
    vsql = "select numero as c from asientosdetalle where numero= " + Str(vnroasiento) + " limit 1"
    v3 = Val(TraerDato2(vsql, "c", pathDBMySQL))
Else
    v3 = 1
End If




If v1 = 0 Or v2 = 0 Then

    MsgBox "Atención. Hubo un error grabe al intentar  grabar el movimiento." + Chr(13) + "Debe repetir la operación ", vbCritical

    vsql = "delete from asientos where numero =" + Str(vnrointerno)
    Call EjecutarScript(vsql, pathDBMySQL)

    vsql = "delete from asientosdetalle where numero =" + Str(vnroasiento)
    Call EjecutarScript(vsql, pathDBMySQL)
    
    vsql = "delete from bancosmovimientos where nrointerno =" + Str(vnroasiento)
    Call EjecutarScript(vsql, pathDBMySQL)

    validado = False

Else

    
   Call frmAlert.DisplayAlert("El movimiento se guardó correctamente", 1000)
    
End If


End Sub


Function hayFacturasImpagas(vcodigo As String)
Dim vsql As String

hayFacturasImpagas = False
 
vsql = " select " + _
" count(t.idpFactura) as c " + _
"  from pfactura t " + _
" where ( estadodocumento is null or  estadodocumento = 'adeudado') " + _
"  and codigo = '" + vcodigo + "'"

If Val(traerDatos2(vsql, "c", pathDBMySQL)) > 0 Then

    If MsgBox("Hay facturas pendientes de pago" + Chr(13) + "Quiere ir a seleccionarlas ?", vbYesNo) = vbYes Then
        hayFacturasImpagas = True
    End If
    
End If

End Function



Function getSaldoCliente2(vcod As String) As Double
Dim vsql As String

vsql = " select " + _
" sum(t.Debito) - sum(t.Credito) as c " + _
" from  " + _
" cuentascorrientes t " + _
" where codigo = '" + vcod + "'"

getSaldoCliente2 = Val(TraerDato2(vsql, "c", pathDBMySQL))

End Function

Function getSaldoProveedor2(vcod As String) As Double
Dim vsql As String

vsql = " select " + _
" sum(t.Debito) - sum(t.Credito) as c " + _
" from  " + _
" cuentascorrientes t " + _
" where codigo = '" + vcod + "'"

getSaldoProveedor2 = Val(TraerDato2(vsql, "c", pathDBMySQL))

End Function


Function getSaldoProveedor22(vcod As String) As Double
Dim vsql As String

vsql = " select " + _
" sum(t.Debito) - sum(t.Credito) as c " + _
" from  " + _
" pcuentascorrientes t " + _
" where codigo = '" + vcod + "'"

getSaldoProveedor22 = Val(TraerDato2(vsql, "c", pathDBMySQL))

End Function


'Public Sub FlexGrid_To_Excel(TheFlexgrid As MSFlexGrid, _
'  TheRows As Integer, TheCols As Integer, _
'  Optional GridStyle As Integer = 1, Optional WorkSheetName _
'  As String)
'
'Dim objXL As New Excel.Application
'Dim wbXL As New Excel.Workbook
'Dim wsXL As New Excel.Worksheet
'Dim intRow As Integer ' counter
'Dim intCol As Integer ' counter
'
'If Not IsObject(objXL) Then
'    MsgBox "You need Microsoft Excel to use this function", _
'       vbExclamation, "Print to Excel"
'    Exit Sub
'End If
'
''On Error Resume Next is necessary because
''someone may pass more rows
''or columns than the flexgrid has
'
''you can instead check for this,
''or rewrite the function so that
''it exports all non-fixed cells
''to Excel
'
'On Error Resume Next
'
'' open Excel
'objXL.Visible = True
'Set wbXL = objXL.Workbooks.Add
'Set wsXL = objXL.ActiveSheet
'
'' name the worksheet
'With wsXL
'    If Not WorkSheetName = "" Then
'        .Name = WorkSheetName
'    End If
'End With
'
'' fill worksheet
'For intRow = 1 To TheRows
'    For intCol = 1 To TheCols
'        With TheFlexgrid
'            wsXL.Cells(intRow, intCol).Value = _
'               .TextMatrix(intRow - 1, intCol - 1) & " "
'        End With
'    Next
'Next
'
'' format the look
'For intCol = 1 To TheCols
'    wsXL.Columns(intCol).AutoFit
'    'wsXL.Columns(intCol).AutoFormat (1)
'    wsXL.Range("a1", Right(wsXL.Columns(TheCols).AddressLocal, _
'         1) & TheRows).AutoFormat GridStyle
'Next
'
'End Sub

Public Sub rsToExcel(ByVal vsql As String)

On Error Resume Next

Dim rss As New ADODB.Recordset
Call rss.Open(vsql, ConnDDBB, adOpenStatic, adLockPessimistic)
        
            Dim iErr As Integer
            iErr = 0
            On Error GoTo Proc_Err
           ' Screen.MousePointer = vbHourglass
            Dim i, ii  As Long
            Dim x  As Long
            Dim Cols As Integer
            Dim Rows As Integer
            Dim sLine As String
            Open App.Path + "\l.csv" For Output As #1
            
rss.MoveFirst
            
        Do Until rss.EOF
           
           sLine = ""
           
           'For Each fld In rs.Fields
           For ii = 0 To rss.Fields.Count - 1
                 sLine = sLine & rss.Fields(i) & IIf(i < Cols - 1, ";", "")
           Next ii
               Print #1, sLine
               rss.NextRecordset
        Loop
            
            Close #1
    
    Call Shell("excel.bat", 1)

Proc_Exit:
            Screen.MousePointer = vbDefault
            Exit Sub
Proc_Err:
            If iErr > 3 Then
            Close #1
               ' Log your error here...
               Resume Proc_Exit
            Else
             Close #1
               iErr = iErr + 1
               Resume
            End If
       
End Sub


Public Sub grillaToExcel(grilla As MSHFlexGrid, Optional vtitulo As String)
On Error Resume Next
Dim iErr As Integer
            iErr = 0
            On Error GoTo Proc_Err
            Screen.MousePointer = vbHourglass
            Dim i, ii  As Long
            Dim x  As Long
            Dim Cols As Integer
            Dim Rows As Integer
            Dim sLine As String
            Open App.Path + "\l.csv" For Output As #1
            
            Cols = grilla.Cols
            
            
          ' imprimio el tìtulo
          Print #1, Replace(vtitulo, vbCrLf, "")
            
           For ii = 0 To grilla.Rows - 1
            
            'Do While Not rs.EOF
               sLine = ""
               For i = 1 To Cols - 1
                  'sLine = sLine & rs.Fields(i).Value & IIf(i < Cols - 1, ";", "")
                  sLine = sLine & Trim(grilla.TextMatrix(ii, i)) & IIf(i < Cols - 1, ";", "")
               Next i
               Print #1, Replace(sLine, vbCrLf, "")
              ' rs.MoveNext
            'Loop
            
            Next
            
            Close #1
    
    Call Shell("excel.bat", 1)

Proc_Exit:
            Screen.MousePointer = vbDefault
            Exit Sub
Proc_Err:
            If iErr > 3 Then
               ' Log your error here...
               Resume Proc_Exit
            Else
               iErr = iErr + 1
               Resume
            End If
       
End Sub

Public Sub grillaToExcel3(grilla As DataGrid, vrows As Integer)
On Error Resume Next
Dim iErr As Integer
            iErr = 0
            On Error GoTo Proc_Err
            Screen.MousePointer = vbHourglass
            Dim i, ii  As Long
            Dim x  As Long
            Dim Cols As Integer
            Dim Rows As Integer
            Dim sLine As String
            Open App.Path + "\l.csv" For Output As #1
            
            Cols = grilla.Columns.Count
            
            
           For ii = 0 To vrows - 1
            
            'Do While Not rs.EOF
               sLine = ""
               For i = 1 To Cols - 1
                  'sLine = sLine & rs.Fields(i).Value & IIf(i < Cols - 1, ";", "")
                  grilla.Col = i
                  grilla.Row = ii
                  
                 
                  sLine = sLine & Trim(Left(grilla.Text, 40)) & IIf(i < Cols - 1, ";", "")
               Next i
               Debug.Print "Linea: " + Trim(sLine)
               Print #1, sLine
              ' rs.MoveNext
            'Loop
            
            Next
            
            Close #1
    
   ' Call Shell("excel.bat", 1)

Proc_Exit:
            
            Close #1
            Debug.Print "II  " + Str(ii)
           Call Shell("excel.bat", 1)
            
            Screen.MousePointer = vbDefault
            Exit Sub
Proc_Err:
            If iErr > 3 Then
               ' Log your error here...
               Resume Proc_Exit
            Else
               iErr = iErr + 1
               Resume
            End If
       
End Sub



Public Sub grillaToExcel2(grilla As KlexGrid)
On Error Resume Next
Dim iErr As Integer
            iErr = 0
            On Error GoTo Proc_Err
            Screen.MousePointer = vbHourglass
            Dim i, ii  As Long
            Dim x  As Long
            Dim Cols As Integer
            Dim Rows As Integer
            Dim sLine As String
            Open App.Path + "\l.csv" For Output As #1
            
            Cols = grilla.Cols
            
            
           For ii = 0 To grilla.Rows - 1
            
            'Do While Not rs.EOF
               sLine = ""
               For i = 1 To Cols - 1
                  'sLine = sLine & rs.Fields(i).Value & IIf(i < Cols - 1, ";", "")
                  sLine = sLine & grilla.TextMatrix(ii, i) & IIf(i < Cols - 1, ";", "")
               Next i
               Print #1, sLine
              ' rs.MoveNext
            'Loop
            
            Next
            
            Close #1
    
    Call Shell("excel.bat", 1)

Proc_Exit:
            Screen.MousePointer = vbDefault
            Exit Sub
Proc_Err:
            If iErr > 3 Then
               ' Log your error here...
               Resume Proc_Exit
            Else
               iErr = iErr + 1
               Resume
            End If
       
End Sub

Public Function RecordsetToCSV(rsData As ADODB.Recordset, Optional ShowColumnNames As Boolean = True, Optional NULLStr As String = "") As String
    Dim k As Long, RetStr As String
    
    
    If ShowColumnNames Then
        For k = 0 To rsData.Fields.Count - 1
            RetStr = RetStr & "," & rsData.Fields(k).Name
        Next k
            RetStr = RetStr + ""
        RetStr = Mid(RetStr, 2) & vbNewLine
    End If
    
    RetStr = RetStr & """" & rsData.GetString(adClipString, -1, """,""", """" & vbNewLine & """", NULLStr)
    RetStr = Left(RetStr, Len(RetStr) - 3)
    
    RecordsetToCSV = RetStr
End Function


Public Sub FPDelay()
'
' Delay Sequence
'

    Dim Start1 As Single
    Start1 = Timer                  '
    Do While frmPrincipal.FiscalEpson2.State = EFP_S_Busy
       

        Do While Timer < Start1 + 0.125     '   Timer delay
            DoEvents
            
            If Start1 > Timer Then          '   This is to
                Exit Do                     '   compensate for the
            End If                          '   AM to PM change
        Loop
    Loop

End Sub


Public Sub ShowMsg()
Exit Sub ' ojo que esto está agregado
    MsgBox "Código de Retorno: " + Format(Hex(frmPrincipal.FiscalEpson2.ReturnCode), "0000") _
                + vbCrLf + "Estado Impresora: " + Format(Hex(frmPrincipal.FiscalEpson2.PrinterStatus), "0000") _
                + vbCrLf + "Estado Fiscal: " + Format(Hex(frmPrincipal.FiscalEpson2.FiscalStatus), "0000"), _
                vbOKOnly + vbExclamation, "Información de respuesta"
End Sub


Public Function validarCUIT2(cuit As String) As Integer
    'definimos variables
    Dim verificador As Integer  'verificador es un número entero
    Dim resultado As Integer    'resultado es un número entero
    Dim mult() As Variant     'mult es una matriz de 10 números
    Dim x As Integer
    
    'eliminamos los guiones de la CUIT en caso de que existan
    cuit = Replace(cuit, "-", "")
    
    'Si la cantidad de dígitos es distinto a 11, la función retorna -1
    If Len(cuit) <> 11 Then
        validarCUIT2 = -1
        Exit Function
        'el código que sigue no se ejecuta
    End If
    
    'Asignamos valores a las variables definidas arriba
    mult = Array(5, 4, 3, 2, 7, 6, 5, 4, 3, 2)
    verificador = Right(cuit, 1)    'verificador es igual al último dígito de la CUIT
    
    For x = 0 To 9  'recorremos todos los dígitos de la CUIT y acumulamos en la variable "resultado"
        resultado = resultado + (mult(x) * Mid(cuit, x + 1, 1))
    Next
    
    'obtenemos el resto de dividir "resultado" con 11
    resultado = resultado Mod 11
    
    'restamos
    resultado = 11 - resultado
    
    If resultado = 11 Then
        'si resultado es igual a 11, cambiamos a 0
        resultado = 0
    ElseIf resultado = 10 Then
        'si resultado es igual a 10, cambiamos a 9
        resultado = 9
    End If
    
    If resultado = verificador Then
        'Si ambas variables coinciden, significa que la CUIT es válida y
        'retornamos 1
        validarCUIT2 = 1
    Else
        'Si la CUIT no es válida, retornamos 0
        validarCUIT2 = 0
    End If

End Function

Public Sub logform(vdatos As String)
' log para identificar error en el nro de cuit para el citi
frmLog.Show
frmLog.log2.AddItem vdatos

With frmLog.log
    .AddItem (vdatos)
End With
End Sub


Function getvcodEmpresa(vid As Long) As String
Dim vsql As String
Dim vvid As Long

vsql = "select idEmpresa as c from t_rel where idfactura=" + Str(vid)
vvid = traerDatos2(vsql, "c", pathDBMySQL)

vsql = "select codigo as c from proveedores where idproveedores = " + Str(vvid)
getvcodEmpresa = traerDatos2(vsql, "c", pathDBMySQL)

End Function



Function getvdesEmpresa(vid As Long) As String
Dim vsql As String
Dim vvid As Long

vsql = "select idEmpresa as c from t_rel where idfactura=" + Str(vid)
vvid = traerDatos2(vsql, "c", pathDBMySQL)

vsql = "select nombre as c from proveedores where idproveedores = " + Str(vvid)
getvdesEmpresa = traerDatos2(vsql, "c", pathDBMySQL)

End Function


Function getRepartidor2idFactura(vcod As String) As String
Dim vsql As String
Dim vvid As Long

        If Not LeerXml("Puesto") = "Empresas" Then
        
            getRepartidor2idFactura = " (select codigo as c from clientes where idvendedor2 = " + vcod + ")"
 
        Else
               
            getRepartidor2idFactura = " (select t.idfactura as c from t_rel t inner join proveedores p on t.idvendedor = p.idProveedores where Codigo = '" + Trim(vcod) + "')"
       '    getRepartidor2idFactura = Val(traerDatos2(vsql, "c", pathDBMySQL))

        End If
End Function


Function getEmpresa2idFactura(vcod As String) As String
Dim vsql As String
Dim vvid As Long
        getEmpresa2idFactura = " (select t.idfactura as c from t_rel t inner join proveedores p on t.idEmpresa = p.idProveedores where codigo = '" + Trim(vcod) + "')"
        'getEmpresa2idFactura = Val(traerDatos2(vsql, "c", pathDBMySQL))
End Function


Function getvcodRepartidor(vid As Long) As String
Dim vsql As String
Dim vvid As Long

vsql = "select idvendedor as c from t_rel where idfactura=" + Str(vid)
vvid = traerDatos2(vsql, "c", pathDBMySQL)

vsql = "select codigo as c from proveedores where idproveedores = " + Str(vvid)
getvcodRepartidor = traerDatos2(vsql, "c", pathDBMySQL)

End Function



Function getCuitFE(ByVal vempresa As String) As String
Dim vsql As String

vsql = "select cuit as c  from proveedores where codigo  = '" + vempresa + "'"

    If UCase(LeerXml("Puesto")) = UCase("Empresas") And vempresa > 0 Then
        getCuitFE = traerDatos2(vsql, "c", pathDBMySQL)
    Else
        getCuitFE = Trim(LeerXml("vcuit"))
    End If

End Function




Function getLicenciaFE(ByVal vempresa As String) As String
    
    If UCase(LeerXml("Puesto")) = UCase("Empresas") And Val(vempresa) > 0 Then
        getLicenciaFE = "Empresa" + Trim(vempresa) + ".lic"
    Else
        getLicenciaFE = Trim(LeerXml("LicenciaWSAFIP"))
    End If

End Function



Function getCertificadoFE(ByVal vempresa As String) As String
    
    If UCase(LeerXml("Puesto")) = UCase("Empresas") And Val(vempresa) > 0 Then
        getCertificadoFE = "Empresa" + Trim(vempresa) + ".pfx"
    Else
        getCertificadoFE = Trim(LeerXml("vcertificado"))
    End If

End Function


Function getvdesRepartidor(vid As Long) As String
Dim vsql As String
Dim vvid As Long

vsql = "select idvendedor as c from t_rel where idfactura=" + Str(vid)
vvid = traerDatos2(vsql, "c", pathDBMySQL)

vsql = "select nombre as c from proveedores where idproveedores = " + Str(vvid)
getvdesRepartidor = traerDatos2(vsql, "c", pathDBMySQL)

End Function


Function verifico_ult_nrointerno_todasTablas()

mensaje "+ verifico_ulr_nrointero"

Dim vsql, vsql2, vsql1 As String
Dim vnumero, vmaxNroInterno As Long

vsql = "select max(fact) as c from (select max(nrointerno) as fact from factura  union select max(nrointerno) as pfact from  pfactura  union select max(t.NroInterno) as cc from  cuentascorrientes t union select max(nrointerno) as pcc from  pcuentascorrientes union select max(nrointerno) as asiento from asientos union select max(nrointerno) as bmovi from  bancosmovimientos union select max(nrointerno) as cheq from  cheques ) a"

vnumero = traerDatos2(vsql, "c", pathDBMySQL)

vsql1 = "select max(numero) as c from t_nrointerno"

vmaxNroInterno = traerDatos2(vsql1, "c", pathDBMySQL)


If vmaxNroInterno < vnumero Then

    vsql2 = "ALTER TABLE t_nrointerno AUTO_INCREMENT =" + Str(vnumero)
   
    Call EjecutarScript(vsql2, pathDBMySQL)
    

End If


mensaje "- verifico_ulr_nrointero"

End Function

