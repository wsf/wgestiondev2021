VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "KlexGrid.ocx"
Begin VB.Form frmMigracion 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox log 
      Height          =   3765
      Left            =   180
      TabIndex        =   11
      Top             =   2520
      Width           =   6285
   End
   Begin XtremeSuiteControls.PushButton cmdCtaCte 
      Height          =   375
      Left            =   9720
      TabIndex        =   10
      Top             =   6960
      Visible         =   0   'False
      Width           =   1935
      _Version        =   851968
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
   End
   Begin Grid.KlexGrid KlexCtaCte 
      Height          =   4215
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   7435
      EnterKeyBehaviour=   0
      BackColorAlternate=   0
      GridLinesFixed  =   2
      BackColorFixed  =   -2147483626
      Cols            =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColorFixed  =   8421504
      MouseIcon       =   "frmMigracion.frx":0000
      Rows            =   10
   End
   Begin XtremeSuiteControls.GroupBox GBCliente 
      Height          =   1215
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   6255
      _Version        =   851968
      _ExtentX        =   11033
      _ExtentY        =   2143
      _StockProps     =   79
      Caption         =   "Elegir Cliente"
      Appearance      =   6
      Begin XtremeSuiteControls.RadioButton RBCliente 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "O. Valentini"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton RBCliente 
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   5
         Top             =   480
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "WServicios"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton RBCliente 
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   6
         Top             =   480
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "M. Tavani"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton RBCliente 
         Height          =   255
         Index           =   3
         Left            =   4440
         TabIndex        =   7
         Top             =   480
         Width           =   1365
         _Version        =   851968
         _ExtentX        =   2408
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "G. Arneri (ARV)"
         Appearance      =   6
         Value           =   -1  'True
      End
   End
   Begin XtremeSuiteControls.ProgressBar Barra 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1380
      Width           =   6255
      _Version        =   851968
      _ExtentX        =   11033
      _ExtentY        =   661
      _StockProps     =   93
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton PBAcciones 
      Height          =   435
      Index           =   0
      Left            =   3360
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
      _Version        =   851968
      _ExtentX        =   2778
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "Ejecutar !!"
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton PBAcciones 
      Default         =   -1  'True
      Height          =   435
      Index           =   1
      Left            =   4920
      TabIndex        =   0
      Top             =   1920
      Width           =   1575
      _Version        =   851968
      _ExtentX        =   2778
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "Salir"
      Appearance      =   6
   End
   Begin Grid.KlexGrid KlexCtaCte 
      Height          =   4215
      Index           =   1
      Left            =   5880
      TabIndex        =   9
      Top             =   2640
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7435
      EnterKeyBehaviour=   0
      BackColorAlternate=   0
      GridLinesFixed  =   2
      BackColorFixed  =   -2147483626
      Cols            =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColorFixed  =   8421504
      MouseIcon       =   "frmMigracion.frx":001C
      Rows            =   10
   End
End
Attribute VB_Name = "frmMigracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim connDDBBMigrar As ADODB.Connection
Dim vnroremito As Long
Private Sub cmdCtaCte_Click()
On Error Resume Next
    
    Dim i As Integer, j As Integer
    Dim vcredito As Double, vImporteParcial As Double
    Dim vPagado As Double
    
    With KlexCtaCte(0)
        For i = 1 To Val(.Rows - 1)
            vcredito = 0
            vcredito = .TextMatrix(i, 4)
            
            
            For j = 1 To Val(KlexCtaCte(1).Rows - 1)
                If Val(KlexCtaCte(0).TextMatrix(i, 5)) = 0 Then
                    If vcredito <= Val(KlexCtaCte(1).TextMatrix(j, 3)) Then
                        .TextMatrix(i, 5) = Val(KlexCtaCte(1).TextMatrix(j, 5))
                        KlexCtaCte(1).TextMatrix(j, 6) = vcredito
                        Exit For
                        
                    Else
                    
                    
                        
                    End If
            
                End If
            Next j
    
        Next
        
        
    End With


If Err Then GrabarLog "cmdCtaCte_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Load()
On Error Resume Next

    Me.Show
    Call AsignaPagos(0, "Rodave")
    Call AsignaPagos(1, "Rodave")

If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub AsignaPagos(Index As Integer, vcodigo As String)
On Error Resume Next

    Dim rsPagosACtaCte As New ADODB.Recordset, sqlPagosACtaCte As String
    
    If Index = 0 Then
        sqlPagosACtaCte = "SELECT * FROM PCuentasCorrientes WHERE Codigo = '" & vcodigo & "' AND (Credito >0) ORDER BY Fecha ASC"
    Else
        sqlPagosACtaCte = "SELECT * FROM PCuentasCorrientes WHERE Codigo = '" & vcodigo & "' AND (Debito >0) ORDER BY Fecha ASC"
    End If
    
    With rsPagosACtaCte
        Call .Open(sqlPagosACtaCte, ConnDDBB, adOpenStatic, adLockReadOnly)
    
        If Not .EOF = True Then
            .MoveFirst
            Call FormatoGrilla(Index, .RecordCount)
        Else
            Call FormatoGrilla(Index, 1)
        End If
        
        Do Until .EOF = True
            KlexCtaCte(Index).TextMatrix(.AbsolutePosition, 1) = .Fields("idPCuentasCorrientes").Value
            KlexCtaCte(Index).TextMatrix(.AbsolutePosition, 2) = .Fields("Fecha").Value
            KlexCtaCte(Index).TextMatrix(.AbsolutePosition, 3) = .Fields("Debito").Value
            KlexCtaCte(Index).TextMatrix(.AbsolutePosition, 4) = .Fields("Credito").Value
            KlexCtaCte(Index).TextMatrix(.AbsolutePosition, 5) = .Fields("Remito").Value
            KlexCtaCte(Index).TextMatrix(.AbsolutePosition, 6) = ""
            
            .MoveNext
        Loop
    
    End With

    sqlPagosACtaCte = ""

    If rsPagosACtaCte.State = 1 Then
        rsPagosACtaCte.Close
        Set rsPagosACtaCte = Nothing
    End If
    
If Err Then GrabarLog "AsignaPagos", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub FormatoGrilla(Index As Integer, vCantidadRenglones As Long)
On Error Resume Next

    Dim i As Integer
   

    With Me.KlexCtaCte(Index)
        
        .FixedRows = 1
        .FixedCols = 1

        .Cols = 7
        .Rows = vCantidadRenglones + 1
    
        If vCantidadRenglones = 1 Then
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
                .ColWidth(i) = 0
            Next
        End If
    
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 200
    
        .TextMatrix(0, 1) = "ID"
        .ColWidth(1) = 750
    
        .TextMatrix(0, 2) = "Fecha"
        .ColWidth(2) = 1000
           
        .TextMatrix(0, 3) = "Debito"
        .ColWidth(3) = 1000
        .ColDisplayFormat(3) = "#0.00"
    
        .TextMatrix(0, 4) = "Credito"
        .ColWidth(4) = 1000
        .ColDisplayFormat(4) = "#0.00"
        
        .TextMatrix(0, 5) = "Remito"
        .ColWidth(5) = 1000
    
        .TextMatrix(0, 6) = "Pagado"
        .ColWidth(6) = 500
            
    
        .BackColorAlternate = &HC0C0C0
    End With

If Err Then GrabarLog "FormatoGriila", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub PbAcciones_Click(Index As Integer)
On Error Resume Next

    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
    
    PbAcciones(0).Enabled = False
    PbAcciones(1).Enabled = False
    
    Dim i As Integer

    For i = 0 To Val(RBCliente.Count - 1)
    
        If RBCliente(i).Value = True Then
            Select Case i
                
                Case 0
                    ConectarDB ("OscarValentini")
                    
                    MigrarClientes
                    MigrarArticulos
                    MigrarFacturas
                    MigrarFDetalles
                    MigrarCtaCteClientes
                    MigrarRubros
                    MigrarLibretaPagos
                
                Case 1
                    ConectarDB ("WGestionServicios")

                    MigrarClientes
                    MigrarArticulos
                    MigrarRubros
                    MigrarEmpleados
                
                Case 2
                    ConectarDB ("MTavani")

                    MigrarProveedores
                    MigrarClientes
                    MigrarRubros
                    MigrarArticulos
                    MigrarFacturas
                    MigrarFDetalles
                    MigrarCtaCteClientes
                    MigrarLibretaPagos

        
                Case 3
                    'ARV (G. Arneri)
                    
                    
                    ''''Call ConectarDB(".DBF")
                    ''''ABMO00.DBF (MOVIMIENTOS DE CAJA Y BANCOS EGRESOS E INGRESOS)
                    ''''ABIC00.DBF (IMPUTACION CONTABLE DE BANCOS)
                    ''''APIC00.DBF (IMPUTACION CONTABLE DE PROVEEDORES)
                    ''''AVIC00.DBF (IMPUTACION CONTABLE DE CLIENTES)
                    ''''Call MigrarCuentasCorrientesCSA("AVCC00.DBF")
                    ''''Call MigrarCuentasCorrientesPSA("APCC00.DBF")
                    
                    Call EjecutarScript("CALL SP_VaciarTablas()") 'ojo: sacar las que no hay que vaciar
                    Call ConectarDB("ARV")
                    
'                    Call MigrarAsientosTipoSA("APCA00")
'                    Call MigrarCuentasSA("AXCU00")
'                    Call MigrarProveedoresSA("APPR00")
'                    Call MigrarClientesSA("AXCL00")
'
'                    Call MigrarBancosSA("AXBA00")
'                    Call MigrarBancosCuentasSA("APBC00")
'
'                    Call MigrarBancosMovimientosSA("ABIC00")
'                    Call MigrarBancosClientesSA("BancosCajaClientes")
'                    Call MigrarBancosProveedoresSA("BancosCajaProveedores")
'
'                   Call MigrarAsientosSA("ACEA00")
'                    Call MigrarAsientosDetalleSA("ACDA00")
'
'                    Call MigrarCuentasCorrientesCSA("CtaCteV")
                    Call MigrarCuentasCorrientesPSA("CtaCteP")
                    
                    Call MigrarChequesCSA("AVCH00")
                    Call MigrarChequesPSA("APCH00")
                    
                    Call MigrarCobrosSA("AVRE00_c")
                    Call MigrarPagosSA("APRE00")
                    
                    ' APC00  es la tabla de ctacte proveedores que es llamada por la consulta CtaCteP en el access
                    
                    'Call MigrarCajasSA("AVCH00")
                    'Call MigrarCajasSA("APCH00")
                    'Call MigrarBancosMovSA("AVCH00")
                    'Call MigrarBancosMovSA("APCH00")
                    'Call MigrarPagosPorNotaCSA("APRE00")
                    'Call MigrarCobrosPorNotaCSA("AVRE00")
                    
            End Select
    
        End If
    
    Next
    
    PbAcciones(0).Enabled = True
    PbAcciones(1).Enabled = True

If Err Then GrabarLog "PbAcciones_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub MigrarEmpleados()
On Error Resume Next

    Dim rsOrigen As New ADODB.Recordset, rsDestino As New ADODB.Recordset
    Dim sqlOrigen As String, sqlDestino As String
    
    sqlOrigen = "SELECT * FROM Empleados WHERE 1=1"
    sqlDestino = "SELECT * FROM Empleados WHERE 1=2"
    
    With rsOrigen
        .CursorLocation = adUseClient
        Call .Open(sqlOrigen, connDDBBMigrar, adOpenStatic, adLockPessimistic)
        
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
        
        .MoveFirst
        barra.Value = 0
        barra.Max = .RecordCount
        
    End With
    
    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            Do Until rsOrigen.EOF = True
                .AddNew
                
                .Fields("Codigo").Value = EsNulo(rsOrigen.Fields("Codigo").Value)
                .Fields("CodigoNum").Value = Val(rsOrigen.Fields("Codigo").Value)
                .Fields("Nombre").Value = EsNulo(rsOrigen.Fields("Nombre").Value)
                .Fields("Direccion").Value = EsNulo(rsOrigen.Fields("Direccion").Value)
                .Fields("Localidad").Value = EsNulo(rsOrigen.Fields("Localidad").Value)
                .Fields("Telefono").Value = EsNulo(rsOrigen.Fields("Telefono").Value)
                .Fields("Iva").Value = EsNulo(rsOrigen.Fields("Iva").Value)
                .Fields("Cuit").Value = EsNulo(rsOrigen.Fields("Cuit").Value)
                .Fields("Credito").Value = EsNulo(rsOrigen.Fields("Credito").Value)
                .Fields("Responsable").Value = EsNulo(rsOrigen.Fields("Responsable").Value)
                .Fields("IBrutos").Value = EsNulo(rsOrigen.Fields("IBrutos").Value)
                .Fields("Quebranto").Value = EsNulo(rsOrigen.Fields("Quebranto").Value)
                
                .Update
                
                rsOrigen.MoveNext
                barra.Value = barra.Value + 1
            Loop
        End If
        
    End With
    
    sqlOrigen = ""
    sqlDestino = ""

    If rsOrigen.State = 1 Then
        rsOrigen.Close
        Set rsOrigen = Nothing
    End If
    
If Err Then GrabarLog "MigrarEmpleados", Err.Number & " " & Err.Description, "Migracion"
End Sub
Private Sub MigrarRubros()
On Error Resume Next

    Dim rsOrigen As New ADODB.Recordset, rsDestino As New ADODB.Recordset
    Dim sqlOrigen As String, sqlDestino As String
    
    sqlOrigen = "SELECT * FROM Rubros WHERE 1=1"
    sqlDestino = "SELECT * FROM Rubros WHERE 1=2"
    
    With rsOrigen
        .CursorLocation = adUseClient
        Call .Open(sqlOrigen, connDDBBMigrar, adOpenStatic, adLockPessimistic)
        
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
        
        .MoveFirst
        barra.Value = 0
        barra.Max = .RecordCount
        
    End With
    
    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            Do Until rsOrigen.EOF = True
                .AddNew
                
                .Fields("idRubros").Value = EsNulo(rsOrigen.Fields("Codigo").Value)
                .Fields("Rubro").Value = EsNulo(rsOrigen.Fields(2).Value)
                
                .Update
                
                rsOrigen.MoveNext
                barra.Value = barra.Value + 1
            Loop
        End If
        
    End With
    
    sqlOrigen = ""
    sqlDestino = ""

    If rsOrigen.State = 1 Then
        rsOrigen.Close
        Set rsOrigen = Nothing
    End If
    
If Err Then GrabarLog "MigrarRubros", Err.Number & " " & Err.Description, "Migracion"
End Sub
Private Sub MigrarLibretaPagos()
On Error Resume Next

    Dim rsOrigen As New ADODB.Recordset, rsDestino As New ADODB.Recordset
    Dim sqlOrigen As String, sqlDestino As String
    
    sqlOrigen = "SELECT * FROM Libreta_Pagos WHERE 1=1"
    sqlDestino = "SELECT * FROM Libreta_Pagos WHERE 1=2"
    
    With rsOrigen
        .CursorLocation = adUseClient
        Call .Open(sqlOrigen, connDDBBMigrar, adOpenStatic, adLockPessimistic)
        
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
        
        .MoveFirst
        barra.Value = 0
        barra.Max = .RecordCount
        
    End With
    
    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            Do Until rsOrigen.EOF = True
                .AddNew
                
                .Fields("id_FDetalle").Value = EsNulo(rsOrigen.Fields("id_FDetalle").Value)
                .Fields("id_Ctacte").Value = EsNulo(rsOrigen.Fields("id_Ctacte").Value)
                .Fields("Codigo").Value = EsNulo(rsOrigen.Fields("Codigo").Value)
                .Fields("Resta").Value = Val(rsOrigen.Fields("Resta").Value)
                .Fields("Pago").Value = Val(rsOrigen.Fields("Pago").Value)
                .Fields("Pagado").Value = EsNulo(rsOrigen.Fields("Pagado").Value)
                
                .Update
                
                rsOrigen.MoveNext
                barra.Value = barra.Value + 1
            Loop
        End If
        
    End With
    
    sqlOrigen = ""
    sqlDestino = ""

    If rsOrigen.State = 1 Then
        rsOrigen.Close
        Set rsOrigen = Nothing
    End If
    
If Err Then GrabarLog "MigrarLibretaPagos", Err.Number & " " & Err.Description, "Migracion"
End Sub
Private Sub MigrarCtaCteClientes()
On Error Resume Next

    Dim rsOrigen As New ADODB.Recordset, rsDestino As New ADODB.Recordset
    Dim sqlOrigen As String, sqlDestino As String
    
    sqlOrigen = "SELECT * FROM CuentasCorrientes WHERE 1=1"
    sqlDestino = "SELECT * FROM CuentasCorrientes WHERE 1=2"
    
    With rsOrigen
        .CursorLocation = adUseClient
        Call .Open(sqlOrigen, connDDBBMigrar, adOpenStatic, adLockPessimistic)
        
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
        
        .MoveFirst
        barra.Value = 0
        barra.Max = .RecordCount
        
    End With
    
    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            Do Until rsOrigen.EOF = True
                DoEvents
                
                .AddNew
                
                .Fields("Fecha").Value = EsNulo(Trim(rsOrigen.Fields("Fecha").Value))
                .Fields("Codigo").Value = EsNulo(Trim(rsOrigen.Fields("Codigo").Value))
                .Fields("Nombre").Value = EsNulo(Trim(rsOrigen.Fields("Nombre").Value))
                .Fields("Debito").Value = Val(Format(rsOrigen.Fields("Debito").Value, "#####0.00"))
                .Fields("Credito").Value = Val(Format(rsOrigen.Fields("Credito").Value, "#####0.00"))
                .Fields("Saldo").Value = Val(Format(rsOrigen.Fields("Saldo").Value, "#####0.00"))
                .Fields("Comentario").Value = EsNulo(rsOrigen.Fields("Comentario").Value)
                .Fields("Remito").Value = Val(rsOrigen.Fields("Remito").Value)
                .Fields("FechaInput").Value = Val(rsOrigen.Fields("FechaInput").Value)
                .Fields("AnoMes").Value = Val(rsOrigen.Fields("AnoMes").Value)
                
                .Update
                
                rsOrigen.MoveNext
                barra.Value = barra.Value + 1
            Loop
        End If
        
    End With
    
    sqlOrigen = ""
    sqlDestino = ""

    If rsOrigen.State = 1 Then
        rsOrigen.Close
        Set rsOrigen = Nothing
    End If
    
If Err Then
    GrabarLog "MigrarCtaCteClientes", Err.Number & " " & Err.Description, "Migracion"
End If
End Sub
Private Sub MigrarFDetalles()
On Error Resume Next

    Dim rsOrigen As New ADODB.Recordset, rsDestino As New ADODB.Recordset
    Dim sqlOrigen As String, sqlDestino As String
    
    sqlOrigen = "SELECT * FROM FDetalle WHERE 1=1"
    sqlDestino = "SELECT * FROM FDetalle WHERE 1=2"
    
    With rsOrigen
        .CursorLocation = adUseClient
        Call .Open(sqlOrigen, connDDBBMigrar, adOpenStatic, adLockPessimistic)
        
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
        
        .MoveFirst
        barra.Value = 0
        barra.Max = .RecordCount
        
    End With
    
    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            Do Until rsOrigen.EOF = True
                DoEvents
                
                .AddNew
                
                .Fields("Fecha").Value = EsNulo(Trim(rsOrigen.Fields("Fecha").Value))
                .Fields("Remito").Value = Val(Trim(rsOrigen.Fields("Remito").Value))
                .Fields("Codigo").Value = EsNulo(Trim(rsOrigen.Fields("Codigo").Value))
                .Fields("Cantidad").Value = Val(Format(rsOrigen.Fields("Cantidad").Value, "#####0.00"))
                .Fields("Detalle").Value = EsNulo(Trim(rsOrigen.Fields("Detalle").Value))
                .Fields("Precio").Value = Val(Format(rsOrigen.Fields("Precio").Value, "#####0.00"))
                .Fields("Total").Value = Val(Format(rsOrigen.Fields("Total").Value, "#####0.00"))
                .Fields("Total_CtaCte").Value = Val(Format(rsOrigen.Fields("Total_CtaCte").Value, "#####0.00"))
                .Fields("TIva").Value = Val(Format(rsOrigen.Fields("TIva").Value, "#####0.00"))
                .Fields("Confirmado").Value = EsNulo(rsOrigen.Fields("Confirmado").Value)
                .Fields("Pagado").Value = EsNulo(rsOrigen.Fields("Pagado").Value)
                .Fields("Pago").Value = Val(Format(rsOrigen.Fields("Pago").Value, "#####0.00"))
                .Fields("Resta").Value = Val(Format(rsOrigen.Fields("Resta").Value, "#####0.00"))
                .Fields("TotalIva").Value = Val(Format(rsOrigen.Fields("TotalIva").Value, "#####0.00"))
                .Fields("id_Ctacte").Value = EsNulo(rsOrigen.Fields("id_Ctacte").Value)

                .Update
                
                rsOrigen.MoveNext
                barra.Value = barra.Value + 1
            Loop
        End If
        
    End With
    
    sqlOrigen = ""
    sqlDestino = ""

    If rsOrigen.State = 1 Then
        rsOrigen.Close
        Set rsOrigen = Nothing
    End If
    
If Err Then
    GrabarLog "MigrarFDetalles", Err.Number & " " & Err.Description, "Migracion"
End If
End Sub
Private Sub MigrarFacturas()
On Error Resume Next

    Dim rsOrigen As New ADODB.Recordset, rsDestino As New ADODB.Recordset
    Dim sqlOrigen As String, sqlDestino As String
    
    sqlOrigen = "SELECT * FROM Factura WHERE 1=1"
    sqlDestino = "SELECT * FROM Factura WHERE 1=2"
    
    With rsOrigen
        .CursorLocation = adUseClient
        Call .Open(sqlOrigen, connDDBBMigrar, adOpenStatic, adLockPessimistic)
        
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
        
        .MoveFirst
        barra.Value = 0
        barra.Max = .RecordCount
        
    End With
    
    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            Do Until rsOrigen.EOF = True
                DoEvents
                
                .AddNew
                
                .Fields("NComprobante").Value = EsNulo(Trim(rsOrigen.Fields("NComprobante").Value))
                .Fields("Fecha").Value = Val(Trim(rsOrigen.Fields("Fecha").Value))
                .Fields("Codigo").Value = EsNulo(Trim(rsOrigen.Fields("Codigo").Value))
                .Fields("Nombre").Value = EsNulo(Trim(rsOrigen.Fields("Nombre").Value))
                .Fields("Domicilio").Value = EsNulo(Trim(rsOrigen.Fields("Domicilio").Value))
                .Fields("Localidad").Value = EsNulo(rsOrigen.Fields("Localidad").Value)
                .Fields("Telefono").Value = EsNulo(rsOrigen.Fields("Telefono").Value)
                .Fields("Iva").Value = EsNulo(rsOrigen.Fields("Iva").Value)
                .Fields("CVenta").Value = EsNulo(rsOrigen.Fields("CVenta").Value)
                .Fields("Remito").Value = EsNulo(rsOrigen.Fields("Remito").Value)
                .Fields("SubTotal").Value = Val(Format(rsOrigen.Fields("SubTotal").Value, "#####0.00"))
                .Fields("Total").Value = Val(Format(rsOrigen.Fields("Total").Value, "#####0.00"))
                .Fields("Total_CtaCte").Value = Val(Format(rsOrigen.Fields("Total_CtaCte").Value, "#####0.00"))
                .Fields("Comentario").Value = EsNulo(rsOrigen.Fields("Comentario").Value)
                .Fields("Cuit").Value = EsNulo(rsOrigen.Fields("Cuit").Value)
                .Fields("Tipo").Value = EsNulo(rsOrigen.Fields("Tipo").Value)

                .Update
                
                rsOrigen.MoveNext
                barra.Value = barra.Value + 1
            Loop
        End If
        
    End With
    
    sqlOrigen = ""
    sqlDestino = ""

    If rsOrigen.State = 1 Then
        rsOrigen.Close
        Set rsOrigen = Nothing
    End If
    
If Err Then GrabarLog "MigrarFacturas", Err.Number & " " & Err.Description, "Migracion"
End Sub
Private Sub MigrarArticulos()
On Error Resume Next

    Dim rsOrigen As New ADODB.Recordset, rsDestino As New ADODB.Recordset
    Dim sqlOrigen As String, sqlDestino As String
    
    sqlOrigen = "SELECT * FROM Articulos WHERE 1=1"
    sqlDestino = "SELECT * FROM Articulos WHERE 1=2"
    
    With rsOrigen
        .CursorLocation = adUseClient
        Call .Open(sqlOrigen, connDDBBMigrar, adOpenStatic, adLockPessimistic)
        
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
        
        .MoveFirst
        barra.Value = 0
        barra.Max = .RecordCount
        
    End With
    
    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            Do Until rsOrigen.EOF = True
                DoEvents
                .AddNew
                
                .Fields("Codigo").Value = EsNulo(Trim(rsOrigen.Fields("Codigo").Value))
                .Fields("CodigoNum").Value = Val(Trim(rsOrigen.Fields("Codigo").Value))
                .Fields("Descrip").Value = EsNulo(Trim(rsOrigen.Fields("Descrip").Value))
                
                If IsNull(rsOrigen.Fields("Rubro").Value) = True Or Trim(rsOrigen.Fields("Rubro").Value) = "" Then
                    .Fields("idRubros").Value = "001"
                Else
                    .Fields("idRubros").Value = rsOrigen.Fields("Rubro").Value
                End If
                .Fields("idPorcentajeIva").Value = "001"
                .Fields("PCosto").Value = EsNulo(rsOrigen.Fields("PCosto").Value)
                .Fields("PVenta1").Value = EsNulo(rsOrigen.Fields("PVenta1").Value)
                .Fields("PVenta2").Value = EsNulo(rsOrigen.Fields("PVenta2").Value)
                .Fields("PVenta3").Value = EsNulo(rsOrigen.Fields("PVenta3").Value)
    
                .Fields("FechaAlta").Value = Date
                .Fields("Stock").Value = Val(Format(rsOrigen.Fields("Stock").Value, "#####0.00"))

                .Update
                
                Call GuardarEnStock("Articulo-Nuevo", Trim(rsOrigen.Fields("Codigo").Value), Date, Val(Format(rsOrigen.Fields("Stock").Value, "#####0.00")), "Stock Inicial", 0, 0)
                
                rsOrigen.MoveNext
                barra.Value = barra.Value + 1
            Loop
        End If
        
    End With
    
    sqlOrigen = ""
    sqlDestino = ""

    If rsOrigen.State = 1 Then
        rsOrigen.Close
        Set rsOrigen = Nothing
    End If
    
If Err Then GrabarLog "MigrarArticulos", Err.Number & " " & Err.Description, "Migracion"
End Sub
Private Sub MigrarClientes()
On Error Resume Next

    Dim rsOrigen As New ADODB.Recordset, rsDestino As New ADODB.Recordset
    Dim sqlOrigen As String, sqlDestino As String
    Dim vlocalidad As String, vcp As String, vProvincia As String
    Dim vTipoIva As String, vIdTipoIva As String
    
    sqlOrigen = "SELECT * FROM Clientes WHERE 1=1"
    sqlDestino = "SELECT * FROM Clientes WHERE 1=2"
    
    With rsOrigen
        .CursorLocation = adUseClient
        Call .Open(sqlOrigen, connDDBBMigrar, adOpenStatic, adLockPessimistic)
        
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
        
        .MoveFirst
        barra.Value = 0
        barra.Max = .RecordCount
        
    End With
    
    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            Do Until rsOrigen.EOF = True
                .AddNew
                
        
                .Fields("Codigo").Value = EsNulo(rsOrigen.Fields("Codigo").Value)
                .Fields("CodigoNum").Value = Val(rsOrigen.Fields("Codigo").Value)
                .Fields("Nombre").Value = EsNulo(rsOrigen.Fields("Nombre").Value)
                .Fields("RazonSocial").Value = EsNulo(rsOrigen.Fields("Nombre").Value)
                .Fields("Direccion").Value = EsNulo(rsOrigen.Fields("Direccion").Value)
                .Fields("Localidad").Value = EsNulo(rsOrigen.Fields("Localidad").Value)
                .Fields("Telefono").Value = EsNulo(rsOrigen.Fields("Telefono").Value)
                
                vIdTipoIva = TraerDato("TipoIva", "TipoIva = '" & rsOrigen.Fields("Iva").Value & "'", "idTipoIva")
                
                .Fields("idTipoIva").Value = vIdTipoIva
                .Fields("Cuit").Value = EsNulo(rsOrigen.Fields("Cuit").Value)
                .Fields("idTipoCliente").Value = "001"
                .Fields("idActividad").Value = "001"
                .Fields("idListas").Value = "001"
                
                .Fields("Fecha_Alta").Value = EsNulo(rsOrigen.Fields("Fecha_Alta").Value)
                .Fields("Fecha_Baja").Value = EsNulo(rsOrigen.Fields("Fecha_Baja").Value)
                
                If EsNulo(rsOrigen.Fields("Pasivo").Value) = "NO" Then
                    .Fields("idEstados").Value = "001"
                Else
                    .Fields("idEstados").Value = "003"
                End If
            
                .Fields("Observaciones").Value = EsNulo(rsOrigen.Fields("Comentario").Value)
            
                .Fields("U_Venta").Value = EsNulo(rsOrigen.Fields("U_Venta").Value)
                .Fields("U_Pago").Value = EsNulo(rsOrigen.Fields("U_Pago").Value)
                .Fields("Saldo").Value = EsNulo(rsOrigen.Fields("Saldo").Value)
                
                .Fields("E-Mail").Value = EsNulo(rsOrigen.Fields("EMail").Value)
                .Update
                
                rsOrigen.MoveNext
                barra.Value = barra.Value + 1
            Loop
        End If
        
    End With
    
    sqlOrigen = ""
    sqlDestino = ""

    If rsOrigen.State = 1 Then
        rsOrigen.Close
        Set rsOrigen = Nothing
    End If
    
If Err Then GrabarLog "MigrarClientes", Err.Number & " " & Err.Description, "Migracion"
End Sub
Private Sub MigrarProveedores()
On Error Resume Next

    Dim rsOrigen As New ADODB.Recordset, rsDestino As New ADODB.Recordset
    Dim sqlOrigen As String, sqlDestino As String
    Dim vlocalidad As String, vcp As String, vProvincia As String
    Dim vTipoIva As String, vIdTipoIva As String
    
    sqlOrigen = "SELECT * FROM Proveedores WHERE 1=1"
    sqlDestino = "SELECT * FROM Proveedores WHERE 1=2"
    
    With rsOrigen
        .CursorLocation = adUseClient
        Call .Open(sqlOrigen, connDDBBMigrar, adOpenStatic, adLockPessimistic)
        
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
        
        .MoveFirst
        barra.Value = 0
        barra.Max = .RecordCount
        
    End With
    
    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            Do Until rsOrigen.EOF = True
                .AddNew
                
        
                .Fields("Codigo").Value = EsNulo(rsOrigen.Fields("Codigo").Value)
                .Fields("CodigoNum").Value = Val(rsOrigen.Fields("Codigo").Value)
                .Fields("Nombre").Value = EsNulo(rsOrigen.Fields("Nombre").Value)
                .Fields("RazonSocial").Value = EsNulo(rsOrigen.Fields("Nombre").Value)
                .Fields("Direccion").Value = EsNulo(rsOrigen.Fields("Direccion").Value)
                .Fields("Localidad").Value = EsNulo(rsOrigen.Fields("Localidad").Value)
                .Fields("Telefono").Value = EsNulo(rsOrigen.Fields("Telefono").Value)
                
                vIdTipoIva = TraerDato("TipoIva", "TipoIva = '" & rsOrigen.Fields("Iva").Value & "'", "idTipoIva")
                
                .Fields("idTipoIva").Value = "001" 'vIdTipoIva
                .Fields("Cuit").Value = EsNulo(rsOrigen.Fields("Cuit").Value)
                .Fields("idTipoCliente").Value = "001"
                .Fields("idActividad").Value = "001"
                .Fields("idListas").Value = "001"
                
                .Fields("Fecha_Alta").Value = Date
                '.Fields("Fecha_Baja").Value = 'EsNulo(rsOrigen.Fields("Fecha_Baja").Value)
                

                .Fields("idEstados").Value = "001"
            
                .Fields("Observaciones").Value = ""
                
                .Fields("E-Mail").Value = ""
                .Update
                
                rsOrigen.MoveNext
                barra.Value = barra.Value + 1
            Loop
        End If
        
    End With
    
    sqlOrigen = ""
    sqlDestino = ""

    If rsOrigen.State = 1 Then
        rsOrigen.Close
        Set rsOrigen = Nothing
    End If
    
If Err Then GrabarLog "MigrarProveedores", Err.Number & " " & Err.Description, "Migracion"
End Sub
Private Sub MigrarCuentasSA(vtabla As String)
On Error Resume Next
    
    vtabla = Replace(vtabla, ".DBF", "")

    Dim rsOrigen As New ADODB.Recordset, rsDestino As New ADODB.Recordset
    Dim sqlOrigen As String, sqlDestino As String
    
    sqlOrigen = "SELECT * FROM " & vtabla & " WHERE (Empresa = '1')"
    sqlDestino = "SELECT * FROM Cuentas WHERE 1=2"
    
    With rsOrigen
        .CursorLocation = adUseClient
        Call .Open(sqlOrigen, connDDBBMigrar, adOpenStatic, adLockPessimistic)
        
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
        
        .MoveFirst
        barra.Value = 0
        barra.Max = .RecordCount
        
    End With

    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            Do Until rsOrigen.EOF = True
                .AddNew

                .Fields("CodigoCuenta").Value = EsNulo(rsOrigen.Fields("Cuenta").Value)
                .Fields("Cuenta").Value = EsNulo(rsOrigen.Fields("Descrip").Value)
                .Fields("Imputable").Value = EsNulo(rsOrigen.Fields("Imp_Asto").Value)
                
                .Update
                
                rsOrigen.MoveNext
                barra.Value = barra.Value + 1
            Loop
        End If
        
    End With
    
    sqlOrigen = ""
    sqlDestino = ""

    If rsOrigen.State = 1 Then
        rsOrigen.Close
        Set rsOrigen = Nothing
    End If

If Err Then GrabarLog "MigrarCuentasSA", Err.Number & " " & Err.Description, "BasSisAgro"
End Sub
Private Sub MigrarProveedoresSA(vtabla As String)
On Error Resume Next
    
    vtabla = Replace(vtabla, ".DBF", "")

    Dim rsOrigen As New ADODB.Recordset, rsDestino As New ADODB.Recordset
    Dim sqlOrigen As String, sqlDestino As String
    Dim vproveedor As String
    
    sqlOrigen = "SELECT * FROM " & vtabla & ""
    sqlDestino = "SELECT * FROM Proveedores WHERE 1=2"
    
    With rsOrigen
        .CursorLocation = adUseClient
        Call .Open(sqlOrigen, connDDBBMigrar, adOpenStatic, adLockPessimistic)
        
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
        
        .MoveFirst
        barra.Value = 0
        barra.Max = .RecordCount
        
    End With
    
    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            Do Until rsOrigen.EOF = True
            
                vproveedor = EsNulo(rsOrigen.Fields("Proveedor").Value)
                
                If Not vproveedor = "" Then
                
                    .AddNew
                
                
                    .Fields("Codigo").Value = EsNulo(rsOrigen.Fields("Proveedor").Value)
                    .Fields("Codigo_Num").Value = Val(rsOrigen.Fields("Razon_Soc").Value)
                    .Fields("Nombre").Value = EsNulo(rsOrigen.Fields("Razon_Soc").Value)
                    .Fields("RazonSocial").Value = EsNulo(rsOrigen.Fields("Razon_Soc").Value)
                    .Fields("Direccion").Value = EsNulo(rsOrigen.Fields("Domicilio").Value)
                    .Fields("Localidad").Value = EsNulo(rsOrigen.Fields("Localidad").Value)
                    .Fields("CodigoPostal").Value = EsNulo(rsOrigen.Fields("Cod_Postal").Value)
                
                    Select Case EsNulo(rsOrigen.Fields("Provincia").Value)
                    
                        Case "B"
                            .Fields("Provincia").Value = "BUENOS AIRES"
                        Case "C"
                            .Fields("Provincia").Value = "CATAMARCA"
                        Case "CF"
                            .Fields("Provincia").Value = "CAPITAL FEDERAL"
                        Case "CO"
                            .Fields("Provincia").Value = "CORRIENTES"
                        Case "ER"
                            .Fields("Provincia").Value = "ENTRE RIOS"
                        Case "F"
                            .Fields("Provincia").Value = "FORMOSA"
                        Case "H"
                            .Fields("Provincia").Value = "CHACO"
                        Case "J"
                            .Fields("Provincia").Value = "JUJUY"
                        Case "L"
                            .Fields("Provincia").Value = "SAN LUIS"
                        Case "M"
                            .Fields("Provincia").Value = "MENDOZA"
                        Case "MS"
                            .Fields("Provincia").Value = "MISIONES"
                        Case "P"
                            .Fields("Provincia").Value = "LA PAMPA"
                        Case "Q"
                            .Fields("Provincia").Value = "NEUQUEN"
                        Case "R"
                            .Fields("Provincia").Value = "RIO NEGRO"
                        Case "RJ"
                            .Fields("Provincia").Value = "LA RIOJA"
                        Case "T"
                            .Fields("Provincia").Value = "TUCUMAN"
                        Case "S"
                            .Fields("Provincia").Value = "SANTIAGO DEL ESTERO"
                        Case "SA"
                            .Fields("Provincia").Value = "SALTA"
                        Case "SE"
                            .Fields("Provincia").Value = "SANTIAGO DEL ESTERO"
                        Case "SJ"
                            .Fields("Provincia").Value = "SAN JUAN"
                        Case "X"
                            .Fields("Provincia").Value = "CORDOBA"
                        Case Else
                            Debug.Print ("OJO........  " & EsNulo(rsOrigen.Fields("Provincia").Value))
                            
                    End Select
                
                    .Fields("Telefono").Value = EsNulo(rsOrigen.Fields("telefono1").Value)
                    .Fields("Fax").Value = EsNulo(rsOrigen.Fields("telefono2").Value)
                    .Fields("Celular").Value = EsNulo(rsOrigen.Fields("telefono3").Value)
                    .Fields("Cuit").Value = Replace(EsNulo(rsOrigen.Fields("Cuit").Value), " ", "-")
                    
                    .Fields("idTipoIva").Value = "0" & EsNulo(rsOrigen.Fields("Cod_Impues").Value)
                
                    .Fields("Saldo").Value = Val(EsNulo(rsOrigen.Fields("Saldo_R").Value))
                    
                    .Update
                Else


                End If
                
                rsOrigen.MoveNext
                barra.Value = barra.Value + 1
            Loop
        End If
        
    End With
    
    sqlOrigen = ""
    sqlDestino = ""

    If rsOrigen.State = 1 Then
        rsOrigen.Close
        Set rsOrigen = Nothing
    End If

If Err Then GrabarLog "MigrarProveedoresSisAgro", Err.Number & " " & Err.Description, "BasSisAgro"
End Sub
Private Sub MigrarCuentasCorrientesCSA(vtabla As String)
On Error Resume Next
    
    vtabla = Replace(vtabla, ".DBF", "")

    Dim rsOrigen As New ADODB.Recordset, rsDestino As New ADODB.Recordset
    Dim sqlOrigen As String, sqlDestino As String
    Dim vcliente As String
    Dim vMedioPago As String
    
    sqlOrigen = "SELECT * FROM " & vtabla & " WHERE (Empresa = '1')"
    sqlDestino = "SELECT * FROM cuentasCorrientes WHERE 1=2"
    
    With rsOrigen
        .CursorLocation = adUseClient
        Call .Open(sqlOrigen, connDDBBMigrar, adOpenStatic, adLockPessimistic)
        
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
        
        .MoveFirst
        barra.Value = 0
        barra.Max = .RecordCount
        
    End With

    vnroremito = 0
    
    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            Do Until rsOrigen.EOF = True
                
                vMedioPago = ""
                DoEvents

                vcliente = EsNulo(rsOrigen.Fields("Cliente").Value)
                
                If Not vcliente = "" Then
                
                    .AddNew
                
                    .Fields("Fecha").Value = strfechaMySQL(rsOrigen.Fields("Fec_Emis").Value)
                    .Fields("Codigo").Value = EsNulo(rsOrigen.Fields("Cliente").Value)
                    .Fields("Nombre").Value = TraerDato("Clientes", "Codigo = '" & vcliente & "'", "Nombre")
                    
                    .Fields("Credito").Value = 0
                    .Fields("Debito").Value = 0
                                                

                    
                    Select Case EsNulo(rsOrigen.Fields("Movimiento").Value)
                        
                        Case "FC"
                            .Fields("Debito").Value = Val(Format(rsOrigen.Fields("Imp_Neto").Value, "######0.00"))
                            
                            vnroremito = MigrarFacturaSA(rsOrigen)
                            .Fields("Remito").Value = Val(vnroremito)
                            
                            
                        Case "RC"
                            'Hago un Pago
                            .Fields("Credito").Value = Val(Format(rsOrigen.Fields("Imp_Neto").Value, "######0.00")) * (-1)
                            
                            'vMedioPago = TraerDato("AVCH00", "Nro_Inter = " & EsNulo(rsOrigen.Fields("Nro_Inter").value) & "", "Tipo_Valor", pathDBMigrar("ARV"))

                            'Select Case vMedioPago
                                'Case "EF"
                                '    .Fields("idMedioPago").value = 1
                                'Case "CH"
                                '    .Fields("idMedioPago").value = 4
                                'Case ""
                                '    MsgBox "ERROR"
                                'Case Else
                                '    MsgBox "ERROR"
                            'End Select
                        Case "AC"
                            .Fields("Credito").Value = Val(Format(rsOrigen.Fields("Imp_Neto").Value, "######0.00")) * (-1)
                        
                            'vMedioPago = TraerDato("AVCH00", "Nro_Inter = " & EsNulo(rsOrigen.Fields("Nro_Inter").value) & "", "Tipo_Valor", pathDBMigrar("ARV"))
                            
                            'Select Case vMedioPago
                            '    Case "EF"
                            '        .Fields("idMedioPago").value = 1
                            '    Case "CH"
                            '        .Fields("idMedioPago").value = 4
                            '    Case ""
                            '        MsgBox "ERROR"
                            '    Case Else
                            '        MsgBox "ERROR"
                            'End Select
                        Case "AD"
                            .Fields("Debito").Value = Val(Format(rsOrigen.Fields("Imp_Neto").Value, "######0.00"))
                        
                        Case "CC"
                            'Hace un Pago de una Factura de Contado
                            .Fields("Credito").Value = Val(Format(rsOrigen.Fields("Imp_Neto").Value, "######0.00")) * (-1)
                            
                            .Fields("Remito").Value = TraerDato("CuentasCorrientes", "NroInterno = " & Val(rsOrigen.Fields("Nro_Inter").Value) & "", "Remito")
                            
                            .Fields("idMedioPago").Value = 99
                            'vMedioPago = TraerDato("AVCH00", "Nro_Inter = " & EsNulo(rsOrigen.Fields("Nro_Inter").value) & "", "Tipo_Valor", pathDBMigrar("ARV"))

                            'Select Case vMedioPago
                            '    Case "EF"
                            '        .Fields("idMedioPago").value = 1
                            '    Case "CH"
                            '        .Fields("idMedioPago").value = 4
                            '    Case ""
                            '        MsgBox "ERROR"
                            '    Case Else
                            '       MsgBox "ERROR"
                            'End Select
                        Case "CD"
                            'Hace un Debito, por una factura que se paga al Contado
                            .Fields("Debito").Value = Val(Format(rsOrigen.Fields("Imp_Neto").Value, "######0.00"))
                            
                            vnroremito = MigrarFacturaSA(rsOrigen)
                            .Fields("Remito").Value = Val(vnroremito)
                        
                        Case "NC"
                            .Fields("Credito").Value = Val(Format(rsOrigen.Fields("Imp_Neto").Value, "######0.00")) * (-1)
                            
                            vnroremito = MigrarFacturaSA(rsOrigen)
                            .Fields("Remito").Value = Val(vnroremito)
                            '.Fields("idMedioPago").value = 8
                        
                            
                        Case "ND"
                            .Fields("Debito").Value = Val(Format(rsOrigen.Fields("Imp_Neto").Value, "######0.00"))
                            
                        Case "RG", "IB", "LV", "SU", "RV"
                            .Fields("Credito").Value = Val(Format(rsOrigen.Fields("Imp_Neto").Value, "######0.00")) * (-1)

                            'vMedioPago = TraerDato("AVCH00", "Nro_Inter = " & EsNulo(rsOrigen.Fields("Nro_Inter").value) & "", "Tipo_Valor", pathDBMigrar("ARV"))

                            vnroremito = MigrarFacturaSA(rsOrigen)
                            .Fields("Remito").Value = Val(vnroremito)
                            '.Fields("idMedioPago").value = 99
    
                        
                        Case "RI", "SS"
                            'MsgBox "Tipo de Movimiento: " & EsNulo(rsOrigen.Fields("Movimiento").Value)
                            
                            .Fields("Debito").Value = Val(Format(rsOrigen.Fields("Imp_Neto").Value, "######0.00"))
                            
                        Case "SI"
                            'Si es Positivo es un DEBITO, si es negativo es un Credito
                            
                            If Val(rsOrigen.Fields("Imp_Neto").Value) > 0 Then
                                .Fields("Debito").Value = Val(Format(rsOrigen.Fields("Imp_Neto").Value, "######0.00"))
                            
                            Else
                                .Fields("Credito").Value = Val(Format(rsOrigen.Fields("Imp_Neto").Value, "######0.00")) * (-1)
                            
                                'vMedioPago = TraerDato("AVCH00", "Nro_Inter = " & EsNulo(rsOrigen.Fields("Nro_Inter").value) & "", "Tipo_Valor", pathDBMigrar("ARV"))

                                'Select Case vMedioPago
                                '    Case "EF"
                                '        .Fields("idMedioPago").value = 1
                                '    Case "CH"
                                '        .Fields("idMedioPago").value = 4
                                '    Case ""
                                '        MsgBox "ERROR"
                                '    Case Else
                                '        MsgBox "ERROR"
                                'End Select
                            
                            End If
                        
                        Case Else
                             
                            MsgBox EsNulo(rsOrigen.Fields("Movimiento").Value)
                            
                    End Select
                
                    
                    .Fields("Comentario").Value = EsNulo(rsOrigen.Fields("Leyenda").Value)
            
                    .Fields("FechaVencimiento").Value = EsNulo(rsOrigen.Fields("Fec_vto").Value)
                    .Fields("FechaIngreso").Value = EsNulo(rsOrigen.Fields("Fec_Ingr").Value)
                
                    .Fields("NroInterno").Value = EsNulo(rsOrigen.Fields("Nro_Inter").Value)
                    .Fields("TipoMovimiento").Value = EsNulo(rsOrigen.Fields("Movimiento").Value)
                        
                    .Fields("NroAsiento").Value = EsNulo(rsOrigen.Fields("Nro_Asto").Value)
                    .Update

                End If
                
                barra.Value = barra.Value + 1
                rsOrigen.MoveNext
            Loop
        End If
        
    End With
    
    sqlOrigen = ""
    sqlDestino = ""

    If rsOrigen.State = 1 Then
        rsOrigen.Close
        Set rsOrigen = Nothing
    End If

If Err Then GrabarLog "MigrarCuentasCorrientesCSA", Err.Number & " " & Err.Description, "BasSisAgro"
End Sub
Private Sub MigrarCuentasCorrientesPSA(vtabla As String)
On Error Resume Next
    
    vtabla = Replace(vtabla, ".DBF", "")

    Dim rsOrigen As New ADODB.Recordset, rsDestino As New ADODB.Recordset
    Dim sqlOrigen As String, sqlDestino As String
    Dim vproveedor As String
    
    sqlOrigen = "SELECT * FROM " & vtabla & " WHERE (Empresa = '1')  ORDER BY Nro_Inter ASC"  ', Letra ASC"
    sqlDestino = "SELECT * FROM pcuentascorrientes WHERE 1=2"
    
    With rsOrigen
        .CursorLocation = adUseClient
        Call .Open(sqlOrigen, connDDBBMigrar, adOpenStatic, adLockPessimistic)
        
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
        
        .MoveFirst
        barra.Value = 0
        barra.Max = .RecordCount
    End With
                
    vnroremito = 0
                
    With rsDestino
        .CursorLocation = adUseClient
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            Do Until rsOrigen.EOF = True
                DoEvents

                
                vproveedor = EsNulo(rsOrigen.Fields("Proveedor").Value)
                
                If Not vproveedor = "" Then
                    
                    .AddNew
                
                    .Fields("Fecha").Value = strfechaMySQL(rsOrigen.Fields("Fec_Emis").Value)
                    .Fields("Codigo").Value = EsNulo(rsOrigen.Fields("Proveedor").Value)
                    .Fields("Nombre").Value = TraerDato("Proveedores", "Codigo = '" & vproveedor & "'", "Nombre")
                    
                    .Fields("Debito").Value = 0
                    .Fields("Credito").Value = 0
                    
                    Select Case EsNulo(rsOrigen.Fields("Movimiento").Value)
                    
                    
                        Case "FC"
                            vnroremito = MigrarPFacturaSA(rsOrigen)
                            
                            .Fields("Debito").Value = Val(Format(rsOrigen.Fields("Imp_Neto").Value, "######0.00"))
                            .Fields("Remito").Value = Val(vnroremito)
                        
                        Case "RC"
                            'Nota: No lleva Factura - Si Remito
                            .Fields("Credito").Value = Val(Format(rsOrigen.Fields("Imp_Neto").Value, "######0.00")) * (-1)
                            
                            'Me hago el langa y lo hago desde acca....

                        Case "AC"
                            'Nota: Estudiar si lleva Doc
                            .Fields("Credito").Value = Val(Format(rsOrigen.Fields("Imp_Neto").Value, "######0.00")) * (-1)
                        
                        Case "AD"
                            'Nota: Estudiar si lleva Doc
                            .Fields("Debito").Value = Val(Format(rsOrigen.Fields("Imp_Neto").Value, "######0.00"))
                            '.Fields("Remito").value = Val(vNroRemito)

                        Case "CC"
                            'Nota:
                            .Fields("Credito").Value = Val(Format(rsOrigen.Fields("Imp_Neto").Value, "######0.00")) * (-1)

                        Case "CD"
                            'Nota: Genera una factura y la paga con CC
                            
                                                
                            vnroremito = MigrarPFacturaSA(rsOrigen)
                
                            .Fields("Debito").Value = Val(Format(rsOrigen.Fields("Imp_Neto").Value, "######0.00"))
                            .Fields("Remito").Value = Val(vnroremito)
                            
                        Case "NC"
                                                
                            vnroremito = MigrarPFacturaSA(rsOrigen)

                            .Fields("Credito").Value = Val(Format(rsOrigen.Fields("Imp_Neto").Value, "######0.00")) * (-1)
                            .Fields("idMedioPago").Value = 8
                            
                            'No lo cargo al remito, porque pongo el del Documento que voy a pagar
                            '.Fields("Remito").value = Val(vNroRemito)
                        
                        
                        Case "ND"
                            .Fields("Debito").Value = Val(Format(rsOrigen.Fields("Imp_Neto").Value, "######0.00"))
                            
                        Case "RI", "SS"
                            MsgBox "Tipo de Movimiento: " & EsNulo(rsOrigen.Fields("Movimiento").Value)
                            
                        Case "SI"
                            'Si es Positivo es un DEBITO, si es negativo es un Credito
                            
                            If Val(rsOrigen.Fields("Imp_Neto").Value) > 0 Then
                                .Fields("Debito").Value = Val(Format(rsOrigen.Fields("Imp_Neto").Value, "######0.00"))
                            Else
                                .Fields("Credito").Value = Val(Format(rsOrigen.Fields("Imp_Neto").Value, "######0.00")) * (-1)
                            End If
                            
                            
                        Case Else
                                MsgBox EsNulo(rsOrigen.Fields("Movimiento").Value)
                            
                            
                            
                    End Select
                
                    .Fields("Comentario").Value = EsNulo(rsOrigen.Fields("Leyenda").Value)
            
                    .Fields("FechaVencimiento").Value = EsNulo(rsOrigen.Fields("Fec_vto").Value)
                    .Fields("FechaIngreso").Value = EsNulo(rsOrigen.Fields("Fec_Ingr").Value)

                    .Fields("NroInterno").Value = EsNulo(rsOrigen.Fields("Nro_Inter").Value)
                    .Fields("TipoMovimiento").Value = EsNulo(rsOrigen.Fields("Movimiento").Value)
                        
                    .Update

                
                End If
                barra.Value = barra.Value + 1
                rsOrigen.MoveNext
            Loop
        End If
        
    End With
    
    sqlOrigen = ""
    sqlDestino = ""

    If rsOrigen.State = 1 Then
        rsOrigen.Close
        Set rsOrigen = Nothing
    End If

If Err Then GrabarLog "MigrarCuentasCorrientesPSA", Err.Number & " " & Err.Description, "BasSisAgro"
End Sub
Private Function MigrarPFacturaSA(rsFacturaOrigen As ADODB.Recordset) As Long
On Error Resume Next
    
    Dim rsPFactura As New ADODB.Recordset, sqlPFactura As String
    
    sqlPFactura = "SELECT * FROM PFactura WHERE 1=2"
    
    vnroremito = vnroremito + 1

    With rsPFactura
        Call .Open(sqlPFactura, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
            If (.EOF = True) Then .AddNew
        
            .Fields("Remito").Value = vnroremito
            .Fields("Sucursal").Value = FormatoUltimoCodigo(4, EsNulo(rsFacturaOrigen.Fields("Sucursal").Value))
            .Fields("NComprobante").Value = FormatoUltimoCodigo(8, EsNulo(rsFacturaOrigen.Fields("Nro_Exter").Value))
            
            Select Case EsNulo(rsFacturaOrigen.Fields("Movimiento").Value)
                
                Case "FC"
                    .Fields("Tipo").Value = "Fact A"
                
                Case "NC"
                    .Fields("Tipo").Value = "Nota C"
                
                Case "ND"
                    .Fields("Tipo").Value = "Nota D"
            
            End Select
            
            .Fields("Letra").Value = EsNulo(rsFacturaOrigen.Fields("Letra").Value)
            .Fields("Fecha").Value = strfechaMySQL(rsFacturaOrigen.Fields("Fec_Emis").Value)
            .Fields("Hora").Value = Time
            
            .Fields("Codigo").Value = EsNulo(rsFacturaOrigen.Fields("Proveedor").Value)
            .Fields("Nombre").Value = TraerDato("Proveedores", "Codigo = '" & EsNulo(rsFacturaOrigen.Fields("Proveedor").Value) & "'", "Nombre")
            .Fields("Domicilio").Value = TraerDato("Proveedores", "Codigo = '" & EsNulo(rsFacturaOrigen.Fields("Proveedor").Value) & "'", "Direccion")
            .Fields("Localidad").Value = TraerDato("Proveedores", "Codigo = '" & EsNulo(rsFacturaOrigen.Fields("Proveedor").Value) & "'", "Localidad")
            .Fields("Telefono").Value = TraerDato("Proveedores", "Codigo = '" & EsNulo(rsFacturaOrigen.Fields("Proveedor").Value) & "'", "Telefono")
            .Fields("Cuit").Value = TraerDato("Proveedores", "Codigo = '" & EsNulo(rsFacturaOrigen.Fields("Proveedor").Value) & "'", "Cuit")
            
            .Fields("SubTotal").Value = Val(rsFacturaOrigen.Fields("Imp_Grava").Value)
            .Fields("Descuento").Value = 0
            .Fields("Total").Value = Val(rsFacturaOrigen.Fields("Imp_Neto").Value)

            .Fields("Comentario").Value = EsNulo(rsFacturaOrigen.Fields("Leyenda").Value)
            .Fields("TipoMovimiento").Value = EsNulo(rsFacturaOrigen.Fields("Movimiento").Value)
            
            .Fields("FVencimiento").Value = strfechaMySQL(rsFacturaOrigen.Fields("Fec_Vto").Value)
            .Fields("NroInterno").Value = EsNulo(rsFacturaOrigen.Fields("Nro_Inter").Value)
            .Fields("NroAsiento").Value = EsNulo(rsFacturaOrigen.Fields("Nro_Asto").Value)
            
            .Update
            
            Select Case EsNulo(rsFacturaOrigen.Fields("Porc_Iva").Value)
            
                Case 10.5
                    Call EjecutarScript("INSERT INTO IvaFacturaCompra (remito, Iva105, Retenciones, Percepciones, NoGravado, ITC) VALUES (" & vnroremito & "," & Val(rsFacturaOrigen.Fields("Imp_Iva").Value) & "," & Val(rsFacturaOrigen.Fields("Imp_Ret").Value) & "," & Val(rsFacturaOrigen.Fields("Imp_Per").Value) & "," & Val(rsFacturaOrigen.Fields("Imp_No_Gr").Value) & "," & Val(rsFacturaOrigen.Fields("Imp_Grava").Value) & ");")
                
                Case 21
                    Call EjecutarScript("INSERT INTO IvaFacturaCompra (remito, Iva210, Retenciones, Percepciones, NoGravado, ITC) VALUES (" & vnroremito & "," & Val(rsFacturaOrigen.Fields("Imp_Iva").Value) & "," & Val(rsFacturaOrigen.Fields("Imp_Ret").Value) & "," & Val(rsFacturaOrigen.Fields("Imp_Per").Value) & "," & Val(rsFacturaOrigen.Fields("Imp_No_Gr").Value) & "," & Val(rsFacturaOrigen.Fields("Imp_Grava").Value) & ");")
                
                Case 27
                    Call EjecutarScript("INSERT INTO IvaFacturaCompra (remito, Iva270, Retenciones, Percepciones, NoGravado, ITC) VALUES (" & vnroremito & "," & Val(rsFacturaOrigen.Fields("Imp_Iva").Value) & "," & Val(rsFacturaOrigen.Fields("Imp_Ret").Value) & "," & Val(rsFacturaOrigen.Fields("Imp_Per").Value) & "," & Val(rsFacturaOrigen.Fields("Imp_No_Gr").Value) & "," & Val(rsFacturaOrigen.Fields("Imp_Grava").Value) & ");")
            
            End Select
            
        End If
    End With
    
    If rsPFactura.State = 1 Then
        rsPFactura.Close
        Set rsPFactura = Nothing
    End If
    
    MigrarPFacturaSA = vnroremito
    
If Err Then GrabarLog "MigrarPFacturaSA", Err.Number & " " & Err.Description, Me.Caption
End Function
Private Function MigrarFacturaSA(rsFacturaCliente As ADODB.Recordset) As Long
On Error Resume Next
    
    Dim rsFactura As New ADODB.Recordset, sqlFactura As String
    
    sqlFactura = "SELECT * FROM Factura WHERE 1=2"
    
    vnroremito = vnroremito + 1

    With rsFactura
        Call .Open(sqlFactura, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
            If (.EOF = True) Then .AddNew
        
            .Fields("Remito").Value = vnroremito
            '.Fields("Sucursal").Value = FormatoUltimoCodigo(4, EsNulo(rsFacturaOrigen.Fields("Sucursal").Value))
            .Fields("NComprobante").Value = FormatoUltimoCodigo(8, EsNulo(rsFacturaCliente.Fields("Nro_Exter").Value))
            
            Select Case EsNulo(rsFacturaCliente.Fields("Movimiento").Value)
                
                Case "FC", "CD"
                    .Fields("Tipo").Value = "Fact A"
                
                Case "NC"
                    .Fields("Tipo").Value = "Nota C"
                
                Case "ND"
                    .Fields("Tipo").Value = "Nota D"
                
                Case Else
                    .Fields("Tipo").Value = "Otros"
                
                                
            End Select
            
            .Fields("Letra").Value = EsNulo(rsFacturaCliente.Fields("Letra").Value)
            .Fields("PuntoDeVenta").Value = EsNulo(rsFacturaCliente.Fields("Sucursal").Value)
            .Fields("Fecha").Value = strfechaMySQL(rsFacturaCliente.Fields("Fec_Emis").Value)
            .Fields("FechaVencimiento").Value = strfechaMySQL(rsFacturaCliente.Fields("Fec_VTO").Value)
            .Fields("Hora").Value = Time
            
            .Fields("Codigo").Value = EsNulo(rsFacturaCliente.Fields("Cliente").Value)
            .Fields("Nombre").Value = TraerDato("Clientes", "Codigo = '" & EsNulo(rsFacturaCliente.Fields("Cliente").Value) & "'", "Nombre")
            .Fields("Domicilio").Value = TraerDato("Clientes", "Codigo = '" & EsNulo(rsFacturaCliente.Fields("Cliente").Value) & "'", "Direccion")
            .Fields("Localidad").Value = TraerDato("Clientes", "Codigo = '" & EsNulo(rsFacturaCliente.Fields("Cliente").Value) & "'", "Localidad")
            .Fields("Telefono").Value = TraerDato("Clientes", "Codigo = '" & EsNulo(rsFacturaCliente.Fields("Cliente").Value) & "'", "Telefono")
            '.Fields("Iva").Value = TraerDato("Clientes", "Codigo = '" & vProveedor & "'", "idTipoIva")
            .Fields("Cuit").Value = TraerDato("Clientes", "Codigo = '" & EsNulo(rsFacturaCliente.Fields("Cliente").Value) & "'", "Cuit")
            
            
            .Fields("SubTotal").Value = Val(rsFacturaCliente.Fields("Imp_Bruto").Value)
            .Fields("Descuento").Value = 0
            .Fields("Total").Value = Val(rsFacturaCliente.Fields("Imp_Neto").Value)

            
            .Fields("Comentario").Value = EsNulo(rsFacturaCliente.Fields("Leyenda").Value)
            .Fields("TipoMovimiento").Value = EsNulo(rsFacturaCliente.Fields("Movimiento").Value)
            
            '.Fields("FVencimiento").Value = strfechaMySQL(rsFacturaCliente.Fields("Fec_Vto").Value)
            .Fields("NroInterno").Value = EsNulo(rsFacturaCliente.Fields("Nro_Inter").Value)
            .Fields("NroAsiento").Value = EsNulo(rsFacturaCliente.Fields("Nro_Asto").Value)
            
            .Update
            
            Select Case EsNulo(rsFacturaCliente.Fields("Porc_Iva").Value)
            
                Case 10.5
                    Call EjecutarScript("INSERT INTO IvaFacturaVenta (remito, Iva105, Retenciones, Percepciones, NoGravado, ITC) VALUES (" & vnroremito & "," & Val(rsFacturaCliente.Fields("Imp_Iva").Value) & "," & Val(rsFacturaCliente.Fields("Imp_Ret").Value) & "," & Val(0) & "," & Val(rsFacturaCliente.Fields("Imp_No_Gr").Value) & "," & Val(0) & ");")
                
                Case 21
                    Call EjecutarScript("INSERT INTO IvaFacturaVenta (remito, Iva210, Retenciones, Percepciones, NoGravado, ITC) VALUES (" & vnroremito & "," & Val(rsFacturaCliente.Fields("Imp_Iva").Value) & "," & Val(rsFacturaCliente.Fields("Imp_Ret").Value) & "," & Val(0) & "," & Val(rsFacturaCliente.Fields("Imp_No_Gr").Value) & "," & Val(0) & ");")
                
                Case 27
                    Call EjecutarScript("INSERT INTO IvaFacturaVenta (remito, Iva270, Retenciones, Percepciones, NoGravado, ITC) VALUES (" & vnroremito & "," & Val(rsFacturaCliente.Fields("Imp_Iva").Value) & "," & Val(rsFacturaCliente.Fields("Imp_Ret").Value) & "," & Val(0) & "," & Val(rsFacturaCliente.Fields("Imp_No_Gr").Value) & "," & Val(0) & ");")
            
            End Select
        
        End If
    
    End With
    
    sqlFactura = ""
    
    If rsFactura.State = 1 Then
        rsFactura.Close
        Set rsFactura = Nothing
    End If
    
    
    MigrarFacturaSA = vnroremito
    
If Err Then GrabarLog "MigrarFacturaSA", Err.Number & " " & Err.Description, Me.Caption
End Function
Private Sub MigrarClientesSA(vtabla As String)
On Error Resume Next
    
    vtabla = Replace(vtabla, ".DBF", "")

    Dim rsOrigen As New ADODB.Recordset, rsDestino As New ADODB.Recordset
    Dim sqlOrigen As String, sqlDestino As String
    Dim vcliente As String
    
    sqlOrigen = "SELECT * FROM " & vtabla
    sqlDestino = "SELECT * FROM Clientes WHERE 1=2"
    
    With rsOrigen
        .CursorLocation = adUseClient
        Call .Open(sqlOrigen, connDDBBMigrar, adOpenStatic, adLockPessimistic)
        
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
        
        .MoveFirst
        barra.Value = 0
        barra.Max = .RecordCount
        
    End With
    
    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            Do Until rsOrigen.EOF = True

                vcliente = EsNulo(rsOrigen.Fields("Cliente").Value)
                
                If Not vcliente = "" Then
                
                    .AddNew
                
                
                    .Fields("Codigo").Value = EsNulo(rsOrigen.Fields("Cliente").Value)
                    .Fields("Codigo_Num").Value = Val(rsOrigen.Fields("Razon_Soc").Value)
                    .Fields("Nombre").Value = EsNulo(rsOrigen.Fields("Razon_Soc").Value)
                    .Fields("RazonSocial").Value = EsNulo(rsOrigen.Fields("Razon_Soc").Value)
                    .Fields("Direccion").Value = EsNulo(rsOrigen.Fields("Domicilio").Value)
                    .Fields("Localidad").Value = EsNulo(rsOrigen.Fields("Localidad").Value)
                    .Fields("CodigoPostal").Value = EsNulo(rsOrigen.Fields("Cod_Postal").Value)
                
                    Select Case EsNulo(rsOrigen.Fields("Provincia").Value)
                    
                        Case "B"
                            .Fields("Provincia").Value = "BUENOS AIRES"
                        Case "C"
                            .Fields("Provincia").Value = "CATAMARCA"
                        Case "CF"
                            .Fields("Provincia").Value = "CAPITAL FEDERAL"
                        Case "CO"
                            .Fields("Provincia").Value = "CORRIENTES"
                        Case "ER"
                            .Fields("Provincia").Value = "ENTRE RIOS"
                        Case "F"
                            .Fields("Provincia").Value = "FORMOSA"
                        Case "H"
                            .Fields("Provincia").Value = "CHACO"
                        Case "J"
                            .Fields("Provincia").Value = "JUJUY"
                        Case "L"
                            .Fields("Provincia").Value = "SAN LUIS"
                        Case "M"
                            .Fields("Provincia").Value = "MENDOZA"
                        Case "MS"
                            .Fields("Provincia").Value = "MISIONES"
                        Case "P"
                            .Fields("Provincia").Value = "LA PAMPA"
                        Case "Q"
                            .Fields("Provincia").Value = "NEUQUEN"
                        Case "R"
                            .Fields("Provincia").Value = "RIO NEGRO"
                        Case "RJ"
                            .Fields("Provincia").Value = "LA RIOJA"
                        Case "T"
                            .Fields("Provincia").Value = "TUCUMAN"
                        Case "S"
                            .Fields("Provincia").Value = "SANTIAGO DEL ESTERO"
                        Case "SA"
                            .Fields("Provincia").Value = "SALTA"
                        Case "SE"
                            .Fields("Provincia").Value = "SANTIAGO DEL ESTERO"
                        Case "SJ"
                            .Fields("Provincia").Value = "SAN JUAN"
                        Case "X"
                            .Fields("Provincia").Value = "CORDOBA"
                        Case Else
                            Debug.Print ("OJO........  " & EsNulo(rsOrigen.Fields("Provincia").Value))
                            
                    End Select
                
                    .Fields("Telefono").Value = EsNulo(rsOrigen.Fields("telefono1").Value)
                    .Fields("Fax").Value = EsNulo(rsOrigen.Fields("telefono2").Value)
                    .Fields("Celular").Value = EsNulo(rsOrigen.Fields("telefono3").Value)
                
                    .Fields("Cuit").Value = Replace(EsNulo(rsOrigen.Fields("Cuit").Value), " ", "-")
                    .Fields("idTipoIva").Value = "0" & EsNulo(rsOrigen.Fields("Cod_Impues").Value)
                
                    .Fields("Saldo").Value = Val(EsNulo(rsOrigen.Fields("Saldo_R").Value))
                    
                    .Update
                Else


                End If
                
                rsOrigen.MoveNext
                barra.Value = barra.Value + 1
            Loop
        End If
        
    End With
    
    sqlOrigen = ""
    sqlDestino = ""

    If rsOrigen.State = 1 Then
        rsOrigen.Close
        Set rsOrigen = Nothing
    End If

If Err Then GrabarLog "MigrarClientesSA", Err.Number & " " & Err.Description, "BasSisAgro"
End Sub
Private Sub MigrarBancosSA(vtabla As String)
On Error Resume Next
    
    vtabla = Replace(vtabla, ".DBF", "")

    Dim rsOrigen As New ADODB.Recordset, rsDestino As New ADODB.Recordset
    Dim sqlOrigen As String, sqlDestino As String

    
    sqlOrigen = "SELECT * FROM " & vtabla & " WHERE (Empresa = '1') OR (Empresa = '') OR (Empresa is Null)"
    sqlDestino = "SELECT * FROM Bancos WHERE 1=2"
    
    With rsOrigen
        .CursorLocation = adUseClient
        Call .Open(sqlOrigen, connDDBBMigrar, adOpenStatic, adLockPessimistic)
        
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
        
        .MoveFirst
        barra.Value = 0
        barra.Max = .RecordCount
        
    End With

    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            Do Until rsOrigen.EOF = True
                DoEvents
                .AddNew

                .Fields("idBancos").Value = EsNulo(rsOrigen.Fields("Banco").Value)
                .Fields("Descripcion").Value = EsNulo(rsOrigen.Fields("Descrip").Value)
                .Fields("EsCaja").Value = EsNulo(rsOrigen.Fields("Caja").Value)
                .Fields("CuentaContableAsociada").Value = EsNulo(rsOrigen.Fields("Cuenta").Value)

                .Update

                
                rsOrigen.MoveNext
                barra.Value = barra.Value + 1
            Loop
        End If
        
    End With
    
    sqlOrigen = ""
    sqlDestino = ""

    If rsOrigen.State = 1 Then
        rsOrigen.Close
        Set rsOrigen = Nothing
    End If

If Err Then GrabarLog "MigrarBancosSA", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub MigrarBancosCuentasSA(vtabla As String)
On Error Resume Next
    
    vtabla = Replace(vtabla, ".DBF", "")

    Dim rsOrigen As New ADODB.Recordset, rsDestino As New ADODB.Recordset
    Dim sqlOrigen As String, sqlDestino As String
    
    sqlOrigen = "SELECT * FROM " & vtabla & " WHERE (Empresa = '1') OR (Empresa = '') OR (Empresa is Null)"
    sqlDestino = "SELECT * FROM BancosCuentas WHERE 1=2"
    
    With rsOrigen
        .CursorLocation = adUseClient
        Call .Open(sqlOrigen, connDDBBMigrar, adOpenStatic, adLockPessimistic)
        
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
        
        .MoveFirst
        barra.Value = 0
        barra.Max = .RecordCount
        
    End With

    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            Do Until rsOrigen.EOF = True
                DoEvents
                
                .AddNew

                .Fields("idBancos").Value = EsNulo(rsOrigen.Fields("cod_bco").Value)
                .Fields("Descripcion").Value = EsNulo(rsOrigen.Fields("Descrip").Value)
                .Fields("Cuenta").Value = EsNulo(rsOrigen.Fields("Cuenta").Value)
                .Fields("CuentaContableAsociada").Value = EsNulo(rsOrigen.Fields("Cuenta_Con").Value)
                .Fields("idTipoCuentaBanco").Value = EsNulo(rsOrigen.Fields("Tipo").Value)

                .Update

                rsOrigen.MoveNext
                barra.Value = barra.Value + 1
            Loop
        End If
        
    End With
    
    sqlOrigen = ""
    sqlDestino = ""

    If rsOrigen.State = 1 Then
        rsOrigen.Close
        Set rsOrigen = Nothing
    End If

If Err Then GrabarLog "MigrarBancosCuentasSA", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub MigrarAsientosSA(vtabla As String)
On Error Resume Next
    
    vtabla = Replace(vtabla, ".DBF", "")

    Dim rsOrigen As New ADODB.Recordset, rsDestino As New ADODB.Recordset
    Dim sqlOrigen As String, sqlDestino As String
    Dim vnrointerno() As Long
    
    sqlOrigen = "SELECT * FROM " & vtabla & " WHERE (Empresa = '1')"
    sqlDestino = "SELECT * FROM Asientos WHERE 1=2"
    
    With rsOrigen
        .CursorLocation = adUseClient
        Call .Open(sqlOrigen, connDDBBMigrar, adOpenStatic, adLockPessimistic)
        
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
        
        .MoveFirst
        barra.Value = 0
        barra.Max = .RecordCount
        ReDim vnrointerno(2)
    End With
    
    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            Do Until rsOrigen.EOF = True
                vnrointerno(0) = 0
                vnrointerno(1) = 0
                vnrointerno(2) = 0
                DoEvents
                
                .AddNew

                .Fields("Fecha").Value = EsNulo(rsOrigen.Fields("Fec_Asto").Value)
                .Fields("Numero").Value = EsNulo(rsOrigen.Fields("Nro_Asto").Value)
                .Fields("Leyenda").Value = EsNulo(rsOrigen.Fields("Leyenda").Value)

                
                'vNroInterno(0) = Val(TraerDato("InternoAsientoBancos", "NRO_ASTO = " & Val(EsNulo(rsOrigen.Fields("Nro_Asto").value)) & "", "Nro_Inter", pathDBMigrar("ARV")))
                
                'If vNroInterno(0) = 0 Then
                '    vNroInterno(1) = Val(TraerDato("InternoAsientoCompras", "NRO_ASTO = " & Val(EsNulo(rsOrigen.Fields("Nro_Asto").value)) & "", "Nro_Inter", pathDBMigrar("ARV")))
                    
                '    If vNroInterno(1) = 0 Then
                '        vNroInterno(2) = TraerDato("InternoAsientoVentas", "NRO_ASTO = " & Val(EsNulo(rsOrigen.Fields("Nro_Asto").value)) & "", "Nro_Inter", pathDBMigrar("ARV"))
                        
                '        .Fields("NroInterno").value = vNroInterno(2)
                        
                        
                '    Else
                '        .Fields("NroInterno").value = vNroInterno(1)
                '    End If
                'Else
                '    .Fields("NroInterno").value = vNroInterno(0)
                'End If
                    

                .Update

                rsOrigen.MoveNext
                barra.Value = barra.Value + 1
            Loop
        End If
        
    End With
    
    sqlOrigen = ""
    sqlDestino = ""

    If rsOrigen.State = 1 Then
        rsOrigen.Close
        Set rsOrigen = Nothing
    End If

If Err Then GrabarLog "MigrarAsientosSA", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub MigrarAsientosDetalleSA(vtabla As String)
On Error Resume Next
    
    vtabla = Replace(vtabla, ".DBF", "")

    Dim rsOrigen As New ADODB.Recordset, rsDestino As New ADODB.Recordset
    Dim sqlOrigen As String, sqlDestino As String
    Dim vNroAsientoAn As Long, vlinea As Long
    
    sqlOrigen = "SELECT * FROM " & vtabla & " WHERE (Empresa = '1')"
    sqlDestino = "SELECT * FROM AsientosDetalle WHERE 1=2"
    
    With rsOrigen
        .CursorLocation = adUseClient
        Call .Open(sqlOrigen, connDDBBMigrar, adOpenStatic, adLockPessimistic)
        
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
        
        .MoveFirst
        barra.Value = 0
        barra.Max = .RecordCount
        
    End With
    
    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
            
            vlinea = 1
            vNroAsientoAn = 0
            vNroAsientoAn = EsNulo(rsOrigen.Fields("Nro_Asto").Value)
            
            Do Until rsOrigen.EOF = True
                DoEvents
               .AddNew

                .Fields("Numero").Value = EsNulo(rsOrigen.Fields("Nro_Asto").Value)
                
                If vNroAsientoAn = EsNulo(rsOrigen.Fields("Nro_Asto").Value) Then
                    .Fields("Linea").Value = vlinea
                Else
                    .Fields("Linea").Value = 1
                    vNroAsientoAn = EsNulo(rsOrigen.Fields("Nro_Asto").Value)
                End If
                
                .Fields("CodigoCuenta").Value = EsNulo(rsOrigen.Fields("Cuenta").Value)
                
                If EsNulo(rsOrigen.Fields("DB_CR").Value) = "D" Then
                    .Fields("Debe").Value = Val(rsOrigen.Fields("Importe").Value)
                Else
                    .Fields("Haber").Value = Val(rsOrigen.Fields("Importe").Value)
                End If

                .Update

                rsOrigen.MoveNext
                barra.Value = barra.Value + 1
                vlinea = vlinea + 1
            
                If Not Err.Description = "" Then
                    MsgBox ""
                End If
            Loop
        End If
        
        
        
    End With
    
    sqlOrigen = ""
    sqlDestino = ""

    If rsOrigen.State = 1 Then
        rsOrigen.Close
        Set rsOrigen = Nothing
    End If

If Err Then GrabarLog "MigrarAsientosDetalleSA", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub MigrarBancosMovimientosSA(vtabla As String)
On Error Resume Next
    
    vtabla = Replace(vtabla, ".DBF", "")

    Dim rsOrigen As New ADODB.Recordset, rsDestino As New ADODB.Recordset
    Dim sqlOrigen As String, sqlDestino As String, vIDBC As Long
    
    sqlOrigen = "SELECT * FROM " & vtabla & " WHERE (Empresa = '1') AND (Tipo_Valor = 'EF' OR Tipo_Valor = 'CH')"
    sqlDestino = "SELECT * FROM BancosMovimientos WHERE 1=2"
    
    With rsOrigen
        .CursorLocation = adUseClient
        Call .Open(sqlOrigen, connDDBBMigrar, adOpenStatic, adLockPessimistic)
        
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
        
        .MoveFirst
        barra.Value = 0
        barra.Max = .RecordCount
        
    End With
    
    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            Do Until rsOrigen.EOF = True
                DoEvents

                vIDBC = 0
 
                If Not EsNulo(rsOrigen.Fields("Cod_Bco").Value) = "" Then
               
                    Select Case EsNulo(rsOrigen.Fields("Nro_Cuenta").Value)
                    
                        Case "11339/9"
                            vIDBC = 1
                        Case "1463-8 034-3"
                            vIDBC = 2
                        Case "01"
                            vIDBC = 3
                        Case "3-124-1995-9"
                            vIDBC = 4
                        Case "107-008409/3"
                            vIDBC = 5
                        Case "454-20-000439/9"
                            vIDBC = 6
                        Case "20000.207/37"
                            vIDBC = 7
                        Case "BBVA-TARJ.VISA"
                            vIDBC = 8
                        Case "CABAL AGRO"
                            'vIDBC = 9
                        Case "CABAL AGRO"
                            vIDBC = 10
                        Case "BCO.MACRO-VISA"
                            vIDBC = 11
                    
                    End Select
                    
                    .AddNew

                    .Fields("idBancos").Value = EsNulo(rsOrigen.Fields("Cod_Bco").Value)
                    
                    .Fields("idBancosCuentas").Value = vIDBC
                    
                    .Fields("Fecha").Value = strfechaMySQL(rsOrigen.Fields("Fec_Valor").Value)
                    .Fields("FechaEmision").Value = strfechaMySQL(rsOrigen.Fields("Fec_Valor").Value)

                    Select Case EsNulo(rsOrigen.Fields("DB_CR").Value)
                
                        Case "D"
                            .Fields("Debito").Value = Val(Format(rsOrigen.Fields("Importe").Value, "######0.00"))
                    
                        Case "H"
                            .Fields("Credito").Value = Val(Format(rsOrigen.Fields("Importe").Value, "######0.00"))
                    
                        Case Else
                            MsgBox "Algo raro paso", vbExclamation, "Mensaje ..."
                
                    End Select

                    .Fields("NroInterno").Value = Val(rsOrigen.Fields("Nro_Inter").Value)
                    .Fields("NroAsiento").Value = Val(rsOrigen.Fields("Nro_Asto").Value)
                    
                    .Fields("Comentario").Value = EsNulo(rsOrigen.Fields("Observac").Value)
                    .Fields("TipoMovimiento").Value = EsNulo(rsOrigen.Fields("Movimiento").Value)
                        
                    .Fields("idTipoValor").Value = EsNulo(rsOrigen.Fields("Tipo_Valor").Value)
                    If EsNulo(.Fields("idTipoValor").Value) = "CH" Then
                        .Fields("FechaValor").Value = strfechaMySQL(rsOrigen("Fec_Valor").Value)
                        .Fields("NroCheque").Value = Val(rsOrigen.Fields("Nro_Valor").Value)
                    Else
                        .Fields("NroCheque").Value = 0
                    End If
                    

                
                    .Update
                    
                    
                    If IsNull(.Fields("idBancos").Value) = True Then
                    
                        MsgBox Null
                    End If
                End If
                    
                rsOrigen.MoveNext
                If (barra.Max - 5) = (barra.Value) Then MsgBox "terminando"
                barra.Value = barra.Value + 1
            Loop
        
        End If
        
    End With
    
    sqlOrigen = ""
    sqlDestino = ""

    If rsOrigen.State = 1 Then
        rsOrigen.Close
        Set rsOrigen = Nothing
    End If

If Err Then GrabarLog "MigrarBancosMovimientosSA", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub MigrarCajasSA(vtabla As String)
On Error Resume Next
    
    vtabla = Replace(vtabla, ".DBF", "")

    Dim rsOrigen As New ADODB.Recordset, rsDestino As New ADODB.Recordset
    Dim sqlOrigen As String, sqlDestino As String
    
    sqlOrigen = "SELECT * FROM " & vtabla & " WHERE (Empresa = '1') AND ((Caja <> '') AND NOT (Caja IS NULL))"
    sqlDestino = "SELECT * FROM BancosMovimientos"
    
    With rsOrigen
        .Close
        .CursorLocation = adUseClient
        Call .Open(sqlOrigen, connDDBBMigrar, adOpenStatic, adLockPessimistic)
        
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
        
        .MoveFirst
        barra.Value = 0
        barra.Max = .RecordCount
        
    End With
    
    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            Do Until rsOrigen.EOF = True
                DoEvents
                
                'If rsOrigen.Fields("Nro_Inter").value = 889 Then MsgBox rsOrigen.Fields("Nro_Inter").value
                
                If Not EsNulo(rsOrigen.Fields("Caja").Value) = "" Then
                    .Close
                    
                    sqlOrigen = "SELECT * FROM BancosMovimientos WHERE ((NroInterno = " & Val(rsOrigen.Fields("Nro_Inter").Value) & ") AND (idTipoValor = 'CH') AND (NroCheque = " & Val(rsOrigen.Fields("Nro_Valor").Value) & "))"
                    
                    Call .Open(sqlOrigen, ConnDDBB, adOpenDynamic, adLockBatchOptimistic)
                    
                    If .EOF = True Then
                        If vtabla = "APCH00" Then
                            If rsOrigen.Fields("Tipo_Valor").Value = "CH" Then
                                Call EjecutarScript("INSERT INTO BancosMovimientos (idBancos, Fecha, FechaEmision, Credito, NroInterno, NroCheque, Comentario, TipoMovimiento, idTipoValor) VALUES ('" & EsNulo(rsOrigen.Fields("Caja").Value) & "','" & strfechaMySQL(rsOrigen.Fields("Fec_Valor").Value) & "','" & strfechaMySQL(rsOrigen.Fields("Fec_Valor").Value) & "'," & EsNulo(rsOrigen.Fields("Importe").Value) & "," & EsNulo(rsOrigen.Fields("Nro_Inter").Value) & "," & EsNulo(rsOrigen.Fields("Nro_Valor").Value) & ",'" & EsNulo(rsOrigen.Fields("Observ").Value) & "','" & EsNulo(rsOrigen.Fields("Movimiento").Value) & "','" & EsNulo(rsOrigen.Fields("Tipo_Valor").Value) & "')")
                            End If
                        Else
                            If rsOrigen.Fields("Tipo_Valor").Value = "CH" Then
                                Call EjecutarScript("INSERT INTO BancosMovimientos (idBancos, Fecha, FechaEmision, Debito, NroInterno, NroCheque, Comentario, TipoMovimiento, idTipoValor) VALUES ('" & EsNulo(rsOrigen.Fields("Caja").Value) & "','" & strfechaMySQL(rsOrigen.Fields("Fec_Valor").Value) & "','" & strfechaMySQL(rsOrigen.Fields("Fec_Valor").Value) & "'," & EsNulo(rsOrigen.Fields("Importe").Value) & "," & EsNulo(rsOrigen.Fields("Nro_Inter").Value) & "," & EsNulo(rsOrigen.Fields("Nro_Valor").Value) & ",'" & EsNulo(rsOrigen.Fields("Observ").Value) & "','" & EsNulo(rsOrigen.Fields("Movimiento").Value) & "','" & EsNulo(rsOrigen.Fields("Tipo_Valor").Value) & "')")
                            Else
                                Call EjecutarScript("INSERT INTO BancosMovimientos (idBancos, Fecha, FechaEmision, Debito, NroInterno, NroCheque, Comentario, TipoMovimiento, idTipoValor) VALUES ('" & EsNulo(rsOrigen.Fields("Caja").Value) & "','" & strfechaMySQL(rsOrigen.Fields("Fec_Valor").Value) & "','" & strfechaMySQL(rsOrigen.Fields("Fec_Valor").Value) & "'," & EsNulo(rsOrigen.Fields("Importe").Value) & "," & EsNulo(rsOrigen.Fields("Nro_Inter").Value) & "," & EsNulo(rsOrigen.Fields("Nro_Valor").Value) & ",'" & EsNulo(rsOrigen.Fields("Observ").Value) & "','" & EsNulo(rsOrigen.Fields("Movimiento").Value) & "','" & EsNulo(rsOrigen.Fields("Tipo_Valor").Value) & "')")
                            End If
                        End If
                    Else
                        Call EjecutarScript("UPDATE BancosMovimientos SET Fecha = '" & strfechaMySQL(rsOrigen.Fields("Fec_Valor").Value) & "',FechaEmision = '" & strfechaMySQL(rsOrigen.Fields("Fec_Valor").Value) & "' WHERE (NroInterno = " & Val(rsOrigen.Fields("Nro_Inter").Value) & ") AND NroCheque = " & Val(rsOrigen.Fields("Nro_Valor").Value) & ")")
                    End If
                   
                    '.Fields("idBancos").Value = EsNulo(rsOrigen.Fields("Caja").Value)
                    '.Fields("Fecha").Value = strfechaMySQL(rsOrigen.Fields("Fec_Valor").Value)
                    '.Fields("FechaEmision").Value = strfechaMySQL(rsOrigen.Fields("Fec_Valor").Value)
    
                    'If vTabla = "APCH00" Then
                    '    .Fields("Credito").Value = rsOrigen.Fields("Importe").Value
                    'Else
                    '    .Fields("Debito").Value = rsOrigen.Fields("Importe").Value
                    'End If
                    '.Fields("NroInterno").Value = Val(rsOrigen.Fields("Nro_Inter").Value)
                    '.Fields("NroCheque").Value = Val(rsOrigen.Fields("Nro_Valor").Value)
                    '.Fields("Comentario").Value = EsNulo(rsOrigen.Fields("Observ").Value)
                    '.Fields("TipoMovimiento").Value = EsNulo(rsOrigen.Fields("Movimiento").Value)
                    '.Fields("idTipoValor").Value = EsNulo(rsOrigen.Fields("Tipo_Valor").Value)
    
                    '.Update
                Else
                    
                    MsgBox Err.Description
                    
                End If
                
                'If rsOrigen.Fields("Nro_Inter").value = 12236 Then MsgBox rsOrigen.Fields("Nro_Inter").value
                
                rsOrigen.MoveNext
                barra.Value = barra.Value + 1
            Loop
        
        End If
        
    End With
    
    sqlOrigen = ""
    sqlDestino = ""

    If rsOrigen.State = 1 Then
        rsOrigen.Close
        Set rsOrigen = Nothing
    End If

If Err Then GrabarLog "MigrarCajasSA", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub MigrarBancosMovSA(vtabla As String)
On Error Resume Next
    
    vtabla = Replace(vtabla, ".DBF", "")

    Dim rsOrigen As New ADODB.Recordset, rsDestino As New ADODB.Recordset
    Dim sqlOrigen As String, sqlDestino As String, vIDBC As Long
    
    sqlOrigen = "SELECT * FROM " & vtabla & " WHERE (Empresa = '1') AND ((COD_BCO <> '') AND NOT (COD_BCO IS NULL))"
    sqlDestino = "SELECT * FROM BancosMovimientos"
    
    With rsOrigen
        .CursorLocation = adUseClient
        Call .Open(sqlOrigen, connDDBBMigrar, adOpenStatic, adLockPessimistic)
        
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
        
        .MoveFirst
        barra.Value = 0
        barra.Max = .RecordCount
        
    End With
    
    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenDynamic, adLockBatchOptimistic)
    
        If .State = 1 Then
        
            Do Until rsOrigen.EOF = True
                DoEvents
                
                
                'If Val(rsOrigen.Fields("Nro_Inter").value) = 889 Then MsgBox Val(rsOrigen.Fields("Nro_Inter").value)
                
                If Not EsNulo(rsOrigen.Fields("Cod_Bco").Value) = "" Then
                    .Close
                    
                    sqlOrigen = "SELECT * FROM BancosMovimientos WHERE (TipoMovimiento = '" & EsNulo(rsOrigen.Fields("Movimiento").Value) & "') AND (NroInterno = " & Val(rsOrigen.Fields("Nro_Inter").Value) & ") AND (NroCheque = " & Val(rsOrigen.Fields("Nro_Valor").Value) & ")  AND (idBancos = '" & EsNulo(rsOrigen.Fields("Cod_Bco").Value) & "') AND (idTipoValor = 'CH')"
                    
                    Call .Open(sqlOrigen, ConnDDBB, adOpenDynamic, adLockBatchOptimistic)
                        
                    vIDBC = 0
                
                    Select Case EsNulo(rsOrigen.Fields("Nro_Cuenta").Value)
                    
                        Case "11339/9"
                            vIDBC = 1
                        Case "1463-8 034-3"
                            vIDBC = 2
                        Case "01"
                            vIDBC = 3
                        Case "3-124-1995-9"
                            vIDBC = 4
                        Case "107-008409/3"
                            vIDBC = 5
                        Case "454-20-000439/9"
                            vIDBC = 6
                        Case "20000.207/37"
                            vIDBC = 7
                        Case "BBVA-TARJ.VISA"
                            vIDBC = 8
                        Case "CABAL AGRO"
                            'vIDBC = 9
                        Case "CABAL AGRO"
                            vIDBC = 10
                        Case "BCO.MACRO-VISA"
                            vIDBC = 11
                    
                    End Select
                    
                    If .EOF = True Then
                        
                        If vtabla = "APCH00" Then
                            If rsOrigen.Fields("Tipo_Valor").Value = "CH" Then
                                Call EjecutarScript("INSERT INTO BancosMovimientos (idBancos, idBancosCuentas, Fecha, FechaEmision, Credito, NroInterno, NroCheque, Comentario, TipoMovimiento, idTipoValor) VALUES ('" & EsNulo(rsOrigen.Fields("COD_BCO").Value) & "'," & vIDBC & ",'" & strfechaMySQL(rsOrigen.Fields("Fec_Valor").Value) & "','" & strfechaMySQL(rsOrigen.Fields("Fec_Valor").Value) & "'," & EsNulo(rsOrigen.Fields("Importe").Value) & "," & EsNulo(rsOrigen.Fields("Nro_Inter").Value) & "," & EsNulo(rsOrigen.Fields("Nro_Valor").Value) & ",'" & EsNulo(rsOrigen.Fields("Observ").Value) & "','" & EsNulo(rsOrigen.Fields("Movimiento").Value) & "','" & EsNulo(rsOrigen.Fields("Tipo_Valor").Value) & "')")
                            End If
                        Else
                            If rsOrigen.Fields("Tipo_Valor").Value = "CH" Then
                                Call EjecutarScript("INSERT INTO BancosMovimientos (idBancos, idBancosCuentas, Fecha, FechaEmision, Debito, NroInterno, NroCheque, Comentario, TipoMovimiento, idTipoValor) VALUES ('" & EsNulo(rsOrigen.Fields("COD_BCO").Value) & "', " & vIDBC & ",'" & strfechaMySQL(rsOrigen.Fields("Fec_Valor").Value) & "','" & strfechaMySQL(rsOrigen.Fields("Fec_Valor").Value) & "'," & EsNulo(rsOrigen.Fields("Importe").Value) & "," & EsNulo(rsOrigen.Fields("Nro_Inter").Value) & "," & EsNulo(rsOrigen.Fields("Nro_Valor").Value) & ",'" & EsNulo(rsOrigen.Fields("Observ").Value) & "','" & EsNulo(rsOrigen.Fields("Movimiento").Value) & "','" & EsNulo(rsOrigen.Fields("Tipo_Valor").Value) & "')")
                            End If
                        End If
                    Else
                        
                        Call EjecutarScript("UPDATE BancosMovimientos SET Fecha = '" & strfechaMySQL(rsOrigen.Fields("Fec_Valor").Value) & "',FechaEmision = '" & strfechaMySQL(rsOrigen.Fields("Fec_Valor").Value) & "' WHERE (NroInterno = " & Val(rsOrigen.Fields("Nro_Inter").Value) & ") AND (NroCheque = " & Val(rsOrigen.Fields("Nro_Valor").Value) & ")")
                        
                    End If
                Else
                    MsgBox rsOrigen.Fields("Cod_Bco").Value
                End If
                
                rsOrigen.MoveNext
                barra.Value = barra.Value + 1
            Loop
        
        End If
        
    End With
    
    sqlOrigen = ""
    sqlDestino = ""

    If rsOrigen.State = 1 Then
        rsOrigen.Close
        Set rsOrigen = Nothing
    End If

If Err Then GrabarLog "MigrarCajasSA", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub MigrarChequesCSA(vtabla As String)
On Error Resume Next
    
    vtabla = Replace(vtabla, ".DBF", "")

    Dim rsOrigen As New ADODB.Recordset, rsDestino As New ADODB.Recordset
    Dim sqlOrigen As String, sqlDestino As String, vIDBC As Long
    
    sqlOrigen = "SELECT * FROM " & vtabla & " WHERE (Empresa = '1') AND (Tipo_Valor = 'CH')"
    sqlDestino = "SELECT * FROM Cheques WHERE 1=2"
    
    With rsOrigen
        .CursorLocation = adUseClient
        Call .Open(sqlOrigen, connDDBBMigrar, adOpenStatic, adLockPessimistic)
        
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
        
        .MoveFirst
        barra.Value = 0
        barra.Max = .RecordCount
        
    End With
    
    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            Do Until rsOrigen.EOF = True
                DoEvents
                
                .AddNew
                
                .Fields("idEstadoCheque").Value = 1
                .Fields("Fecha").Value = EsNulo(rsOrigen.Fields("Fec_Valor").Value)
                .Fields("Codigo").Value = EsNulo(rsOrigen.Fields("Cliente").Value)
                .Fields("Nombre").Value = TraerDato("Clientes", "Codigo = '" & EsNulo(rsOrigen.Fields("Cliente").Value) & "'", "Nombre")
                .Fields("NCheque").Value = EsNulo(rsOrigen.Fields("Nro_Valor").Value)
                '.Fields("Firmate").Value = ""
                .Fields("CP").Value = "c"
                
                .Fields("Monto").Value = EsNulo(rsOrigen.Fields("Importe").Value)
                .Fields("Endoso").Value = ""
                
                .Fields("Remito").Value = 0
                .Fields("NroInterno").Value = EsNulo(rsOrigen.Fields("Nro_Inter").Value)
                .Fields("idBancos").Value = EsNulo(rsOrigen.Fields("Cod_Bco").Value)
                
                vIDBC = 0
                
                Select Case EsNulo(rsOrigen.Fields("Nro_Cuenta").Value)
                
                    Case "11339/9"
                        vIDBC = 1
                    Case "1463-8 034-3"
                        vIDBC = 2
                    Case "01"
                        vIDBC = 3
                    Case "3-124-1995-9"
                        vIDBC = 4
                    Case "107-008409/3"
                        vIDBC = 5
                    Case "454-20-000439/9"
                        vIDBC = 6
                    Case "20000.207/37"
                        vIDBC = 7
                    Case "BBVA-TARJ.VISA"
                        vIDBC = 8
                    Case "CABAL AGRO"
                        'vIDBC = 9
                    Case "CABAL AGRO"
                        vIDBC = 10
                    Case "BCO.MACRO-VISA"
                        vIDBC = 11
                
                End Select
                    
                .Fields("idBancosCuentas").Value = Val(vIDBC)

                .Fields("Observaciones").Value = EsNulo(rsOrigen.Fields("Observ").Value)
                .Fields("FechaDeposito").Value = (rsOrigen.Fields("Fec_Depos").Value)
                .Fields("FechaAcreditacion").Value = (rsOrigen.Fields("Fec_Acre").Value)
                
                .Fields("TipoMovimiento").Value = EsNulo(rsOrigen.Fields("Movimiento").Value)
                .Update

                
                rsOrigen.MoveNext
                barra.Value = barra.Value + 1
            Loop
        End If
        
    End With
    
    sqlOrigen = ""
    sqlDestino = ""

    If rsOrigen.State = 1 Then
        rsOrigen.Close
        Set rsOrigen = Nothing
    End If

If Err Then GrabarLog "MigrarChequesCSA", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub MigrarChequesPSA(vtabla As String)
On Error Resume Next

    vtabla = Replace(vtabla, ".DBF", "")

    Dim rsOrigen As New ADODB.Recordset, rsDestino As New ADODB.Recordset
    Dim sqlOrigen As String, sqlDestino As String, vIDBC As Long
    
    sqlOrigen = "SELECT * FROM " & vtabla & " WHERE (Empresa = '1') AND (Tipo_Valor = 'CH')"
    sqlDestino = "SELECT * FROM Cheques WHERE 1=2"
    
    With rsOrigen
        .CursorLocation = adUseClient
        Call .Open(sqlOrigen, connDDBBMigrar, adOpenStatic, adLockPessimistic)
        
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
        
        .MoveFirst
        barra.Value = 0
        barra.Max = .RecordCount
        
    End With
    
    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            Do Until rsOrigen.EOF = True
                DoEvents
                
                .AddNew

                
                .Fields("idEstadoCheque").Value = 1
                .Fields("Fecha").Value = EsNulo(rsOrigen.Fields("Fec_Valor").Value)
                .Fields("Codigo").Value = EsNulo(rsOrigen.Fields("Proveedor").Value)
                .Fields("Nombre").Value = TraerDato("Proveedores", "Codigo = '" & EsNulo(rsOrigen.Fields("Proveedor").Value) & "'", "Nombre")
                .Fields("NCheque").Value = EsNulo(rsOrigen.Fields("Nro_Valor").Value)
                .Fields("Firmate").Value = ""
                .Fields("CP").Value = "p"
                
                .Fields("Deposito").Value = ""
                .Fields("Monto").Value = EsNulo(rsOrigen.Fields("Importe").Value)
                .Fields("Endoso").Value = ""
                
                .Fields("Remito").Value = 0
                .Fields("NroInterno").Value = EsNulo(rsOrigen.Fields("Nro_Inter").Value)
                .Fields("idBancos").Value = EsNulo(rsOrigen.Fields("Cod_Bco").Value)
                
                vIDBC = 0
                
                Select Case EsNulo(rsOrigen.Fields("Nro_Cuenta").Value)
                
                    Case "11339/9"
                        vIDBC = 1
                    Case "1463-8 034-3"
                        vIDBC = 2
                    Case "01"
                        vIDBC = 3
                    Case "3-124-1995-9"
                        vIDBC = 4
                    Case "107-008409/3"
                        vIDBC = 5
                    Case "454-20-000439/9"
                        vIDBC = 6
                    Case "20000.207/37"
                        vIDBC = 7
                    Case "BBVA-TARJ.VISA"
                        vIDBC = 8
                    Case "CABAL AGRO"
                        'vIDBC = 9
                    Case "CABAL AGRO"
                        vIDBC = 10
                    Case "BCO.MACRO-VISA"
                        vIDBC = 11
                End Select
                
                .Fields("idBancosCuentas").Value = Val(vIDBC)

                .Fields("Observaciones").Value = EsNulo(rsOrigen.Fields("Observ").Value)
                .Fields("TipoMovimiento").Value = EsNulo(rsOrigen.Fields("Movimiento").Value)
                .Fields("FechaDeposito").Value = EsNulo(rsOrigen.Fields("Fec_Valor").Value)
                .Fields("FechaAcreditacion").Value = EsNulo(rsOrigen.Fields("Fec_Valor").Value)
                .Update

                rsOrigen.MoveNext
                barra.Value = barra.Value + 1
            Loop
        End If
        
    End With
    
    sqlOrigen = ""
    sqlDestino = ""

    If rsOrigen.State = 1 Then
        rsOrigen.Close
        Set rsOrigen = Nothing
    End If

If Err Then GrabarLog "MigrarChequesPSA", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub MigrarCobrosSA(vtabla As String)
On Error Resume Next

    vtabla = Replace(vtabla, ".DBF", "")

    Dim rsOrigen As New ADODB.Recordset, rsDestino As New ADODB.Recordset
    Dim sqlOrigen As String, sqlDestino As String
    
    sqlOrigen = "SELECT * FROM " & vtabla & " WHERE (Empresa_D = '1')"
    sqlDestino = "SELECT * FROM Cobros WHERE 1=2"
    
    With rsOrigen
        .CursorLocation = adUseClient
        Call .Open(sqlOrigen, connDDBBMigrar, adOpenStatic, adLockPessimistic)
        
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        Else
            .MoveFirst
            barra.Value = 0
            barra.Max = .RecordCount
        End If
        
    End With
    
    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            Do Until rsOrigen.EOF = True
                DoEvents
                
                .AddNew
                
                .Fields("CodigoCliente").Value = EsNulo(rsOrigen.Fields("Cliente").Value)
                .Fields("Fecha").Value = EsNulo(rsOrigen.Fields("Fec_Credit").Value)
                .Fields("Remito").Value = TraerDato("Factura", "(NroInterno = " & EsNulo(rsOrigen.Fields("NRO_INT_DB").Value) & ")", "Remito")
                '.Fields("idMedioPago").Value = ""
                .Fields("Importe").Value = EsNulo(rsOrigen.Fields("Importe").Value)
                .Fields("TipoMovimiento").Value = EsNulo(rsOrigen.Fields("MovimientC").Value)
                .Fields("NroInterno").Value = EsNulo(rsOrigen.Fields("Nro_Int_Cr").Value)
                
                .Update

                rsOrigen.MoveNext
                barra.Value = barra.Value + 1
            Loop
        End If
        
    End With
    
    sqlOrigen = ""
    sqlDestino = ""

    If rsOrigen.State = 1 Then
        rsOrigen.Close
        Set rsOrigen = Nothing
    End If

If Err Then GrabarLog "MigrarChequesPSA", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub MigrarPagosSA(vtabla As String)
On Error Resume Next

    vtabla = Replace(vtabla, ".DBF", "")

    Dim rsOrigen As New ADODB.Recordset, rsDestino As New ADODB.Recordset
    Dim sqlOrigen As String, sqlDestino As String
    
    sqlOrigen = "SELECT * FROM " & vtabla & " WHERE (Empresa_D = '1')"
    sqlDestino = "SELECT * FROM Pagos WHERE 1=2"
    
    With rsOrigen
        .CursorLocation = adUseClient
        Call .Open(sqlOrigen, connDDBBMigrar, adOpenStatic, adLockPessimistic)
        
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        Else
            .MoveFirst
            barra.Value = 0
            barra.Max = .RecordCount
        End If
        
    End With
    
    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            Do Until rsOrigen.EOF = True
                DoEvents
                
                .AddNew
                
                .Fields("CodigoProveedor").Value = EsNulo(rsOrigen.Fields("Proveedor").Value)
                .Fields("Fecha").Value = EsNulo(rsOrigen.Fields("Fec_Credit").Value)
                .Fields("Remito").Value = Val(TraerDato("PFactura", "(NroInterno = " & EsNulo(rsOrigen.Fields("NRO_INT_DB").Value) & ")", "Remito"))
                '.Fields("idMedioPago").Value = ""
                .Fields("Importe").Value = EsNulo(rsOrigen.Fields("Importe").Value)
                .Fields("TipoMovimiento").Value = EsNulo(rsOrigen.Fields("MovimientC").Value)
                .Fields("NroInterno").Value = EsNulo(rsOrigen.Fields("Nro_Int_Cr").Value)
                
                .Update

                rsOrigen.MoveNext
                barra.Value = barra.Value + 1
            Loop
        End If
        
    End With
    
    sqlOrigen = ""
    sqlDestino = ""

    If rsOrigen.State = 1 Then
        rsOrigen.Close
        Set rsOrigen = Nothing
    End If

If Err Then GrabarLog "MigrarPagosSA", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub ConectarDB(vDB As String)
On Error Resume Next

    Set connDDBBMigrar = New ADODB.Connection
    
    With connDDBBMigrar
        .ConnectionString = pathDBMigrar(vDB)
        .Open
        If Not .State = 1 Then
            MsgBox Err.Description
            Exit Sub
        End If
    End With

If Err Then GrabarLog "ConectarDB", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub MigrarPagosPorNotaCSA(vtabla As String)
On Error Resume Next

    vtabla = Replace(vtabla, ".DBF", "")

    Dim rsDestino As New ADODB.Recordset
    Dim sqlDestino As String
    Dim vNroDebito As Long, vnroremito As Long

    sqlDestino = "SELECT * FROM PCuentascorrientes WHERE idMedioPago = 8 ORDER BY NroInterno ASC"
    
    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            Do Until .EOF = True
                DoEvents
                
                vNroDebito = TraerDato(vtabla, "NRO_INT_CR = " & Val(.Fields("NroInterno").Value) & "", "NRO_INT_DB", pathDBMigrar("ARV"))
                vnroremito = TraerDato("PFactura", "NroInterno = " & Val(vNroDebito) & "", "Remito")
                
                .Fields("Remito").Value = vnroremito
                

                .MoveNext
                barra.Value = barra.Value + 1
            Loop
        End If
        
    End With
    

    sqlDestino = ""

    If rsDestino.State = 1 Then
        rsDestino.Close
        Set rsDestino = Nothing
    End If

If Err Then GrabarLog "MigrarPagosPorNotaCSA", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub MigrarCobrosPorNotaCSA(vtabla As String)
On Error Resume Next

    vtabla = Replace(vtabla, ".DBF", "")

    Dim rsDestino As New ADODB.Recordset
    Dim sqlDestino As String
    Dim vNroDebito As Long, vnroremito As Long

    sqlDestino = "SELECT * FROM Cuentascorrientes WHERE idMedioPago = 8"

    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            Do Until .EOF = True
                DoEvents
                
                vNroDebito = TraerDato(vtabla, "NRO_INT_CR = " & Val(.Fields("NroInterno").Value) & "", "NRO_INT_DB", pathDBMigrar("ARV"))
                vnroremito = TraerDato("Factura", "NroInterno = " & Val(vNroDebito) & "", "Remito")
                
                .Fields("Remito").Value = vnroremito
                
                .Update

                barra.Value = barra.Value + 1
            Loop
        End If
        
    End With
    
    sqlDestino = ""

    If rsDestino.State = 1 Then
        rsDestino.Close
        Set rsDestino = Nothing
    End If

If Err Then GrabarLog "MigrarPagosPorNotaCSA", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub MigrarAsientosTipoSA(vtabla As String)
On Error Resume Next

    vtabla = Replace(vtabla, ".DBF", "")

    Dim rsOrigen As New ADODB.Recordset, rsDestino As New ADODB.Recordset
    Dim sqlOrigen As String, sqlDestino As String
    Dim vNroDebito As Long, vnroremito As Long

    sqlDestino = "SELECT * FROM AsientosTipo"
    sqlOrigen = "SELECT * FROM " & vtabla & " WHERE Empresa = '1' ORDER BY Asiento"
    
    With rsOrigen
        .CursorLocation = adUseClient
        Call .Open(sqlOrigen, connDDBBMigrar, adOpenStatic, adLockPessimistic)
        
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        Else
            .MoveFirst
            barra.Value = 0
            barra.Max = .RecordCount
        End If
        
    End With
    
    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            Do Until rsOrigen.EOF = True
                DoEvents
                
                .AddNew
                .Fields("Numero").Value = EsNulo(rsOrigen.Fields("ASIENTO").Value)
                .Fields("CodigoCuenta").Value = EsNulo(rsOrigen.Fields("CUENTA").Value)
                .Fields("DebeHaber").Value = EsNulo(rsOrigen.Fields("DEUDOR_ACR").Value)
                .Fields("Porcentaje").Value = EsNulo(rsOrigen.Fields("PORC_IMP").Value)
                
                .Update

                barra.Value = barra.Value + 1
                rsOrigen.MoveNext
            Loop
        End If
        
    End With
    
    sqlOrigen = ""
    sqlDestino = ""

    If rsDestino.State = 1 Then
        rsDestino.Close
        Set rsDestino = Nothing
    End If


    If rsOrigen.State = 1 Then
        rsOrigen.Close
        Set rsOrigen = Nothing
    End If

If Err Then GrabarLog "MigrarAsientosTipoSA", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub MigrarBancosProveedoresSA(vtabla As String)
On Error Resume Next
Err.Clear
    vtabla = Replace(vtabla, ".DBF", "")

    Dim rsOrigen As New ADODB.Recordset, rsDestino As New ADODB.Recordset
    Dim sqlOrigen As String, sqlDestino As String
    Dim vNroDebito As Long, vnroremito As Long, vIDBC As Long

    sqlDestino = "SELECT * FROM BancosMovimientos"
    sqlOrigen = "SELECT * FROM " & vtabla & ""
    
    With rsOrigen
        .CursorLocation = adUseClient
        Call .Open(sqlOrigen, connDDBBMigrar, adOpenStatic, adLockPessimistic)
        
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        Else
            .MoveFirst
            barra.Value = 0
            barra.Max = .RecordCount
        End If
        
    End With
    
    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
    
        If .State = 1 Then
        
            Do Until rsOrigen.EOF = True
                DoEvents
                
                .Close
                sqlDestino = ""
                sqlDestino = "SELECT * FROM BancosMovimientos WHERE (TipoMovimiento = '" & EsNulo(rsOrigen.Fields("Movimiento").Value) & "') AND (NroInterno = " & Val(rsOrigen.Fields("Nro_Inter").Value) & ") AND (NroCheque = " & Val(rsOrigen.Fields("Nro_Valor").Value) & ")  AND (idBancos = '" & EsNulo(rsOrigen.Fields("COD_BCO").Value) & "' OR idBancos = '" & EsNulo(rsOrigen.Fields("CAJA").Value) & "') AND (idTipoValor = '" & EsNulo(rsOrigen.Fields("Tipo_Valor").Value) & "')"
                
                
                Call .Open(sqlDestino, ConnDDBB, adOpenDynamic, adLockPessimistic)
                
                If .EOF = True Then
                    .AddNew
                
                    vIDBC = 0
                
                    If rsOrigen.Fields("Nro_Inter").Value = 4198 Then MsgBox rsOrigen.Fields("Nro_Inter").Value
                
                    Select Case EsNulo(rsOrigen.Fields("Nro_Cuenta").Value)
                
                        Case "11339/9"
                            vIDBC = 1
                        Case "1463-8 034-3"
                            vIDBC = 2
                        Case "01"
                            vIDBC = 3
                        Case "3-124-1995-9"
                            vIDBC = 4
                        Case "107-008409/3"
                            vIDBC = 5
                        Case "454-20-000439/9"
                            vIDBC = 6
                        Case "20000.207/37"
                            vIDBC = 7
                        Case "BBVA-TARJ.VISA"
                            vIDBC = 8
                        Case "CABAL AGRO"
                            'vIDBC = 9
                        Case "CABAL AGRO"
                            vIDBC = 10
                        Case "BCO.MACRO-VISA"
                            vIDBC = 11
                    End Select
                
                    If IsNull(rsOrigen.Fields("Cod_BCO").Value) = True Or EsNulo(rsOrigen.Fields("Cod_BCO").Value) = "" Then
                        .Fields("idBancos").Value = EsNulo(rsOrigen.Fields("CAJA").Value)
                    Else
                        .Fields("idBancos").Value = EsNulo(rsOrigen.Fields("Cod_BCO").Value)
                    End If
                
                    .Fields("idBancosCuentas").Value = Val(vIDBC)
                    
                    .Fields("Fecha").Value = strfechaMySQL(rsOrigen.Fields("Fec_Valor").Value)
                
                    .Fields("Credito").Value = Val(rsOrigen.Fields("Importe").Value)
                
                    .Fields("Saldo").Value = 0
                
                    '.Fields("Comentario").value = TraerDato("APCC00", "Nro_Inter = " & Val(rsOrigen.Fields("Nro_Inter").value) & "", "Leyenda", pathdbamigrar)
                
                    .Fields("NroCheque").Value = Val(rsOrigen.Fields("Nro_Valor").Value)
                    .Fields("TipoMovimiento").Value = EsNulo(rsOrigen.Fields("Movimiento").Value)
                    .Fields("NroInterno").Value = Val(rsOrigen.Fields("Nro_Inter").Value)
                    .Fields("idTipoValor").Value = EsNulo(rsOrigen.Fields("Tipo_Valor").Value)
                
                    .Fields("FechaValor").Value = strfechaMySQL(rsOrigen.Fields("Fec_Valor").Value)
                
                    If EsNulo(rsOrigen.Fields("PROVEEDOR").Value) = "" Or IsNull(EsNulo(rsOrigen.Fields("PROVEEDOR").Value)) = True Then
                        .Fields("Comentario").Value = EsNulo(rsOrigen.Fields("OBSERV").Value)
                    Else
                        .Fields("CP").Value = "P"
                        .Fields("ClienteProveedor").Value = EsNulo(rsOrigen.Fields("PROVEEDOR").Value)
                        .Fields("Comentario").Value = TraerDato("APCC00", "Nro_Inter = " & Val(rsOrigen.Fields("Nro_Inter").Value) & "", "Leyenda", pathDBMigrar("ARV"))
                    End If
                Me.log.AddItem (Str(.Fields("idBancosCuentas").Value) + Str(.Fields("Fecha").Value) + Str(.Fields("Credito").Value))
                    .Update
                Else
                    If Not IsNull(rsOrigen.Fields("CAJA").Value) = True Then
                        If Not IsNull(rsOrigen.Fields("PROVEEDOR").Value) = True Then
                            If Not rsOrigen.Fields("PROVEEDOR").Value = "" Then
                            
                                .AddNew
                    
                                vIDBC = 0
                    
                                
                                .Fields("idBancos").Value = EsNulo(rsOrigen.Fields("CAJA").Value)
                                .Fields("idBancosCuentas").Value = Val(vIDBC)
                                .Fields("Fecha").Value = strfechaMySQL(rsOrigen.Fields("Fec_Valor").Value)
                                .Fields("Credito").Value = Val(rsOrigen.Fields("Importe").Value)
                                .Fields("Saldo").Value = 0
                    
                                .Fields("NroCheque").Value = Val(rsOrigen.Fields("Nro_Valor").Value)
                                .Fields("TipoMovimiento").Value = EsNulo(rsOrigen.Fields("Movimiento").Value)
                                .Fields("NroInterno").Value = Val(rsOrigen.Fields("Nro_Inter").Value)
                                .Fields("idTipoValor").Value = EsNulo(rsOrigen.Fields("Tipo_Valor").Value)
                    
                                .Fields("FechaValor").Value = strfechaMySQL(rsOrigen.Fields("Fec_Valor").Value)
    
                                .Fields("CP").Value = "P"
                                .Fields("ClienteProveedor").Value = EsNulo(rsOrigen.Fields("PROVEEDOR").Value)
                                .Fields("Comentario").Value = TraerDato("APCC00", "Nro_Inter = " & Val(rsOrigen.Fields("Nro_Inter").Value) & "", "Leyenda", pathDBMigrar("ARV"))
                    
                                .Update
                        
                        
                            End If
                        End If
                    End If
                End If
                
                If barra.Max < barra.Value Then
                    MsgBox "error"
                End If
                
                barra.Value = barra.Value + 1
                rsOrigen.MoveNext
                
            Loop
        End If
        
    End With
    
    sqlOrigen = ""
    sqlDestino = ""

    If rsDestino.State = 1 Then
        rsDestino.Close
        Set rsDestino = Nothing
    End If


    If rsOrigen.State = 1 Then
        rsOrigen.Close
        Set rsOrigen = Nothing
    End If

If Err Then GrabarLog "MigrarBancosProveedoresSA", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub MigrarBancosClientesSA(vtabla As String)
On Error Resume Next
Err.Clear
    vtabla = Replace(vtabla, ".DBF", "")

    Dim rsOrigen As New ADODB.Recordset, rsDestino As New ADODB.Recordset
    Dim sqlOrigen As String, sqlDestino As String
    Dim vNroDebito As Long, vnroremito As Long, vIDBC As Long

    sqlDestino = "SELECT * FROM BancosMovimientos WHERE 1=2"
    sqlOrigen = "SELECT * FROM " & vtabla & ""
    
    With rsOrigen
        .CursorLocation = adUseClient
        Call .Open(sqlOrigen, connDDBBMigrar, adOpenStatic, adLockPessimistic)
        
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        Else
            .MoveFirst
            barra.Value = 0
            barra.Max = .RecordCount
        End If
        
    End With
    
    With rsDestino
        Call .Open(sqlDestino, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .State = 1 Then
        
            Do Until rsOrigen.EOF = True
                'If rsOrigen.Fields("Nro_Inter").value = 4198 Then MsgBox rsOrigen.Fields("Nro_Inter").value
                
                DoEvents
                
                .Close
                sqlDestino = ""
                sqlDestino = "SELECT * FROM BancosMovimientos WHERE (TipoMovimiento = '" & EsNulo(rsOrigen.Fields("Movimiento").Value) & "') AND (NroInterno = " & Val(rsOrigen.Fields("Nro_Inter").Value) & ") AND (NroCheque = " & Val(rsOrigen.Fields("Nro_Valor").Value) & ")  AND (idBancos = '" & EsNulo(rsOrigen.Fields("COD_BCO").Value) & "' or idBancos = '" & EsNulo(rsOrigen.Fields("CAJA").Value) & "') AND (idTipoValor = '" & EsNulo(rsOrigen.Fields("Tipo_Valor").Value) & "')"
                
                Call .Open(sqlDestino, ConnDDBB, adOpenDynamic, adLockPessimistic)
                
                If .EOF = True Then
                
                    .AddNew
                
                    vIDBC = 0
                
                    Select Case EsNulo(rsOrigen.Fields("Nro_Cuenta").Value)
                
                        Case "11339/9"
                            vIDBC = 1
                        Case "1463-8 034-3"
                            vIDBC = 2
                        Case "01"
                            vIDBC = 3
                        Case "3-124-1995-9"
                            vIDBC = 4
                        Case "107-008409/3"
                            vIDBC = 5
                        Case "454-20-000439/9"
                            vIDBC = 6
                        Case "20000.207/37"
                            vIDBC = 7
                        Case "BBVA-TARJ.VISA"
                            vIDBC = 8
                        Case "CABAL AGRO"
                            'vIDBC = 9
                        Case "CABAL AGRO"
                            vIDBC = 10
                        Case "BCO.MACRO-VISA"
                            vIDBC = 11
                    End Select

                    If IsNull(rsOrigen.Fields("Cod_BCO").Value) = True Or EsNulo(rsOrigen.Fields("Cod_BCO").Value) = "" Then
                        .Fields("idBancos").Value = EsNulo(rsOrigen.Fields("CAJA").Value)
                    Else
                        .Fields("idBancos").Value = EsNulo(rsOrigen.Fields("Cod_BCO").Value)
                    End If
                    
                    .Fields("idBancosCuentas").Value = Val(vIDBC)
                    .Fields("Fecha").Value = strfechaMySQL(rsOrigen.Fields("Fec_Valor").Value)
                    .Fields("Debito").Value = Val(rsOrigen.Fields("Importe").Value)
                    .Fields("Saldo").Value = 0
                
                    .Fields("NroCheque").Value = Val(rsOrigen.Fields("Nro_Valor").Value)
                    .Fields("TipoMovimiento").Value = EsNulo(rsOrigen.Fields("Movimiento").Value)
                    .Fields("NroInterno").Value = Val(rsOrigen.Fields("Nro_Inter").Value)
                    .Fields("idTipoValor").Value = EsNulo(rsOrigen.Fields("Tipo_Valor").Value)
                
                    .Fields("FechaValor").Value = strfechaMySQL(rsOrigen.Fields("Fec_Valor").Value)
                
                    If EsNulo(rsOrigen.Fields("CLIENTE").Value) = "" Or IsNull(EsNulo(rsOrigen.Fields("CLIENTE").Value)) = True Then
                        .Fields("Comentario").Value = EsNulo(rsOrigen.Fields("OBSERV").Value)
                    Else
                        .Fields("CP").Value = "C"
                        .Fields("ClienteProveedor").Value = EsNulo(rsOrigen.Fields("Cliente").Value)
                        .Fields("Comentario").Value = TraerDato("AVCC00", "Nro_Inter = " & Val(rsOrigen.Fields("Nro_Inter").Value) & "", "Leyenda", pathDBMigrar("ARV"))
                    End If
                
                    .Update
                    
                Else
                    If Not IsNull(rsOrigen.Fields("CAJA").Value) = True Then
                        If Not IsNull(rsOrigen.Fields("CLIENTE").Value) = True Then
                            If Not rsOrigen.Fields("CLIENTE").Value = "" Then
                            
                                .AddNew
                    
                                vIDBC = 0
                    
                                
                                .Fields("idBancos").Value = EsNulo(rsOrigen.Fields("CAJA").Value)
                                .Fields("idBancosCuentas").Value = Val(vIDBC)
                                .Fields("Fecha").Value = strfechaMySQL(rsOrigen.Fields("Fec_Valor").Value)
                                .Fields("Debito").Value = Val(rsOrigen.Fields("Importe").Value)
                                .Fields("Saldo").Value = 0
                    
                                .Fields("NroCheque").Value = Val(rsOrigen.Fields("Nro_Valor").Value)
                                .Fields("TipoMovimiento").Value = EsNulo(rsOrigen.Fields("Movimiento").Value)
                                .Fields("NroInterno").Value = Val(rsOrigen.Fields("Nro_Inter").Value)
                                .Fields("idTipoValor").Value = EsNulo(rsOrigen.Fields("Tipo_Valor").Value)
                    
                                .Fields("FechaValor").Value = strfechaMySQL(rsOrigen.Fields("Fec_Valor").Value)
    
                                .Fields("CP").Value = "C"
                                .Fields("ClienteProveedor").Value = EsNulo(rsOrigen.Fields("Cliente").Value)
                                .Fields("Comentario").Value = TraerDato("AVCC00", "Nro_Inter = " & Val(rsOrigen.Fields("Nro_Inter").Value) & "", "Leyenda", pathDBMigrar("ARV"))
                    
                                .Update
                        
                        
                            End If
                        End If
                    End If
                End If
               
                barra.Value = barra.Value + 1
                rsOrigen.MoveNext
            
            Loop
        End If
        
    End With
    
    sqlOrigen = ""
    sqlDestino = ""

    If rsDestino.State = 1 Then
        rsDestino.Close
        Set rsDestino = Nothing
    End If


    If rsOrigen.State = 1 Then
        rsOrigen.Close
        Set rsOrigen = Nothing
    End If

If Err Then GrabarLog "MigrarBancosClientesSA", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub
