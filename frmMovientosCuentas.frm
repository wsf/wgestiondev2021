VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{9746E3DA-06E1-4D26-9CE4-D9F6411A9C70}#1.0#0"; "SMGA_OcxTxt2008.ocx"
Begin VB.Form frmMovimientosCuentas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos Efectuados por Cuentas"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   9525
   Begin XtremeSuiteControls.GroupBox GroupBox4 
      Height          =   570
      Left            =   6390
      TabIndex        =   40
      Top             =   2460
      Width           =   3030
      _Version        =   851968
      _ExtentX        =   5345
      _ExtentY        =   1005
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton PusImprimir 
         Height          =   345
         Left            =   2010
         TabIndex        =   43
         Top             =   180
         Width           =   915
         _Version        =   851968
         _ExtentX        =   1614
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton rbBalanceI 
         Height          =   255
         Left            =   60
         TabIndex        =   41
         Top             =   210
         Width           =   1005
         _Version        =   851968
         _ExtentX        =   1773
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ingresos"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton rbBalanceE 
         Height          =   255
         Left            =   1110
         TabIndex        =   42
         Top             =   210
         Width           =   1005
         _Version        =   851968
         _ExtentX        =   1773
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Egresos"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   555
      Left            =   135
      TabIndex        =   33
      Top             =   2430
      Width           =   6225
      _Version        =   851968
      _ExtentX        =   10980
      _ExtentY        =   979
      _StockProps     =   79
      Caption         =   "Tipos de movimientos"
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox vmarcapropia 
         Height          =   225
         Left            =   150
         TabIndex        =   37
         Top             =   240
         Width           =   1515
      End
      Begin XtremeSuiteControls.RadioButton rdInterno 
         Height          =   195
         Left            =   3270
         TabIndex        =   34
         Top             =   240
         Width           =   1485
         _Version        =   851968
         _ExtentX        =   2619
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Asiento interno"
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton rdNormal 
         Height          =   255
         Left            =   4860
         TabIndex        =   35
         Top             =   210
         Width           =   1485
         _Version        =   851968
         _ExtentX        =   2619
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Asiento normal"
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton rdtodos 
         Height          =   165
         Left            =   1710
         TabIndex        =   36
         Top             =   240
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   291
         _StockProps     =   79
         Caption         =   "Todos los asientos"
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   555
      Left            =   -60
      TabIndex        =   26
      Top             =   -60
      Width           =   9495
      _Version        =   851968
      _ExtentX        =   16748
      _ExtentY        =   979
      _StockProps     =   79
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   120
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Vista Previa"
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Picture         =   "frmMovientosCuentas.frx":0000
         BorderGap       =   10
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   375
         Index           =   2
         Left            =   8070
         TabIndex        =   28
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cerrar"
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Picture         =   "frmMovientosCuentas.frx":6862
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   375
         Index           =   1
         Left            =   1620
         TabIndex        =   29
         Top             =   120
         Visible         =   0   'False
         Width           =   1425
         _Version        =   851968
         _ExtentX        =   2514
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Imprimir"
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Picture         =   "frmMovientosCuentas.frx":6C62
         BorderGap       =   10
      End
   End
   Begin VB.ListBox log 
      Height          =   3180
      Left            =   60
      TabIndex        =   24
      Top             =   4770
      Width           =   9285
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   525
      Left            =   60
      TabIndex        =   18
      Top             =   3150
      Width           =   9345
      _Version        =   851968
      _ExtentX        =   16484
      _ExtentY        =   926
      _StockProps     =   79
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit vleyenda 
         Height          =   315
         Left            =   3420
         TabIndex        =   20
         Top             =   180
         Width           =   5835
         _Version        =   851968
         _ExtentX        =   10292
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vnumero 
         Height          =   315
         Left            =   1290
         TabIndex        =   19
         Top             =   180
         Visible         =   0   'False
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1940
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "25218"
         Alignment       =   1
      End
      Begin VB.Label lblLeyenda 
         Caption         =   "Leyenda:"
         Height          =   255
         Left            =   2520
         TabIndex        =   22
         Top             =   240
         Width           =   825
      End
      Begin VB.Label lblNroAsiento 
         Caption         =   "IdAsiento"
         Height          =   255
         Left            =   60
         TabIndex        =   21
         Top             =   210
         Visible         =   0   'False
         Width           =   1125
      End
   End
   Begin WGestion.LedProgress Barra 
      Height          =   135
      Left            =   60
      TabIndex        =   17
      Top             =   4590
      Width           =   9195
      _extentx        =   16219
      _extenty        =   238
   End
   Begin XtremeSuiteControls.GroupBox GBOtros 
      Height          =   435
      Left            =   60
      TabIndex        =   13
      Top             =   3720
      Width           =   9315
      _Version        =   851968
      _ExtentX        =   16431
      _ExtentY        =   767
      _StockProps     =   79
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin VB.CheckBox chkBalanceIE 
         Caption         =   "balanceIE"
         Height          =   255
         Left            =   7170
         TabIndex        =   39
         Top             =   120
         Width           =   1185
      End
      Begin VB.TextBox vclave 
         Height          =   225
         Left            =   5460
         TabIndex        =   25
         Top             =   180
         Width           =   1425
      End
      Begin VB.CheckBox chkAgrupado 
         Caption         =   "Agrupar detalles de Asientos"
         Height          =   195
         Left            =   2910
         TabIndex        =   14
         Top             =   180
         Width           =   2325
      End
      Begin VB.CheckBox chkCuentasVacias 
         Caption         =   "Mostrar Cuentas sin Movientos"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   180
         Width           =   2985
      End
   End
   Begin XtremeSuiteControls.GroupBox GBParametros 
      Height          =   1905
      Left            =   30
      TabIndex        =   0
      Top             =   420
      Width           =   9375
      _Version        =   851968
      _ExtentX        =   16536
      _ExtentY        =   3360
      _StockProps     =   79
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton PusMesActual 
         Height          =   675
         Left            =   4140
         TabIndex        =   44
         Top             =   990
         Width           =   675
         _Version        =   851968
         _ExtentX        =   1191
         _ExtentY        =   1191
         _StockProps     =   79
         Caption         =   "Mes Actual"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.CheckBox chkVarios 
         Caption         =   "Varios ejercicios juntos"
         Height          =   255
         Left            =   7170
         TabIndex        =   38
         Top             =   1020
         Width           =   1995
      End
      Begin VB.CheckBox chkFechas 
         Caption         =   "Todas las Fechas"
         Height          =   255
         Left            =   5160
         TabIndex        =   16
         Top             =   1050
         Width           =   1575
      End
      Begin XtremeSuiteControls.PushButton pbContabilidad 
         Height          =   315
         Index           =   0
         Left            =   4080
         TabIndex        =   7
         Tag             =   "CodigoCuentaD"
         Top             =   240
         Width           =   345
         _Version        =   851968
         _ExtentX        =   609
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit vnro_asto0 
         Height          =   315
         Left            =   1530
         TabIndex        =   8
         Top             =   240
         Width           =   2505
         _Version        =   851968
         _ExtentX        =   4419
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vnro_asto1 
         Height          =   315
         Left            =   4500
         TabIndex        =   9
         Top             =   210
         Width           =   4695
         _Version        =   851968
         _ExtentX        =   8281
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vnro_asto2 
         Height          =   315
         Left            =   1530
         TabIndex        =   10
         Top             =   600
         Width           =   2505
         _Version        =   851968
         _ExtentX        =   4419
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbContabilidad 
         Height          =   285
         Index           =   1
         Left            =   4080
         TabIndex        =   11
         Tag             =   "CodigoCuentaH"
         Top             =   600
         Width           =   345
         _Version        =   851968
         _ExtentX        =   609
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "..."
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit vnro_asto3 
         Height          =   315
         Left            =   4500
         TabIndex        =   12
         Top             =   570
         Width           =   4665
         _Version        =   851968
         _ExtentX        =   8229
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   285
         Left            =   8700
         TabIndex        =   30
         Top             =   1410
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "..."
         BackColor       =   -2147483633
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit vcbalance 
         Height          =   285
         Left            =   7470
         TabIndex        =   31
         Top             =   1410
         Width           =   1065
         _Version        =   851968
         _ExtentX        =   1879
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
      End
      Begin Aplisoft_CajasDeTexto.TxF dtpCuentas 
         Height          =   300
         Index           =   0
         Left            =   1530
         TabIndex        =   5
         Top             =   1020
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   0
      End
      Begin Aplisoft_CajasDeTexto.TxF dtpCuentas 
         Height          =   300
         Index           =   1
         Left            =   1530
         TabIndex        =   6
         Top             =   1410
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   0
      End
      Begin VB.Label lblBalance 
         BackStyle       =   0  'Transparent
         Caption         =   "> Nombre del Balance:"
         Height          =   195
         Index           =   6
         Left            =   5460
         TabIndex        =   32
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Cuenta Hasta :"
         Height          =   195
         Index           =   1
         Left            =   30
         TabIndex        =   4
         Top             =   600
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Cuenta Desde :"
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   3
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> Fecha Desde:"
         Height          =   195
         Index           =   2
         Left            =   30
         TabIndex        =   2
         Top             =   1080
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> Fecha Hasta:"
         Height          =   195
         Index           =   3
         Left            =   30
         TabIndex        =   1
         Top             =   1440
         Width           =   1335
      End
   End
   Begin VB.Label vmensaje 
      BackColor       =   &H80000004&
      Caption         =   "....."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   30
      TabIndex        =   23
      Top             =   4260
      Width           =   9345
   End
End
Attribute VB_Name = "frmMovimientosCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sqlA, sqlC, vcodigoBalance As String
Dim vTotalD, vTotalH As Double
Dim vsaldoactual As Double
Dim vsaldoanterior, vgneto As Double
Dim vnrocomprobantegral As Long

Private Sub chkFechas_Click()
On Error Resume Next
    
    dtpCuentas(0).Enabled = Not CBool(chkFechas.Value)
    dtpCuentas(1).Enabled = Not CBool(chkFechas.Value)

If Err Then GrabarLog "chkFechas_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub GenerarDatos()
On Error Resume Next

    Dim rsCuentas As New ADODB.Recordset, sqlCuentas As String
    
    Dim vnrobalance As Integer
    
    If Not (Val(vnro_asto0.Text) = 0) And Not (Val(vnro_asto2.Text) = 0) Then
        sqlC = sqlC + " AND (CodigoCuenta >= '" & vnro_asto0.Text & "' AND CodigoCuenta <= '" & vnro_asto2.Text & "')"
    End If
    
    sqlCuentas = "SELECT * FROM cuentas WHERE 1=1 " + sqlC + " ORDER BY CodigoCuenta ASC"

    With rsCuentas
        Call .Open(sqlCuentas, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then
            .MoveFirst
            'Barra.Value = 0
            'Barra.Max = .RecordCount
        End If
        
        Dim vCodCuenta As String, vcuenta As String
        
        
        vnrobalance = traerDatos2("select nrobalance from balances where Activo='S' order by nrobalance desc", "nrobalance", pathDBMySQL)
   
        vnrobalance = selectNrobalance(dtpCuentas(0).Value, dtpCuentas(1).Value, vnrobalance)
        
        
        Do Until .EOF = True
            'DoEvents
            
            vCodCuenta = .Fields("CodigoCuenta").Value
            vcuenta = .Fields("Cuenta").Value
            
            vmensaje.Caption = "Cuenta a procesar : " + vcuenta
            
            
            Call BuscarMovimientos(vCodCuenta, vcuenta, vnrobalance) ' paso2
            
            barra.Value = .AbsolutePosition
            .MoveNext
        Loop
    
    End With
    
    sqlCuentas = ""
    
    If rsCuentas.State = 1 Then
        rsCuentas.Close
        Set rsCuentas = Nothing
    End If
    
If Err Then GrabarLog "GenerarDatos", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub MostrarReporte(Index As Integer)
On Error Resume Next

    Unload Mantenimiento
    Load Mantenimiento

    MsgBox "Prepare la Impresora!!!", vbInformation, "Mensaje ..."
    
    Unload Mantenimiento
    Load Mantenimiento
        
    With Mantenimiento.rsMCuentas
        If Not chkAgrupado.Value = 1 Then
            '.Source = "SHAPE {SELECT * FROM TempCuentas ORDER BY codigo ASC} AS MCuentas APPEND ({SELECT * FROM Temp2 ORDER BY C10 ASC, C02 ASC, C01 ASC}  AS TempCuentas RELATE 'Codigo' TO 'C10') AS TempCuentas"
            .Source = "SHAPE {SELECT * FROM TempCuentas ORDER BY codigo ASC} AS MCuentas APPEND ({SELECT * FROM Temp2}  AS TempCuentas RELATE 'Codigo' TO 'C10') AS TempCuentas"
        Else
           ' .Source = "SHAPE {SELECT * FROM TempCuentas ORDER BY codigo ASC} AS MCuentas APPEND ({SELECT Temp2.C09, Temp2.C11, Temp2.C01, Max(Temp2.C02) as C02, Sum(Temp2.C03) AS C03, Sum(Temp2.C04) AS C04, Temp2.C06, Temp2.C10, Max(Temp2.Id) AS Id FROM Temp2 GROUP BY Temp2.C01, Temp2.C06, Temp2.C10 ORDER BY C10 ASC, C01 ASC}  AS TempCuentas RELATE 'Codigo' TO 'C10') AS TempCuentas"
             .Source = "SHAPE {SELECT * FROM TempCuentas ORDER BY codigo ASC} AS MCuentas APPEND ({SELECT Temp2.C09, Temp2.C11, Temp2.C01, Max(Temp2.C02) as C02, Sum(Temp2.C03) AS C03, Sum(Temp2.C04) AS C04, Temp2.C06, Temp2.C10, Max(Temp2.Id) AS Id FROM Temp2 GROUP BY Temp2.C01, Temp2.C06, Temp2.C10}  AS TempCuentas RELATE 'Codigo' TO 'C10') AS TempCuentas"
  
        
        End If

        If .State = 0 Then .Open
        .Close
        .Open
    
        If .RecordCount = 0 Then
            MsgBox "No existen datos para mostrarse!!", vbInformation, "Mensaje ..."
            Exit Sub
        End If
    End With
    
    With drMovimientosCuentas
        .Sections(2).Controls("lblTitulo").Caption = "[ Movimientos Efectuados por Cuenta desde el Código:  " & Val(vnro_asto0.Text) & " Hasta : " & Val(vnro_asto2.Text) & " ]"
        .Sections(2).Controls("gnombre").Caption = vDatosEmpresa.Nombre
        .Sections(2).Controls("gdireccion").Caption = vDatosEmpresa.Direccion & "  /  " & vDatosEmpresa.Telefono
        .Sections(2).Controls("gtelefono").Caption = vDatosEmpresa.Telefono
        '.Sections(2).Controls("semail").Caption = vDatosEmpresa.Email
        If Index = 0 Then
            .Show
        Else
            .Hide
            Call .PrintReport(False, rptRangeAllPages)
        End If
    End With
    
If Err Then GrabarLog "MostrarReporte", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub BuscarMovimientos(ByRef vCodCuenta As String, ByRef vcuenta As String, vnrobalance As Integer)
On Error Resume Next
    Dim vsqlmarca, vmarca As String
    
    Dim vCodigoCliente, vCodigoProveedor, sqlauxiliar As String
    'Dim vnrobalance As Integer
    Dim rsAsientos As New ADODB.Recordset, sqlAsientos As String
    Dim vneto As Double
    
   
    
   ' vnrobalance = 0
    vTotalD = 0
    vTotalH = 0
    
    sqlA = ""
    
    sqlA = sqlA + " AND (CodigoCuenta = '" & vCodCuenta & "')"
    
    If Not chkFechas.Value = 1 Then
        sqlA = sqlA + " AND (Fecha >= '" & strfechaMySQL(dtpCuentas(0).Value) + "' AND Fecha <= '" & strfechaMySQL(dtpCuentas(1).Value) + "')"
    End If
    
    '13-08-2012
    
    
  ' vnrobalance = traerDatos2("select nrobalance from balances where Activo='S' order by nrobalance desc", "nrobalance", pathDBMySQL)
   
   
   '------------------------------------
   ' vnrobalance = selectNrobalance(dtpCuentas(0).Value, dtpCuentas(1).Value, vnrobalance)
    If vnrobalance = 0 And chkvarios.Value = 0 Then Exit Sub
    
    vcodigoBalance = traerDatos2("select * from balances where nrobalance=" + Str(vnrobalance), "codigo", pathDBMySQL)
    Me.vcbalance = vcodigoBalance
   '-------------------------------------
   
   vsqlmarca = fsqlmarca ' arma el where para las marcas
   vmarca = fmarca       ' tipo de marca
   
   
   If vnrobalance = 0 Then
    'sqlAsientos = "SELECT * FROM AsientosDetalle INNER JOIN Asientos ON Asientos.Numero = AsientosDetalle.Numero  WHERE 1=1 and  AsientosDetalle.nrobalance is null " + sqlA + " ORDER BY Asientos.Fecha ASC" ' moificado 10-06-2011
    
    sqlauxiliar = "and (((asientos.timestamp) <= (asientosdetalle.timestamp)) or (asientosdetalle.idAsientosDetalle >=87246))"
    'sqlauxiliar = "and (((asientos.timestamp) <= (asientosdetalle.timestamp)) or (asientos.timestamp < '2012-08-08'))"
    sqlAsientos = "SELECT * FROM AsientosDetalle INNER JOIN Asientos ON Asientos.Numero = AsientosDetalle.Numero  WHERE 1=1 " + sqlauxiliar + sqlA + " ORDER BY Asientos.Fecha ASC" ' modificado 23-10-2012 con el propósito
   
   Else
   
    If Me.chkvarios Then
        sqlAsientos = "SELECT * FROM AsientosDetalle INNER JOIN Asientos ON Asientos.Numero = AsientosDetalle.Numero  WHERE 1=1 and  AsientosDetalle.nrobalance =Asientos.nrobalance " + sqlA + vsqlmarca + " ORDER BY Asientos.Fecha ASC"    ' nrobalance = 16
    Else
        sqlAsientos = "SELECT * FROM AsientosDetalle INNER JOIN Asientos ON Asientos.Numero = AsientosDetalle.Numero  WHERE 1=1 and  AsientosDetalle.nrobalance =Asientos.nrobalance " + sqlA + vsqlmarca + " and  Asientos.nrobalance= " + Str(vnrobalance) + " ORDER BY Asientos.Fecha ASC"    ' nrobalance = 16
    End If
    ' sqlAsientos = "SELECT * FROM AsientosDetalle INNER JOIN Asientos ON Asientos.Numero = AsientosDetalle.Numero  WHERE 1=1 and  AsientosDetalle.nrobalance =Asientos.nrobalance " + sqlA + vsqlmarca + " ORDER BY Asientos.Fecha ASC"    ' nrobalance = 16

   End If
    
    ' guardo en temp_cuenta los datos del agrupamiento y calculo el saldo anterior que se tomará para el neto.
    ' aca se calcula el saldo actual para poner en la cuenta
    'Call GrabarCuentaTemp(vCodCuenta, vCuenta) ' paso3
    
  '  vneto = vsaldoanterior ' paso 3: calcula saldo anterior
     vneto = 0  ' pongo el neto en cero porque los saldos parciales son del periodo que pido. No tiene que arrastrar el saldo anterior.
     
    With rsAsientos
        vmensaje.Caption = "Filtrando los movimientos de la cuenta: " + vCodCuenta
        Call .Open(sqlAsientos, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        Debug.Print sqlAsientos
        
        If Not .EOF = True Then
            .MoveFirst
        
            ' acà estaba
        
            Dim vcp As String
            
            Do Until .EOF = True
                DoEvents
                vTotalD = 0
                vTotalH = 0
                
                vTotalD = vTotalD + Val(Format(.Fields("Debe").Value, "#######0.000"))
                vTotalH = vTotalH + Val(Format(.Fields("Haber").Value, "#######0.000"))
                
                vneto = vneto + vTotalD - vTotalH
                
                ' guardo en temp2 fecha, numero, d,c
                vmensaje.Caption = "Generando del movimiento:" + EsNulo(.Fields("Leyenda").Value)
                
                
                If vclave.Text = "dalascp" Then ' esto es un arreglo de los cli y proveedores
                
                    vCodigoCliente = traerDatos2("select * from cuentascorrientes where nrointerno=" + EsNulo(.Fields("nrointerno").Value), "Codigo", pathDBMySQL)
                    
                    vCodigoProveedor = traerDatos2("select * from pcuentascorrientes where nrointerno=" + EsNulo(.Fields("nrointerno").Value), "Codigo", pathDBMySQL)



                    If Not vCodigoCliente = "" Then
                        Call EjecutarScript("update asientos set codigoCliente='" + vCodigoCliente + "'" + " where nrointerno=" + EsNulo(.Fields("nrointerno").Value), pathDBMySQL)
                        log.AddItem ("NroAsiento: " + EsNulo(.Fields("Numero").Value) + "  NroInterno: " + EsNulo(.Fields("nrointerno").Value) + " Cliente: " + vCodigoCliente)
                    End If
                    
                    
                    If Not vCodigoCliente = "" Then
                        Call EjecutarScript("update asientos set codigoProveedor='" + vCodigoProveedor + "'" + " where nrointerno=" + EsNulo(.Fields("nrointerno").Value), pathDBMySQL)
                        log.AddItem ("NroAsiento: " + EsNulo(.Fields("Numero").Value) + "  NroInterno: " + EsNulo(.Fields("nrointerno").Value) + " Cliente: " + vCodigoProveedor)
                    End If
                
                
                End If
                
                
                vcp = fvCP(EsNulo(.Fields("CodigoProveedor").Value) + EsNulo(.Fields("CodigoCliente").Value), EsNulo(.Fields("NroInterno").Value))
                
                vgneto = 0
                vgneto = vneto ' guardo el ultimoneto en una variable global
               ' Call GuardarTemp(.Fields("Numero").Value, vCP + " - " + EsNulo(.Fields("Leyenda").Value) + fmarcaImpreso(EsNulo(.Fields("marca").Value)), .Fields("Fecha").Value, vCodCuenta, vneto, EsNulo(.Fields("tipoMovimiento").Value), EsNulo(.Fields("NroInterno").Value)) ' paso 4:
                 Call GuardarTemp(.Fields("Numero").Value, vcp + " - " + EsNulo(.Fields("Leyenda").Value) + fmarcaImpreso(EsNulo(.Fields("marca").Value)), .Fields("Fecha").Value, vCodCuenta, vneto, TraerTipoMovimiento(.Fields("nrointerno")), EsNulo(.Fields("NroInterno").Value)) ' paso 4:
                
                .MoveNext
            Loop
            
         Call GrabarCuentaTemp(vCodCuenta, vcuenta, vnrobalance, vmarca) ' paso3
             ' cuardo en temp_cuenta los datos del agrupamiento y calculo el saldo anterior que se tomará para el neto.
            'Call GrabarCuentaTemp(vCodCuenta, vCuenta)

            
        Else
            If (chkCuentasVacias.Value = 1) Then
                Call GrabarCuentaTemp(vCodCuenta, vcuenta, vnrobalance, vmarca) ' panic: ojo con esto para ver si tiene sentido
            End If
        End If
        
    End With
    
    sqlAsientos = ""
    
    If rsAsientos.State = 1 Then
        rsAsientos.Close
        Set rsAsientos = Nothing
    End If
    
If Err Then GrabarLog "BuscarMovimientos", Err.Number & " " & Err.Description, Me.Name
End Sub
Function fsqlmarca() As String
If Me.rdinterno.Value Then fsqlmarca = " and (marca = 'INTERNO')"
If Me.rdNormal.Value Then fsqlmarca = " and (marca = 'NORMAL')"
If Me.rdtodos Then fsqlmarca = ""
End Function
Function fmarca() As String
If Me.rdinterno.Value Then fmarca = "INTERNO"
If Me.rdNormal.Value Then fmarca = "NORMAL"
If Me.rdtodos Then fmarca = "TODOS"
End Function
Function fmarcaImpreso(vmarca As String) As String
If vmarca = "INTERNO" Then
    fmarcaImpreso = vmarcapropia
Else
    fmarcaImpreso = ""
End If
End Function


Function fvCP(vstr As String, vnrointerno As String) As String



If vstr = "" Then
    fvCP = "[ " + traerDatos2("select * from cuentascorrientes where NroInterno=" + vnrointerno, "Codigo", pathDBMySQL) + traerDatos2("select * from pcuentascorrientes where NroInterno=" + vnrointerno, "Codigo", pathDBMySQL) + " ]"
Else
    fvCP = "[ " + vstr + " ]"
End If

End Function

Private Sub GuardarTemp(vnumero, vDetalle, vfechaAsiento, ByRef vCodCuenta, vvsaldo As Double, vtipo As String, vinterno As String)
On Error Resume Next

    Dim rstemp As New ADODB.Recordset, sqlTemp As String
            
    sqlTemp = "SELECT * FROM Temp2"
    
    With rstemp
        Call .Open(sqlTemp, ConnDDBB, adOpenDynamic, adLockPessimistic)
                
          .AddNew
            
        .Fields("C01").Value = vfechaAsiento
        '.Fields("C02").Value = vnumero 'vnrocomprobantegral
        .Fields("C03").Value = vTotalD
        .Fields("C04").Value = vTotalH
        
        Select Case (vTotalD - vTotalH)
        
            Case Is > 0
                .Fields("C07").Value = Val(Format(vTotalD - vTotalH, "#######0.000"))
                .Fields("C11").Value = 0
            
            Case Is < 0
                .Fields("C07").Value = 0
                .Fields("C07").Value = Val(Format(vTotalH - vTotalD, "#######0.000"))
            
            Case 0
                .Fields("C07").Value = 0
                .Fields("C07").Value = 0
                
        End Select
        
       ' If Trim(vDetalle) = "[  ] -" Then vDetalle = addDetalleBM(Val(vinterno), asientoN2F(vnumero))
        
        vnrocomprobantegral = 0
       '
       ' If Not UCase(LeerXml("Puesto")) = "POLIWHEEL" Then vDetalle = addDetalleBM(Val(vinterno))  ' sacarlo
       
       vDetalle = addDetalleBM(Val(vinterno))
        
        .Fields("C02").Value = fvnrocomprobantegral(vnumero, Val(vinterno)) ' le paso el nro de asiento para que lo devuelva si no hay comprobante
         
         Debug.Print "Comentario mayor : " + vDetalle
         
        .Fields("C06").Value = Trim(Left(vDetalle, 254))
        .Fields("C09").Value = Trim(vtipo)
        
        .Fields("C11").Value = vvsaldo
            
        .Fields("C10").Value = vCodCuenta
        
        .Fields("C13").Value = vinterno
        
        .Update

    End With
    
    sqlTemp = ""
    
    If rstemp.State = 1 Then
        rstemp.Close
        Set rstemp = Nothing
    End If
    
If Err < 0 Then GrabarLog "GuardarTemp", Err.Number & " " & Err.Description, Me.Name
End Sub

Function fvnrocomprobantegral(ByVal vnro As Long, vinterno As Long)
On Error Resume Next
Dim vsql As String

vsql = "select nrocomprobante as c from bancosmovimientos where nrointerno = " + Str(vinterno)

fvnrocomprobantegral = traerDatos2(vsql, "c", pathDBMySQL)


If Err Then
    fvnrocomprobantegral = vnro
End If
End Function


Function asientoN2F(ByVal vnumero As Long) As Date
On Error Resume Next
Dim vsql As String

vsql = "select fecha as c from asientos where numero = " + Str(vnumero)

asientoN2F = traerDatos2(vsql, "c", pathDBMySQL)

If Err Then asientoN2F = CDate(("2010-01-01"))
End Function

Function addDetalleBM(vinterno As Long) As String
On Error Resume Next

Dim vsql, va, valor, vsqlte As String

'vsqlte = " and fecha = '" + strfechaMySQL(vfecha) + "'"


valor = ""



vsql = "select comentario2 as c from bancosmovimientos where nrointerno = " + Str(vinterno)

va = traerDatos2(vsql, "c", pathDBMySQL)

valor = valor + va

vsql = "select comentario as c from bancosmovimientos where nrointerno = " + Str(vinterno)

va = traerDatos2(vsql, "c", pathDBMySQL)

valor = valor + va

vsql = "select codPersona as c from bancosmovimientos where nrointerno = " + Str(vinterno)
va = traerDatos2(vsql, "c", pathDBMySQL)

vsql = "select nombre as c from proveedores where codigo = " + Str(va)

va = traerDatos2(vsql, "c", pathDBMySQL)

valor = valor + "-Interesado :  " + va
' addDetalleBM = valor


vsql = "select Leyenda as c from asientos t where t.NroInterno = " + Str(vinterno)

va = traerDatos2(vsql, "c", pathDBMySQL)

valor = valor + " [" + Right(va, Len(va) - 2) + "]"

addDetalleBM = valor

If Err Then
    'addDetalleBM = ""
    Exit Function
End If

End Function

Function CalcularSaldo(ByRef vCodCuenta As String, vAnterior As Boolean, vfechaHasta As Date) As Double
On Error Resume Next

vmensaje.Caption = "Calculando saldo de la cuenta: " + vCodCuenta

Dim vsqlanterior, vwhereAnterior, vwhereActual As String


' condicionales para el saldo anterior
vwhereAnterior = "(asientosdetalle.CodigoCuenta ='" + vCodCuenta & "' AND " & _
                  " (asientos.Fecha < '" & strfechaMySQL(vfechaHasta) & "')"
                  'and " & _
                  '" asientos.idAsientos >=" & Me.vnumero.Text & " )"
                  
' condicionales para saldo actual
vwhereActual = "asientosdetalle.CodigoCuenta ='" + vCodCuenta & "'  and " & _
                  " asientos.idAsientos >=" & Me.vnumero.Text & " "

    Dim rsSaldo As New ADODB.Recordset, sqlSaldo As String

    If vAnterior = True Then ' si tiene saldo anterior ' calculo del saldo anterior
    
vmensaje.Caption = "Calculando saldo anterior de la cuenta: " + vCodCuenta

  sqlSaldo = "SELECT  sum(asientosdetalle.Debe) as d, sum(asientosdetalle.Haber) As H, (sum(asientosdetalle.Debe) - sum(asientosdetalle.Haber)) as Saldo " & _
                 " From  asientos " & _
                 " INNER JOIN asientosdetalle ON (asientos.Numero=asientosdetalle.Numero) " & _
                 " INNER JOIN cuentas ON (asientosdetalle.CodigoCuenta=cuentas.CodigoCuenta) " & _
                 " Where " & vwhereAnterior
    
    
    
        'sqlSaldo = "SELECT Asientos.Codigo, Sum(Asientos.Debe) AS SumaDeDebe, Sum(Asientos.Haber) AS SumaDeHaber, Sum(Debe-haber) AS Saldo FROM Asientos WHERE (((Asientos.Fecha) < '" & strfechaMySQL(vFechaHasta) + "')) GROUP BY Asientos.Codigo HAVING (((Asientos.Codigo) = '" & vCodCuenta & "'))"
    Else
        
        
              
 vmensaje.Caption = "Calculando saldo de la cuenta: " + vCodCuenta
        
          sqlSaldo = "SELECT  sum(asientosdetalle.Debe) as d, sum(asientosdetalle.Haber) As H, (sum(asientosdetalle.Debe) - sum(asientosdetalle.Haber)) as Saldo " & _
                 " From  asientos " & _
                 " INNER JOIN asientosdetalle ON (asientos.Numero=asientosdetalle.Numero) " & _
                 " INNER JOIN cuentas ON (asientosdetalle.CodigoCuenta=cuentas.CodigoCuenta) " & _
                 " Where " & vwhereActual
        
        
        
        'sqlSaldo = "SELECT (Temp2.C10) AS Codigo, Sum(Temp2.C03) AS SumaDeDebe, Sum(Temp2.C04) AS SumaDeHaber, Sum(C03-C04) AS Saldo FROM Temp2 GROUP BY Temp2.C10 HAVING (((Temp2.C10) = '" & vCodCuenta & "'))"
    End If

    With rsSaldo
        Call .Open(sqlSaldo, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then
            CalcularSaldo = Val(Format(.Fields("Saldo").Value, "######0.00"))
        Else
           ' MsgBox "No se registraron movimientos para ser procesado.", vbCritical, "Error ..."
            CalcularSaldo = 0
        End If
    
    End With
        
    sqlSaldo = ""
    
    If rsSaldo.State = 1 Then
        rsSaldo.Close
        Set rsSaldo = Nothing
    End If
    
If Err Then GrabarLog "CalcularSaldo", Err.Number & " " & Err.Description, Me.Name
End Function
Public Sub dtpCuentas_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
    
    If KeyAscii = 13 Then
        Select Case Index
        
            Case 0
                
                If KeyAscii = 13 Then
                    dtpCuentas(1).SetFocus
                    dtpCuentas(1).Value = DiasDelMes(dtpCuentas(0).Value) & "/" & AjustarMes(Month(dtpCuentas(0).Value)) & "/" & Year(dtpCuentas(0).Value)
                End If
    
    
    

                
                
                dtpCuentas(1).SetFocus
                
                
            
            Case 1
                PbAcciones(0).SetFocus
        
        End Select
    
    End If
    
If Err Then GrabarLog "dtpCuentas_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub Form_Load()
On Error Resume Next

    With Me
        .Show
        .Left = 0
        .Top = 0
        '.Width = 7250
        '.Height = 5205
        .KeyPreview = True
        '.vnro_asto(0).SetFocus
    End With
    
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2 - 1000


If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub GrabarCuentaTemp(ByRef vCodCuenta As String, vcuenta As String, vnrobalance As Integer, vmarca As String)
On Error Resume Next
    'Dim vnrobalance As Integer
    Dim vfd, vfha As Date
    Dim vbc As String
    
    'vnrobalance = TraerDato("balances", " Activo='S' order by NroBalance Desc", "NroBalance", pathDBMySQL)
    vfd = TraerDato("balances", " Activo='S' and nrobalance=" + Str(vnrobalance) + " order by NroBalance Desc", "FechaInicio", pathDBMySQL)
    vfha = TraerDato("balances", " Activo='S' and nrobalance=" + Str(vnrobalance) + " order by NroBalance Desc", "FechaFin", pathDBMySQL)
    vbc = TraerDato("balances", " Activo='S' and nrobalance=" + Str(vnrobalance) + " order by NroBalance Desc", "codigo", pathDBMySQL)
    
    
    '--------`agregarlo panic sacar -------
    
    vfd = TraerDato("balances", " nrobalance=" + Str(vnrobalance) + " order by NroBalance Desc", "FechaInicio", pathDBMySQL)
    vfha = TraerDato("balances", " nrobalance=" + Str(vnrobalance) + " order by NroBalance Desc", "FechaFin", pathDBMySQL)
    vbc = TraerDato("balances", "  nrobalance=" + Str(vnrobalance) + " order by NroBalance Desc", "codigo", pathDBMySQL)
    
    '---------------
    
    
    vtimestampAjuste = "2012-08-08"
    
    Dim rsTempCuentas As New ADODB.Recordset, sqlTempCuentas As String
    
    
    sqlTempCuentas = "SELECT * FROM TempCuentas"
    vsaldoanterior = 0
    vsaldoactual = 0
    
    With rsTempCuentas
        Call .Open(sqlTempCuentas, ConnDDBB, adOpenDynamic, adLockPessimistic)
        
        .AddNew
        
        .Fields("Imprimir").Value = 1
        .Fields("Codigo").Value = vCodCuenta
        .Fields("Cuenta").Value = vcuenta
        
        If Not chkFechas.Value = 1 Then ' hay saldo anterior porque filtra por fecha
          '  vsaldoanterior = CalcularSaldo(vCodCuenta, True, dtpCuentas(0).Value)
             vsaldoanterior = CalSaldoAnteriorCtaContable(vmarca, vCodCuenta, vnrobalance, dtpCuentas(0).Value, vbc, vfd) ' paso 4
              
              
            If vsaldoanterior > 0 Then
                .Fields("SaldoAnteriorD").Value = vsaldoanterior
                .Fields("SaldoAnteriorH").Value = 0
            Else
                .Fields("SaldoAnteriorH").Value = vsaldoanterior * (-1)
                .Fields("SaldoAnteriorD").Value = 0
            End If
        
        Else
            
            .Fields("SaldoAnteriorD").Value = 0
            .Fields("SaldoAnteriorH").Value = 0
        
        End If
        
        
        ' calcula en saldo actual de la cuenta (renglón del mayor) ' panic: ojo, ver por que pasa False
        'vsaldoactual = CalcularSaldo(vCodCuenta, False, dtpCuentas(0).Value)
        'vsaldoactual = CalSaldoAnteriorCtaContable(vCodCuenta, vnroBalance, dtpCuentas(1).Value + 1)
        vsaldoactual = vsaldoanterior + vgneto
        
        
        
        If vsaldoactual > 0 Then
            .Fields("SaldoD").Value = vsaldoactual '+ Val(.Fields("SaldoAnteriorD").Value) - Val(.Fields("SaldoAnteriorH").Value)
        Else
           ' If ((Val(.Fields("SaldoAnteriorD").Value) - Val(.Fields("SaldoAnteriorH").Value)) > 0) And (vsaldoactual > 0) Then
           '     .Fields("SaldoH").Value = vsaldoactual + vsaldoanterior
           ' Else
           '     .Fields("SaldoH").Value = vsaldoactual * (-1) - vsaldoanterior
           ' End If
           .Fields("SaldoH").Value = -1 * vsaldoactual
        End If
        
        .Update
        
    End With
    
    sqlTempCuentas = ""
    
    If rsTempCuentas.State = 1 Then
        rsTempCuentas.Close
        Set rsTempCuentas = Nothing
    End If

If Err Then GrabarLog "GrabarCuentaTemp", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub

Public Sub fbalanceie()
Dim vul, sqlSaldoCaja As String
On Error Resume Next
    
    Unload Mantenimiento
    Load Mantenimiento
    
    With Mantenimiento.rsbalanceIE
        If .State = 1 Then .Close
        
      .Source = sqlBalanceIE(dtpCuentas(0).Value, dtpCuentas(1).Value, rbBalanceI.Value, rbBalanceE.Value)
      
        If Not .State = 1 Then .Open
        .Close
        .Open
        
    End With
    
    With drBalanceIE
    .Sections("TituloEmpresa").Controls("eperiodo").Caption = "Correspondiente al período desde: " + Format(dtpCuentas(0).Value, "dd-mm-yyyy") + " hasta: " + Format(dtpCuentas(1).Value, "dd-mm-yyyy")
    .Show
    End With

If Err Then Exit Sub

End Sub




Private Sub PbAcciones_Click(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 0, 1
   
            sqlC = ""
            sqlA = ""
            vmensaje.Caption = "Preparando espacio de datos ..."
            Call BorrarBase("Temp2", pathDBMySQL)
            Call BorrarBase("TempCuentasAcumuladas", pathDBMySQL)
            Call BorrarBase("TempCuentasMovimientos", pathDBMySQL)
            Call BorrarBase("TempCuentas", pathDBMySQL)
            
            vmensaje.Caption = "Construyendo información  ..."
            log.Clear
            
            GenerarDatos ' llena tempCuenta y temp2   ' paso1
            
            vmensaje.Caption = "Mostrando listado  ..."
            
            
            MostrarReporte (Index)

            
            
        Case 2
            Unload Me
    End Select
    
If Err Then GrabarLog "", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub pbContabilidad_Click(Index As Integer)
If Index = 0 Then
    Call fbuscarGrilla("(select * from cuentas where Imputable ='S') as t", "Cuenta", "CodigoCuenta", Me.vnro_asto1.Name, Me)    ' ema:
Else
    Call fbuscarGrilla("(select * from cuentas where Imputable ='S') as t", "Cuenta", "CodigoCuenta", Me.vnro_asto3.Name, Me)    ' ema:
End If

End Sub

Private Sub PusImprimir_Click()
    Call fbalanceie
End Sub

Private Sub vmensaje_Change()
log.AddItem (vmensaje.Caption)
End Sub

Private Sub vnro_asto_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
    
    If KeyAscii = 13 Then
        
        Select Case Index
    
            Case 0
'                If Not vnro_asto(Index).Text = "" Then
'                    vnro_asto(Index + 1).Text = TraerDato("Cuentas", "(CodigoCuenta = " & Trim(vnro_asto(Index).Text) & ")", "Cuenta")
'                End If
'
'                vnro_asto(Index + 2).SetFocus
'            Case 1
'                vnro_asto(Index + 1).SetFocus
            Case 2
'                If Not vnro_asto(Index).Text = "" Then
'                    vnro_asto(Index + 1).Text = TraerDato("Cuentas", "(CodigoCuenta = " & Trim(vnro_asto(Index).Text) & ")", "Cuenta")
'                End If
'
                dtpCuentas(0).SetFocus
            Case 3
                dtpCuentas(0).SetFocus
        
        End Select
    
    End If

If Err Then GrabarLog "txtBusqueda_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub vnro_asto1_Change()
    vnro_asto0.Text = vnro_asto1.Tag
    
   ' If vnro_asto2.Text = "" Then
    
        vnro_asto2.Text = vnro_asto1.Tag
    
        vnro_asto3.Tag = vnro_asto1.Tag
        
        vnro_asto3.Text = vnro_asto1.Text
        
        vnro_asto3.Tag = vnro_asto1.Tag
    
 '   End If
    
End Sub

Private Sub vnro_asto3_Change()
 vnro_asto2.Text = vnro_asto3.Tag
End Sub

Private Sub vnumero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.vleyenda.Text = TraerDato("asientos", "idasientos=" + Trim(vnumero), "leyenda", pathDBMySQL)
End If
End Sub
