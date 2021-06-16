VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMonitor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monitor del Sistema"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   5670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   5670
   Begin VB.CommandButton Command1 
      Caption         =   "Cheques Todos"
      Height          =   495
      Left            =   3750
      TabIndex        =   10
      Top             =   2790
      Width           =   1875
   End
   Begin VB.CommandButton cmdCheques 
      Caption         =   "Cheques tercero"
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   2790
      Width           =   2025
   End
   Begin VB.CommandButton cmdLog 
      Caption         =   "Prueba"
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   3300
      Width           =   5655
   End
   Begin VB.CommandButton cmdClienteCtaCte 
      Caption         =   "Listado Movimientos de Cte. Cta. por Cliente"
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   1260
      Width           =   5655
   End
   Begin VB.Frame fraFechas 
      Caption         =   "Rango de Fecha"
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5655
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Index           =   0
         Left            =   540
         TabIndex        =   5
         Top             =   300
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58982401
         CurrentDate     =   39713
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Index           =   1
         Left            =   3750
         TabIndex        =   6
         Top             =   300
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58982401
         CurrentDate     =   39713
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   465
      Left            =   0
      TabIndex        =   3
      Top             =   3780
      Width           =   5625
   End
   Begin VB.CommandButton cmdMovimientosCtaCte 
      Caption         =   "Listado Movimientos de Cte. Cta."
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   1770
      Width           =   5655
   End
   Begin VB.CommandButton cmdFacturasDetalle 
      Caption         =   "Listado de Documentos C/Detalle"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   750
      Width           =   5655
   End
   Begin VB.CommandButton cmdPagos 
      Caption         =   "Listado de Pagos"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   2280
      Width           =   5655
   End
End
Attribute VB_Name = "frmMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCheques_Click()
On Error Resume Next

    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox "Prepare la Impresora !!!", vbInformation, "Mensaje ..."
    
    With Mantenimiento.rsChequesAgrupados_Grouping
        If .State = 1 Then .Close
        .Source = "SHAPE {SELECT * FROM cheques WHERE (not propietario = 'propio') and (((Fecha) >= '" & strfechaMySQL(dtpFecha(0).Value) & "' And (Fecha)<= '" & strfechaMySQL(dtpFecha(1).Value) & "'))}  AS ChequesAgrupados COMPUTE ChequesAgrupados BY 'Fecha'"
        If .State = 0 Then .Open
        .Close
        .Open
    End With
    
    With drChequesAgrupados
        .Sections("PageHeader").Controls("lbltitulo").Caption = "Listado de Cheques Dia por Dia"
        .Sections("PageHeader").Controls("lblFecha").Caption = "Fecha Desde : " & dtpFecha(0).Value & " - Fecha Hasta : " & dtpFecha(1).Value

        .Show
    End With

If Err Then GrabarLog "cmdCheques_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cmdClienteCtaCte_Click()
On Error Resume Next

    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox "Prepare la Impresora !!!", vbInformation, "Mensaje ..."
    
    With Mantenimiento.rsClientesCtaCte
        If .State = 1 Then .Close
        
        .Source = "SHAPE {SELECT Clientes.Codigo, Last(Clientes.Nombre) AS Nombre, Last(Clientes.Direccion) AS Direccion, Last(Clientes.Localidad) AS Localidad, Last(Clientes.Iva) AS Iva, Last(Clientes.Cuit) AS Cuit, Clientes.id FROM cuentascorrientes INNER JOIN Clientes ON cuentascorrientes.Codigo = Clientes.Codigo GROUP BY Clientes.Codigo, Clientes.id, cuentascorrientes.Fecha HAVING (((cuentascorrientes.Fecha) >= '" & strfechaMySQL(dtpFecha(0).Value) & "' And (cuentascorrientes.Fecha)<= '" & strfechaMySQL(dtpFecha(1).Value) & "'))}  AS ClientesCtaCte APPEND ({SELECT * FROM cuentascorrientes WHERE (((cuentascorrientes.Fecha) >= '" & strfechaMySQL(dtpFecha(0).Value) & "' And (cuentascorrientes.Fecha)<= '" & strfechaMySQL(dtpFecha(1).Value) & "'))}  AS CtaCte RELATE 'clientes.Codigo' TO 'cuentascorrientes.Codigo') AS CtaCte"
        
        If .State = 0 Then .Open
        .Close
        .Open
       
    End With
    
    With drMonitorClienteCtaCte
        .Sections("PageHeader").Controls("lbltitulo").Caption = "Listado de Movimientos de Cuenta Corriente Por Cliente"
        .Sections("PageHeader").Controls("lblFecha").Caption = "Fecha Desde : " & dtpFecha(0).Value & " - Fecha Hasta : " & dtpFecha(1).Value

        .Show
    End With

If Err Then GrabarLog "cmdClienteCtaCte_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cmdFacturasDetalle_Click()
On Error Resume Next

Close Mantenimiento
Load Mantenimiento

    With Mantenimiento.rsFacturaFDetalle
        If .State = 1 Then .Close
        .Source = " SHAPE {SELECT * FROM factura WHERE ((fecha >=  '" & strfechaMySQL(dtpFecha(0).Value) & "') AND (fecha <= '" & strfechaMySQL(dtpFecha(1).Value) & "')) }  AS FacturaFDetalle APPEND ({SELECT * FROM fdetalle}  AS Fdetalle RELATE 'Remito' TO 'Remito') AS Fdetalle"
       ' .Source = " SHAPE {SELECT * FROM factura AS FacturaFDetalle APPEND ({SELECT * FROM fdetalle}  AS Fdetalle RELATE 'Remito' TO 'Remito') AS Fdetalle"

    If .State = 0 Then .Open
        .Close
        .Open
    End With


    With drMonitorFacturaFDetalle
        .Sections("PageHeader").Controls("lbltitulo").Caption = "Listado de Movimientos Documentos"
        .Sections("PageHeader").Controls("lblFecha").Caption = "Fecha Desde : " & dtpFecha(0).Value & " - Fecha Hasta : " & dtpFecha(1).Value

        .Show
    End With
If Err Then GrabarLog "cmdFacturasDetalle_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cmdLog_Click()

Unload Mantenimiento
Load Mantenimiento

drMonitor.Show

'On Error Resume Next'
'
'    With Mantenimiento.rsLog
'        If .State = 1 Then .Close
'        .Source = "SELECT * FROM log WHERE ((fecha >=  #" & strFecha(dtpFecha(0).Value) & "#) AND (fecha <= #" & strFecha(dtpFecha(1).Value) & "#)) ORDER BY id ASC"
'
'        If .State = 0 Then .Open
'        .Close
'        .Open
'    End With
'
'    With drlog
'        .Sections("PageHeader").Controls("lbltitulo").Caption = "Listado de Errores del Sistema"
'        .Sections("PageHeader").Controls("lblFecha").Caption = "Fecha Desde : " & dtpFecha(0).Value & " - Fecha Hasta : " & dtpFecha(1).Value
'        .Show
'    End With

'If Err Then GrabarLog "cmdLog_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cmdMovimientosCtaCte_Click()
On Error Resume Next

    With Mantenimiento.rsCuentaCorriente
        If .State = 1 Then .Close
        .Source = "SELECT * FROM cuentascorrientes WHERE ((fecha >=  '" & strfechaMySQL(dtpFecha(0).Value) & "') AND (fecha <= '" & strfechaMySQL(dtpFecha(1).Value) & "')) ORDER BY id ASC"
    
        If .State = 0 Then .Open
        .Close
        .Open
    End With
        
    With drMonitorCtaCte
        .Sections("PageHeader").Controls("lbltitulo").Caption = "Listado de Movimientos de Cuenta Corriente"
        .Sections("PageHeader").Controls("lblFecha").Caption = "Fecha Desde : " & dtpFecha(0).Value & " - Fecha Hasta : " & dtpFecha(1).Value
        .Show
    End With


If Err Then GrabarLog "cmdMovimientosCtaCte_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cmdPagos_Click()
On Error Resume Next

    With Mantenimiento.rsCuentaCorriente
        If .State = 1 Then .Close
        .Source = "SELECT * FROM cuentascorrientes WHERE (credito > 0) AND ((fecha >=  '" & strfechaMySQL(dtpFecha(0).Value) & "') AND (fecha <= '" & strfechaMySQL(dtpFecha(1).Value) & "')) ORDER BY id ASC"
    
        If .State = 0 Then .Open
        .Close
        .Open
    End With
    
    With drMonitorCtaCte
        .Sections("PageHeader").Controls("lbltitulo").Caption = "Listado de Pagos de Cuenta Corriente"
        .Sections("PageHeader").Controls("lblFecha").Caption = "Fecha del Listado: " & Date & " / Fecha Desde : " & dtpFecha(0).Value & " - Fecha Hasta : " & dtpFecha(1).Value
        .Show
    End With

If Err Then GrabarLog "cmdPagos_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cmdSalir_Click()
On Error Resume Next

    Unload Me

If Err Then GrabarLog "cmdSalir_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Load()
On Error Resume Next

    Dim i As Integer
    
    For i = 0 To 1
        dtpFecha(i).Value = Date
    Next
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub
