VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmSaldosClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Saldos de Clientes"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11655
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   11655
   Begin VB.Frame Frame3 
      Height          =   315
      Left            =   300
      TabIndex        =   19
      Top             =   2790
      Width           =   4725
      Begin MSComctlLib.ProgressBar barra 
         Height          =   165
         Left            =   30
         TabIndex        =   20
         Top             =   120
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   291
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.CommandButton cmdUltimo 
      Caption         =   "Ult. Listado"
      Height          =   525
      Left            =   1320
      TabIndex        =   21
      Top             =   2280
      Width           =   975
   End
   Begin VB.Frame fecha 
      Caption         =   "Fecha :"
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   2790
      TabIndex        =   14
      Top             =   1410
      Visible         =   0   'False
      Width           =   2205
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   285
         Left            =   720
         TabIndex        =   17
         Top             =   180
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   74842113
         CurrentDate     =   38238
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   285
         Left            =   720
         TabIndex        =   18
         Top             =   510
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   74842113
         CurrentDate     =   38238
      End
      Begin VB.Label Label5 
         Caption         =   "Desde :"
         Height          =   195
         Left            =   60
         TabIndex        =   16
         Top             =   210
         Width           =   705
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta :"
         Height          =   225
         Left            =   60
         TabIndex        =   15
         Top             =   540
         Width           =   675
      End
   End
   Begin VB.CheckBox chkDetalle 
      Caption         =   "Con detalles"
      Enabled         =   0   'False
      Height          =   225
      Left            =   3390
      TabIndex        =   13
      Top             =   2460
      Width           =   1305
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   330
      TabIndex        =   8
      Top             =   1170
      Width           =   2385
      Begin VB.OptionButton Option1 
         Caption         =   "Deudores"
         Height          =   225
         Left            =   30
         TabIndex        =   12
         Top             =   120
         Value           =   -1  'True
         Width           =   2025
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Saldados"
         Height          =   225
         Left            =   30
         TabIndex        =   11
         Top             =   360
         Width           =   2025
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Saldos a favor de Clientes"
         Height          =   195
         Left            =   30
         TabIndex        =   10
         Top             =   600
         Width           =   2295
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Todos los Clientes"
         Height          =   195
         Left            =   30
         TabIndex        =   9
         Top             =   810
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc bcliente 
      Height          =   330
      Left            =   240
      Top             =   3240
      Visible         =   0   'False
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "bcliente"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   525
      Left            =   330
      Picture         =   "lsaldosclientes2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Generar reporte para imprimir"
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Clientes :"
      ForeColor       =   &H00000080&
      Height          =   1185
      Left            =   330
      TabIndex        =   0
      Top             =   0
      Width           =   4665
      Begin VB.TextBox vlocalidad 
         Height          =   315
         Left            =   1440
         TabIndex        =   6
         Top             =   840
         Width           =   2865
      End
      Begin VB.TextBox vchasta 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   510
         Width           =   2865
      End
      Begin VB.TextBox vcdesde 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   180
         Width           =   2865
      End
      Begin VB.Label Label3 
         Caption         =   "> Localidad :"
         Height          =   225
         Left            =   420
         TabIndex        =   7
         Top             =   870
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "> Hasta :"
         Height          =   225
         Left            =   450
         TabIndex        =   2
         Top             =   540
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "> Desde :"
         Height          =   195
         Left            =   450
         TabIndex        =   1
         Top             =   240
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmSaldosClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vcodigodesde, vcodigohasta As String
Dim vufactura As Double
Dim vFiltro As String
Function BuscarSaldo(vcodigo As String) As Double
On Error Resume Next

    Dim connSaldos As New ADODB.Connection
    Dim rsSaldos As New ADODB.Recordset
    Dim sqlSaldos As String
    
    With connSaldos
        .ConnectionString = pathDBMySQL
        .Open
    End With
    
    sqlSaldos = "SELECT * FROM saldo_clientes WHERE (codigo = '" & vcodigo & "')"
        
    With rsSaldos
        .Open sqlSaldos, connSaldos, adOpenDynamic, adLockReadOnly
        
        If Not .EOF = True Then BuscarSaldo = Val(Format(.Fields("saldo").Value, "#######0.00"))
    
    End With
        
    sqlSaldos = ""
    
    rsSaldos.Close
    Set rsSaldos = Nothing
    
    connSaldos.Close
    Set connSaldos = Nothing
    
If Err Then GrabarLog "BuscarSaldo", Err.Number & " " & Err.Description, Me.Caption
End Function
Function CalcularSaldo(vcodigo As String) As Double
On Error Resume Next

    Dim connSaldoCliente As New ADODB.Connection
    Dim rsSaldoCliente As New ADODB.Recordset
    Dim sqlSaldoCliente As String
    
    With connSaldoCliente
        .ConnectionString = pathDBMySQL
        .Open
    End With
    
    sqlSaldoCliente = "SELECT cuentascorrientes.Codigo, Sum(cuentascorrientes.Debito) AS SumaDeDebito, Sum(cuentascorrientes.Credito) AS SumaDeCredito, [SumadeDebito]-[SumaDeCredito] AS Saldo FROM Clientes INNER JOIN cuentascorrientes ON Clientes.Codigo = cuentascorrientes.Codigo GROUP BY cuentascorrientes.Codigo HAVING (((cuentascorrientes.Codigo)= '" + vcodigo + "'))"

    With rsSaldoCliente
        .Open sqlSaldoCliente, connSaldoCliente, adOpenStatic, adLockReadOnly
        
        If Not .EOF = True Then
            CalcularSaldo = .Fields("Saldo").Value
        Else
            CalcularSaldo = 0
        End If
    End With
    
    sqlSaldoCliente = ""
    
    rsSaldoCliente.Close
    Set rsSaldoCliente = Nothing
    
    connSaldoCliente.Close
    Set connSaldoCliente = Nothing
    
If Err Then
    MsgBox "Ocurrió un problema al intentar calcular el saldo del cliente: " + vcodigo + Chr(13) + " Revise los movimiento de este cliente.", vbCritical
    GrabarLog "CalcularSaldo", Err.Number & " " & Err.Description, Me.Caption
End If
End Function
Function UltimaFactura(vcodigo As String) As Double
On Error Resume Next

    Dim connUFactura As New ADODB.Connection
    Dim rsUFactura As New ADODB.Recordset
    Dim sqlUFactura As String
    
    With connUFactura
        .ConnectionString = pathDBMySQL
        .Open
    End With
    
    sqlUFactura = "SELECT fecha, codigo, total FROM factura WHERE codigo =  '" + Trim(vcodigo) + "' ORDER BY fecha DESC"

    With rsUFactura
        .Open sqlUFactura, connUFactura, adOpenStatic, adLockReadOnly
        
        If Not .EOF = True Then
            UltimaFactura = .Fields("total").Value
        Else
            UltimaFactura = 0
        End If
    
    End With

    sqlUFactura = ""
    
    rsUFactura.Close
    Set rsUFactura = Nothing
    
    connUFactura.Close
    Set connUFactura = Nothing
    
If Err Then GrabarLog "UltimaFactura", Err.Number & " " & Err.Description, Me.Caption
End Function
Private Sub BuscarCliente(vcliente As String, dh As String)
On Error Resume Next

    Dim connClientes As New ADODB.Connection
    Dim rsClientes As New ADODB.Recordset
    Dim sqlClientes As String
    
    With connClientes
        .ConnectionString = pathDBMySQL
        .Open
    End With
    
    sqlClientes = "SELECT * FROM clientes WHERE (nombre = '" + Trim(vcliente) + "') OR (codigo = '" + Trim(vcliente) + "')"
    
    With rsClientes
        .Open sqlClientes, connClientes, adOpenStatic, adLockReadOnly
        
        If .EOF = True Then
    
            'frmBuscaCliente.Show
            If dh = "d" Then
                'frmBuscaCliente.o = 8
            Else
                'frmBuscaCliente.o = 9
            End If
    
            'frmBuscaCliente.Show
            'frmBuscaCliente.txtBusca = vcliente
            'frmBuscaCliente.txtBusca.SetFocus
    
        Else
        
            If dh = "d" Then
                vcdesde.Text = .Fields("nombre").Value
                vcodigodesde = .Fields("Codigo").Value
                vchasta.SetFocus
            Else
                vchasta.Text = .Fields("nombre").Value
                vcodigohasta = .Fields("Codigo").Value
                vlocalidad.SetFocus
            End If
        
        End If
    
    End With

    sqlClientes = ""
    
    rsClientes.Close
    Set rsClientes = Nothing
    
    connClientes.Close
    Set connClientes = Nothing
    
If Err Then GrabarLog "BuscarCliente", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub idetallectacte(fdesde As Date, fhasta As Date)
Dim condetalle As Integer

If Not (vcodigodesde = "" Or vcodigohasta = "") Then
    bcliente.RecordSource = "select * from clientes where codigo >= '" + vcodigodesde + "' and codigo <= '" + vcodigohasta + "'"
End If

    
    bcliente.Refresh

    condetalle = 0

    If MsgBox("¿Imprimir movimientos con detalle ?", vbYesNo, " Cuentas Corrientes...") = vbYes Then condetalle = 1

Do Until bcliente.Recordset.EOF
    Mantenimiento.rsccc.Filter = "codigo = '" + bcliente.Recordset("codigo") + "' and importe = 0 and  fecha <= #" + strFecha(fhasta) + "# and fecha >= #" + strFecha(fdesde) + "#"
    Mantenimiento.rsccc.Sort = "fecha ASC, id ASC"
    
    Mantenimiento.rscctacte_detalle.Filter = "codigo = '" + bcliente.Recordset("codigo") + "' and importe = 0 and  fecha <= #" + strFecha(fhasta) + "# and fecha >= #" + strFecha(fdesde) + "#"
    Mantenimiento.rscctacte_detalle.Sort = "fecha ASC, id ASC"
    
    If condetalle = 1 Then
        condetalle = 1
       ' drcuentascorrientes_detalles.Sections(2).Controls("gnombre").Caption = gnombre
       ' drcuentascorrientes_detalles.Sections(2).Controls("gdireccion").Caption = gdireccion
       ' drcuentascorrientes_detalles.Sections(2).Controls("gtelefono").Caption = gtelefono
        
        drcuentascorrientes_detalles.Visible = False
        drcuentascorrientes_detalles.PrintReport False
    Else
        
        'drcuentascorrientes.Sections(2).Controls("gnombre").Caption = gnombre
        'drcuentascorrientes.Sections(2).Controls("gdireccion").Caption = gdireccion
        'drcuentascorrientes.Sections(2).Controls("gtelefono").Caption = gtelefono
        
        drcuentascorrientes.Sections("section2").Controls("vcliente").Caption = bcliente.Recordset("codigo") + " " + bcliente.Recordset("nombre")
        'drcuentascorrientes.Sections("section3").Controls("vsaldo").Caption = saldo.Caption
        
        drcuentascorrientes.Visible = False
        drcuentascorrientes.PrintReport False
    End If
    
    
    Mantenimiento.rsccc.Close
    
    bcliente.Recordset.MoveNext

Loop

End Sub
Private Sub Limpiar()
On Error Resume Next

    vcdesde.Text = ""
    vchasta.Text = ""
    vlocalidad.Text = ""
    vcodigodesde = ""
    vcodigohasta = ""

If Err Then GrabarLog "Limpiar", Err.Number & " " & Err.Description, Me.Caption
End Sub
Public Sub vcdesde_KeyPress(Keyascii As Integer)

    If Keyascii = 13 Then BuscarCliente vcdesde, "d"

End Sub
Public Sub vchasta_KeyPress(Keyascii As Integer)

    If Keyascii = 13 Then BuscarCliente vchasta, "h"

End Sub
Public Sub vlocalidad_KeyPress(Keyascii As Integer)

    If Keyascii = 13 Then cmdImprimir.SetFocus

End Sub
Private Sub cmdUltimo_Click()
    'CambiarImpresora 1
    Unload Mantenimiento
    Load Mantenimiento
    drcliente.Show
End Sub
Private Sub cmdImprimir_Click()
On Error Resume Next

    'CambiarImpresora 1

    If chkDetalle.Value = 1 Then
        idetallectacte fdesde, fhasta
        Exit Sub
    End If

    MousePointer = vbHourglass

    Call Filtrar
    Call Imprimir
    Call Limpiar

    MousePointer = vbDefault

If Err Then GrabarLog "", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Filtrar()
On Error Resume Next

Dim vvalores, vsql As String

    vFiltro = ""
    
    If Not (vcodigodesde = "" And vcodigohasta = "") Then
        vFiltro = vFiltro + " and (codigo_num >= " & vcodigodesde & " and codigo_num <= " & vcodigohasta & ")"
    End If

    If Not vlocalidad = "" Then
        vFiltro = vFiltro + " and Localidad like '%" + Trim(vlocalidad) + "%'"
    End If
    
    With bcliente
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM Clientes WHERE 1=1" + vFiltro + " order by codigo"
        .Refresh
            
        If Not .Recordset.EOF = True Then
            .Recordset.MoveFirst
            barra.Max = .Recordset.RecordCount
            barra.Value = 0
        End If
    
        Do Until .Recordset.EOF = True
             
             
             
             
             vufactura = Val(Format(UltimaFactura(.Recordset("codigo").Value), "########0.00"))
             
             
             vvalores = "ufactura =" + Str(vufactura) + _
             ", saldo= " + Str(Val(Format(BuscarSaldo(.Recordset("codigo").Value) - vufactura, "########0.00"))) + _
             ", saldoTotal=" + Str(Val(Format(.Recordset("Saldo").Value, "#######0.00")) + vufactura) + _
             ", saldof=" + Str(Val(Format(.Recordset("Saldo").Value, "#######0.00")) + vufactura)
             
             
             vsql = "update clientes set " + vvalores + " where codigo = '" + .Recordset("codigo").Value + "'"
             
             Call EjecutarScript(vsql, pathDBMySQL)
             
             
             
            '.Recordset("ufactura").Value = vufactura
            '.Recordset("Saldo").Value = Val(Format(BuscarSaldo(.Recordset("codigo").Value) - vufactura, "########0.00"))
            '.Recordset("SaldoTotal").Value = Val(Format(.Recordset("Saldo").Value, "#######0.00")) + vufactura
            '.Recordset("Saldof").Value = Val(Format(.Recordset("Saldo").Value, "#######0.00")) + vufactura
            
            
            
            .Recordset.MoveNext
            barra.Value = barra.Value + 1
            
        Loop
    
    End With

If Err Then GrabarLog "Filtrar", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Imprimir()
On Error Resume Next

    Unload Mantenimiento
    Load Mantenimiento

    MsgBox "Prepare la impresora !", vbInformation, "Mensaje ..."

    With Mantenimiento.rsclcli
            
        If Not .State = 1 Then .Open
        .Close
        .Open

        If Option1.Value = True Then .Filter = "Saldof > 0" + vFiltro
        If Option2.Value = True Then .Filter = "saldoTotal = 0" + vFiltro
        If Option3.Value = True Then .Filter = "saldoTotal < 0" + vFiltro
        If Option4.Value = True Then .Filter = "id > 0 " + vFiltro

        .Sort = "nombre ASC"
    
    End With
    With drcliente
        '.Sections(2).Controls("gnombre").Caption = gnombre
        '.Sections(2).Controls("gdireccion").Caption = gdireccion
        '.Sections(2).Controls("gtelefono").Caption = gtelefono

        .Show
    End With

If Err Then GrabarLog "Imprimir", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub chkDetalle_Click()
On Error Resume Next
    
    fecha.Visible = CBool(chkDetalle.Value)

If Err Then GrabarLog "chkDetalle_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Load()
On Error Resume Next

    With Me
        .Top = 1000
        .Left = 1300
        .width = 5280
        .height = 3675
    End With
    
    fdesde.Value = Date
    fhasta.Value = Date

If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Unload(Cancel As Integer)
   ' CambiarImpresora 0
End Sub
