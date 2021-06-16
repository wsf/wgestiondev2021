VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "Copia de KlexGrid.ocx"
Begin VB.Form frmAlarmas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alertas !!!"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   11190
   Begin XtremeSuiteControls.TabControl TabAlarmas 
      Height          =   6555
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   11085
      _Version        =   851968
      _ExtentX        =   19553
      _ExtentY        =   11562
      _StockProps     =   68
      Appearance      =   8
      PaintManager.BoldSelected=   -1  'True
      ItemCount       =   7
      Item(0).Caption =   "Información Gral"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "gridInfo"
      Item(0).Control(1)=   "GroupBox1"
      Item(1).Caption =   "Articulos"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "KlexArticulos"
      Item(2).Caption =   "Clientes"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "KlexClientes"
      Item(3).Caption =   "Proveedores"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "KlexProveedores"
      Item(4).Caption =   "Facturación Automática"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "KlexFacturaAutomatica(0)"
      Item(5).Caption =   "Errores de Saldos vs Facturas"
      Item(5).ControlCount=   1
      Item(5).Control(0)=   "gridcf"
      Item(6).Caption =   "Errores"
      Item(6).ControlCount=   1
      Item(6).Control(0)=   "log"
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   690
         Left            =   360
         TabIndex        =   8
         Top             =   4050
         Width           =   10050
         _Version        =   851968
         _ExtentX        =   17727
         _ExtentY        =   1217
         _StockProps     =   79
         Appearance      =   2
         Begin XtremeSuiteControls.PushButton PushButton1 
            Height          =   420
            Left            =   90
            TabIndex        =   9
            Top             =   180
            Width           =   2355
            _Version        =   851968
            _ExtentX        =   4154
            _ExtentY        =   741
            _StockProps     =   79
            Caption         =   "Imprimir Carta Comercio"
            Appearance      =   2
         End
         Begin XtremeSuiteControls.PushButton PushButton2 
            Height          =   420
            Left            =   3915
            TabIndex        =   10
            Top             =   180
            Width           =   2355
            _Version        =   851968
            _ExtentX        =   4154
            _ExtentY        =   741
            _StockProps     =   79
            Caption         =   "Imprimir Carta Comercio"
            Appearance      =   2
         End
         Begin XtremeSuiteControls.PushButton PushButton3 
            Height          =   420
            Left            =   7605
            TabIndex        =   11
            Top             =   180
            Width           =   2355
            _Version        =   851968
            _ExtentX        =   4154
            _ExtentY        =   741
            _StockProps     =   79
            Caption         =   "Imprimir Carta Comercio"
            Appearance      =   2
         End
      End
      Begin XtremeSuiteControls.ListBox log 
         Height          =   4245
         Left            =   -69865
         TabIndex        =   7
         Top             =   1980
         Visible         =   0   'False
         Width           =   10785
         _Version        =   851968
         _ExtentX        =   19024
         _ExtentY        =   7488
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridInfo 
         Height          =   2610
         Left            =   450
         TabIndex        =   6
         Top             =   720
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   4604
         _Version        =   393216
         Rows            =   0
         FixedRows       =   0
         FixedCols       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridcf 
         Height          =   3735
         Left            =   -69880
         TabIndex        =   5
         Top             =   1710
         Visible         =   0   'False
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   6588
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin Grid.KlexGrid KlexFacturaAutomatica 
         Height          =   3690
         Index           =   0
         Left            =   -69910
         TabIndex        =   2
         Top             =   1440
         Visible         =   0   'False
         Width           =   10875
         _ExtentX        =   19182
         _ExtentY        =   6509
         EnterKeyBehaviour=   0
         BackColorAlternate=   0
         GridLinesFixed  =   2
         BackColorFixed  =   -2147483626
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
         MouseIcon       =   "frmAlarmas.frx":0000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid KlexArticulos 
         Height          =   4125
         Left            =   -69880
         TabIndex        =   1
         Top             =   660
         Visible         =   0   'False
         Width           =   10875
         _ExtentX        =   19182
         _ExtentY        =   7276
         _Version        =   393216
         BackColor       =   16777215
         FixedRows       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   4210752
         BackColorSel    =   255
         ScrollBars      =   2
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid KlexClientes 
         Height          =   3855
         Left            =   -69940
         TabIndex        =   3
         Top             =   960
         Visible         =   0   'False
         Width           =   10905
         _ExtentX        =   19235
         _ExtentY        =   6800
         _Version        =   393216
         BackColor       =   16777215
         FixedRows       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   4210752
         BackColorSel    =   255
         ScrollBars      =   2
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid KlexProveedores 
         Height          =   4695
         Left            =   -69940
         TabIndex        =   4
         Top             =   1170
         Visible         =   0   'False
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   8281
         _Version        =   393216
         BackColor       =   16777215
         FixedRows       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   4210752
         BackColorSel    =   255
         ScrollBars      =   2
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "frmAlarmas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim rsAlarma As ADODB.Recordset
Dim vCantidadDeArticulos As Integer
Private Sub Form_Load()
    On Error Resume Next
    
    init
    
    With Me
        .Show
        .Top = (frmPrincipal.Height - .Height) / 2
        .Left = (frmPrincipal.Width - .Width) / 2
    End With
    
   Call CargarCCFactura
    
    
        
    If ControlarClientes = True Then
        MsgBox "Existen clientes SIN CODIGO!!!!!", vbCritical, "Mensaje ..."
    End If

    
    If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub init()


With Me.gridInfo

    .ColWidth(0) = 4000
    
    .ColWidth(1) = 2000
    
End With



With Me.gridcf

    .ColWidth(0) = 0
    
    .ColWidth(1) = 1000
    
    
    .ColWidth(2) = 3500
    
    .ColWidth(3) = 1500
    
    .ColWidth(4) = 1500
    
    .ColWidth(5) = 1500
    
End With

End Sub

Private Sub CargarCCFactura()
On Error Resume Next
Dim vline, vsql As String
Dim rs As New ADODB.Recordset


vsql = " select a.codigo, b.nombre, format(b.Saldo,'###,###,##0.00') as Saldo, format(a.Total,'###,###,##0.00') as Total , format(b.Saldo-a.Total,'###,###,##0.00') as Dif from " + _
" ( select codigo, nombre, sum(debito) - sum(credito) as Saldo From pcuentascorrientes group by codigo " + _
" ) b Inner Join " + _
"( select " + _
" codigo, sum(total) As total from pfactura t where t.estadodocumento is null or not  t.estadodocumento = 'Pagado' " + _
" group by codigo order by t.Codigo  desc ) a " + _
" on a.codigo = b.codigo " + _
" and not a.total = b.Saldo "

With rs
    .CursorLocation = adUseClient
    Call .Open(vsql, ConnDDBB, adOpenStatic, adLockPessimistic)
End With
   

Set gridcf.DataSource = rs

If rs.RecordCount > 0 Then

    vline = "Errores saldos proveedores: " + vbTab + Str(rs.RecordCount)
    Me.gridInfo.AddItem vline
End If





If Err Then Exit Sub
End Sub


Private Function BuscarFacturasAGenerar() As Integer
On Error Resume Next

    Dim rsFacturaAutomatica As New ADODB.Recordset, sqlFacturaAutomatica As String, vnroremito As Long, vCantidadFacturas As Integer
    
    vCantidadFacturas = 0
    sqlFacturaAutomatica = "SELECT * FROM FacturaAutomatica WHERE (FechaProximaEjecucion <= '" & strfechaMySQL(Date) & "') AND (FechaFin IS NULL);"
    
    With rsFacturaAutomatica
        .CursorLocation = adUseClient
        
        Call .Open(sqlFacturaAutomatica, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        Do Until .EOF = True
            vnroremito = 1 ' GenerarFactura
            
            'Traigo el Numero de Remito de la factura y guardar en FacturaAutomaticaEjecuciones
        
            .Fields("FechaProximaEjecucion").Value = ControlarEjecuciones(.Fields("FechaProximaEjecucion").Value, .Fields("idIntervalos").Value)
            .Fields("FechaUltimaEjecucion").Value = Date
            
            
            Call EjecutarScript("INSERT INTO FacturaAutomaticaEjecuciones (idFacturaAutomatica, Remito, Fecha) VALUES (" & .Fields("idFacturaAutomatica").Value & ", " & vnroremito & ",date)")
            
            
            .MoveNext
            vCantidadFacturas = vCantidadFacturas + 1
        Loop

    End With
    
    BuscarFacturasAGenerar = vCantidadFacturas
    
    sqlFacturaAutomatica = ""
    
    If rsFacturaAutomatica.State = 1 Then
        rsFacturaAutomatica.Close
        Set rsFacturaAutomatica = Nothing
    End If
    
If Err Then GrabarLog "BuscarFacturasAGenerar", Err.Number & " " & Err.Description, Me.Caption
End Function

Private Sub gridInfo_Click()
Dim vrow As Integer

vrow = Me.gridInfo.Row

Select Case vrow

    Case 0
        Me.TabAlarmas.SelectedItem = 5
        
 
End Select


End Sub

Private Sub KlexArticulos_DblClick()
On Error Resume Next

    With KlexArticulos
        If Not .TextMatrix(.Row, 1) = "" Then
            With frmArticulosAlta
                .Show
                .ModificarArticulo (KlexArticulos.TextMatrix(KlexArticulos.Row, 1))
                .vaccion = "Modificar"
            End With
        End If
    End With

If Err Then GrabarLog "KlexArticulos_DblClick", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Function CargarArticulosConErrores() As Integer
On Error Resume Next

    Set rsAlarma = New ADODB.Recordset
    Dim sqlAlarma As String

    sqlAlarma = "SELECT idArticulos, Codigo, Descrip, PVenta1, FechaAlta FROM Articulos WHERE (idSubRubros = '') OR (idSubRubros IS null) OR (idRubros = '') OR (idRubros IS null);"

    With rsAlarma
        If .State = 1 Then .Close
        
        .CursorLocation = adUseClient
        
        Call .Open(sqlAlarma, ConnDDBB, adOpenStatic, adLockReadOnly)
    
        CargarArticulosConErrores = .RecordCount
        
        If .EOF = True Then
            Set KlexArticulos.DataSource = Nothing
        Else
            Call ConfigurarGrilla
        End If
    End With
    
    sqlAlarma = ""

    If rsAlarma.State = 1 Then
        rsAlarma.Close
        Set rsAlarma = Nothing
    End If
    
    MousePointer = vbDefault
    
    
If Err Then GrabarLog "CargarArticulosConErrores", Err.Number & " " & Err.Description, Me.Caption
End Function
Private Sub ConfigurarGrilla()
On Error Resume Next

    Dim i As Integer
   
    With KlexArticulos
        .Cols = 7
        
        .FixedRows = 1
        .FixedCols = 1
            
        .ColWidth(0) = 250
        .ColWidth(1) = 0
        .ColWidth(2) = 1250
        .ColWidth(3) = 5000
        .ColWidth(4) = 1250
        .ColWidth(5) = 1250
        
        .Row = KlexArticulos.Rows - 1
    
        Set .DataSource = rsAlarma
        
        .TextMatrix(0, 1) = ""
        .TextMatrix(0, 2) = "Código"
        .TextMatrix(0, 3) = "Código"
        .TextMatrix(0, 4) = "P. Venta 1"
        .TextMatrix(0, 5) = "F. Alta"

    End With
    
If Err Then GrabarLog "ConfigurarGrilla", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub lblAlarma_Click(Index As Integer)

End Sub

