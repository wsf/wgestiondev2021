VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{9746E3DA-06E1-4D26-9CE4-D9F6411A9C70}#1.0#0"; "SMGA_OcxTxt2009.ocx"
Begin VB.Form frmCajaMovimientos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos de Caja"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   240
   ClientWidth     =   10920
   Icon            =   "frmCajaMovimientos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   10920
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   135
      Left            =   60
      TabIndex        =   33
      Top             =   510
      Width           =   10815
      _Version        =   851968
      _ExtentX        =   19076
      _ExtentY        =   238
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   645
      Left            =   30
      TabIndex        =   28
      Top             =   -90
      Width           =   10875
      _Version        =   851968
      _ExtentX        =   19182
      _ExtentY        =   1138
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton PBAcciones 
         Height          =   435
         Index           =   1
         Left            =   9630
         TabIndex        =   29
         Top             =   150
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Cerrar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCajaMovimientos.frx":000C
      End
      Begin XtremeSuiteControls.PushButton PBAcciones 
         Height          =   390
         Index           =   0
         Left            =   3240
         TabIndex        =   30
         Top             =   150
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   688
         _StockProps     =   79
         Caption         =   "Imprimir"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCajaMovimientos.frx":040C
      End
      Begin XtremeSuiteControls.PushButton PBAcciones 
         Height          =   390
         Index           =   2
         Left            =   1980
         TabIndex        =   31
         Top             =   150
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   688
         _StockProps     =   79
         Caption         =   "Buscar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCajaMovimientos.frx":0827
      End
      Begin XtremeSuiteControls.PushButton PBAcciones 
         Height          =   390
         Index           =   3
         Left            =   90
         TabIndex        =   32
         Top             =   150
         Width           =   1875
         _Version        =   851968
         _ExtentX        =   3307
         _ExtentY        =   688
         _StockProps     =   79
         Caption         =   "Ver Transacciones"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCajaMovimientos.frx":0C5E
      End
   End
   Begin XtremeSuiteControls.GroupBox GBBusqueda 
      Height          =   1815
      Left            =   2220
      TabIndex        =   18
      Top             =   3330
      Visible         =   0   'False
      Width           =   7365
      _Version        =   851968
      _ExtentX        =   12991
      _ExtentY        =   3201
      _StockProps     =   79
      Caption         =   "Busqueda "
      BackColor       =   14737632
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.CheckBox chkOcultarNoBuscados 
         Height          =   255
         Left            =   1680
         TabIndex        =   24
         Top             =   840
         Width           =   5415
         _Version        =   851968
         _ExtentX        =   9551
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Solo mostrar criterios buscados"
         BackColor       =   14737632
         UseVisualStyle  =   -1  'True
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   0
         Picture         =   "frmCajaMovimientos.frx":11F8
         ScaleHeight     =   555
         ScaleWidth      =   7305
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1260
         Width           =   7305
         Begin VB.Label lblWGESTION2010 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "WGESTION 2010"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   240
            Index           =   0
            Left            =   50
            TabIndex        =   20
            Top             =   150
            Width           =   1770
         End
         Begin VB.Label lblWGESTION2010 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "WGESTION 2010"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   240
            Index           =   1
            Left            =   75
            TabIndex        =   21
            Top             =   170
            Width           =   1770
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtBusqueda 
         Height          =   315
         Left            =   1680
         TabIndex        =   22
         Top             =   360
         Width           =   5415
         _Version        =   851968
         _ExtentX        =   9551
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   23
         Top             =   400
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Texto a Buscar:"
         Transparent     =   -1  'True
      End
   End
   Begin MSAdodcLib.Adodc bCajaMovimientos 
      Height          =   330
      Left            =   120
      Top             =   7560
      Visible         =   0   'False
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
      Caption         =   "bCajaMovimientos"
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
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "Borrar"
      Height          =   495
      Left            =   4680
      Picture         =   "frmCajaMovimientos.frx":62AB
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7560
      UseMaskColor    =   -1  'True
      Width           =   765
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
      Height          =   495
      Index           =   1
      Left            =   3870
      Picture         =   "frmCajaMovimientos.frx":66B0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7560
      UseMaskColor    =   -1  'True
      Width           =   825
   End
   Begin VB.Frame Frame3 
      Height          =   30
      Left            =   570
      TabIndex        =   5
      Top             =   8400
      Width           =   9465
   End
   Begin XtremeSuiteControls.PushButton cmdVerDetalle 
      Height          =   375
      Left            =   8040
      TabIndex        =   12
      Top             =   7680
      Width           =   4335
      _Version        =   851968
      _ExtentX        =   7646
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Ver detalle del movimiento seleccionado"
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
      Picture         =   "frmCajaMovimientos.frx":6AB1
   End
   Begin XtremeSuiteControls.PushButton PBFiltrar 
      Height          =   405
      Left            =   60
      TabIndex        =   25
      Top             =   7080
      Width           =   10815
      _Version        =   851968
      _ExtentX        =   19076
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Filtrar Movimientos"
      UseVisualStyle  =   -1  'True
      Picture         =   "frmCajaMovimientos.frx":6FFB
   End
   Begin XtremeSuiteControls.TabControl TabBancos 
      Height          =   5655
      Left            =   60
      TabIndex        =   6
      Top             =   1410
      Width           =   10815
      _Version        =   851968
      _ExtentX        =   19076
      _ExtentY        =   9975
      _StockProps     =   68
      ItemCount       =   2
      Item(0).Caption =   "Movimientos de Caja"
      Item(0).ControlCount=   9
      Item(0).Control(0)=   "lblBanco(0)"
      Item(0).Control(1)=   "pbCarga(0)"
      Item(0).Control(2)=   "lblBanco(2)"
      Item(0).Control(3)=   "txtFecha(0)"
      Item(0).Control(4)=   "txtFecha(1)"
      Item(0).Control(5)=   "lblBanco(3)"
      Item(0).Control(6)=   "txtCaja(0)"
      Item(0).Control(7)=   "txtCaja(1)"
      Item(0).Control(8)=   "PusArreglarSaldo"
      Item(1).Caption =   "Ver Datos"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "KlexMovimientos"
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid KlexMovimientos 
         Height          =   5235
         Left            =   -69970
         TabIndex        =   27
         Top             =   390
         Visible         =   0   'False
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   9234
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin XtremeSuiteControls.PushButton PusArreglarSaldo 
         Height          =   255
         Left            =   9000
         TabIndex        =   26
         Top             =   0
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Arreglar Saldo"
         UseVisualStyle  =   -1  'True
      End
      Begin Aplisoft_CajasDeTexto.TxF txtFecha 
         Height          =   315
         Index           =   0
         Left            =   2160
         TabIndex        =   1
         Top             =   1050
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.FlatEdit txtCaja 
         Height          =   315
         Index           =   0
         Left            =   2190
         TabIndex        =   0
         Top             =   600
         Width           =   975
         _Version        =   851968
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtCaja 
         Height          =   315
         Index           =   1
         Left            =   3720
         TabIndex        =   8
         Top             =   600
         Width           =   6675
         _Version        =   851968
         _ExtentX        =   11774
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   0
         Left            =   3300
         TabIndex        =   9
         Tag             =   "Caja"
         Top             =   600
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin Aplisoft_CajasDeTexto.TxF txtFecha 
         Height          =   315
         Index           =   1
         Left            =   2160
         TabIndex        =   2
         Top             =   1470
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label lblBanco 
         Height          =   255
         Index           =   3
         Left            =   870
         TabIndex        =   11
         Top             =   1530
         Width           =   1005
         _Version        =   851968
         _ExtentX        =   1773
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha Hasta:"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblBanco 
         Height          =   315
         Index           =   2
         Left            =   870
         TabIndex        =   10
         Top             =   1080
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Fecha Desde:"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblBanco 
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   7
         Top             =   600
         Width           =   1750
         _Version        =   851968
         _ExtentX        =   3087
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Seleccione la Caja :"
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.TabControl TabSaldo 
      Height          =   705
      Left            =   60
      TabIndex        =   13
      Top             =   570
      Width           =   10815
      _Version        =   851968
      _ExtentX        =   19076
      _ExtentY        =   1244
      _StockProps     =   68
      PaintManager.FixedTabWidth=   0
      Begin XtremeSuiteControls.Label lblSaldo 
         Height          =   330
         Index           =   0
         Left            =   4170
         TabIndex        =   17
         Top             =   360
         Width           =   2685
         _Version        =   851968
         _ExtentX        =   4736
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "0.00"
         ForeColor       =   16744576
         BackColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblSaldo 
         Height          =   330
         Index           =   1
         Left            =   8160
         TabIndex        =   16
         Top             =   360
         Width           =   2655
         _Version        =   851968
         _ExtentX        =   4683
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "0.00"
         ForeColor       =   255
         BackColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin VB.Label lblTituloSaldo 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Anterior al"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   2730
         TabIndex        =   15
         Top             =   420
         Width           =   2535
      End
      Begin VB.Label lblTituloSaldo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Actual :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   7050
         TabIndex        =   14
         Top             =   420
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmCajaMovimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vIDCaja As Long
Dim rsCaja As ADODB.Recordset
Dim rsMovimientosCaja As ADODB.Recordset
Dim dbMovimientosCaja As ADODB.Recordset
Const vColorInicial = &H80C0FF
Dim vColor As String
Dim vSQLBusqueda As String
Private Sub FiltrarMovimientos()
On Error Resume Next

    Dim vsaldoanterior As Double

    If Not txtCaja(0).Text = "" Then
        
        Set rsMovimientosCaja = New ADODB.Recordset
        Dim sqlMovimientosCaja As String
        
        sqlMovimientosCaja = "SELECT DISTINCT  * FROM BancosMovimientos WHERE (idBancos = '" & txtCaja(0).Text & "') AND (Fecha >= '" & strfechaMySQL(txtFecha(0).Value) & "' and fecha <= '" & strfechaMySQL(txtFecha(1).Value) & "') ORDER BY fecha ASC, NroInterno ASC"
        
        With rsMovimientosCaja
            .CursorLocation = adUseServer
                        
            Call .Open(sqlMovimientosCaja, ConnDDBB, adOpenDynamic, adLockPessimistic)
            
            lblTituloSaldo(0).Caption = "Saldo anterior al " & txtFecha(0).Value
            vsaldoanterior = CalcularSaldoAnterior(txtFecha(0).Value)
            lblSaldo(0).Caption = Format(vsaldoanterior, "############0.00")
            
            
            CalcularSaldo (vsaldoanterior)  ' acá lo pasa al temporar que esta en la base listado
        
        End With

    Else
        
        FormatoGrilla (1)
    End If

    
If Err Then GrabarLog "FiltrarMovimientos", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Function CalcularSaldoAnterior(vFechaLimite As Date) As Double
On Error Resume Next

    CalcularSaldoAnterior = Val(GenerarDato("SELECT Sum(Debito),Sum(Credito),Sum(Debito)-Sum(Credito) as TSaldo FROM BancosMovimientos WHERE (idBancos  = '" & Trim(txtCaja(0).Text) & "') AND (Fecha < '" & strfechaMySQL(vFechaLimite) & "')", "TSaldo"))
    
    If Err Then
        GrabarLog "CalcularSaldoAnterior", Err.Number & " " & Err.Description, Me.Name
    End If
End Function
Private Sub CalcularSaldo(vSaldoParcial As Double)
On Error Resume Next

    PBFiltrar.Enabled = Not True
 
' ---- Variables ---------------
    Dim i As Integer
    Dim vsql, vvalores, vvcp As String
' ------------------------------

    
    vsql = "delete from Movimientoscaja"
    Call EjecutarScript(vsql, PathDBListados)  ' vacio la tabla
    

    
    With rsMovimientosCaja
        
        If Not .EOF = True Then
            .MoveFirst
            'fdesde.Value = strfechaMySQL(.Fields("Fecha").Value)
            FormatoGrilla (Val(GenerarDato("SELECT COUNT(idBancosMovimientos) as CantidadDeRegistros FROM BancosMovimientos WHERE (idBancos = '" & Trim(Me.txtCaja(0).Text) & "') AND (Fecha >= '" & strfechaMySQL(txtFecha(0).Text) & "' AND Fecha <= '" & strfechaMySQL(txtFecha(1).Text) & "')", "CantidadDeRegistros")))
        Else
            FormatoGrilla (1)
        End If
        
        i = 1
        
        
        Do Until .EOF = True
            
            vSaldoParcial = vSaldoParcial - Val(Format(.Fields("Credito").Value, "#######0.00")) + Val(Format(.Fields("debito").Value, "#######0.00"))
            
          '  vSaldoParcial = vSaldoParcial - Val(.Fields("Credito").Value) + Val(.Fields("Credito").Value)
                        
            vvalores = ""
            
            KlexMovimientos.RowHeight(i) = 240
            
            KlexMovimientos.TextMatrix(i, 0) = ""
            KlexMovimientos.TextMatrix(i, 1) = EsNulo(.Fields("idBancosMovimientos").Value)
            
            KlexMovimientos.TextMatrix(i, 2) = EsNulo(.Fields("Fecha").Value) 'Fecha
            vvalores = vvalores + "'" + strfecha2(.Fields("Fecha").Value) + "',"
            
            KlexMovimientos.TextMatrix(i, 3) = EsNulo(.Fields("NroInterno").Value) ' NroInterno
            vvalores = vvalores + EsNulo(.Fields("NroInterno").Value) + ","
            
            vsql = "select TipoMovimiento as cp from asientos where nrointerno=" + EsNulo(.Fields("NroInterno").Value)
            vvcp = traerDatos2(vsql, "cp", pathDBMySQL)
            vvalores = vvalores + "'" + vvcp + "'," ' TipoMovimiento
            
            'vsql = "select concat (CodigoProveedor,CodigoCliente) as cp from asientos where nrointerno=" + EsNulo(.Fields("NroInterno").Value)
            
            '--------------
            
            vsql = "select nombre from cuentascorrientes where nrointerno=" + EsNulo(.Fields("NroInterno").Value)

            vvcp = traerDatos2(vsql, "nombre", pathDBMySQL)
            
            vsql = "select nombre from pcuentascorrientes where nrointerno=" + EsNulo(.Fields("NroInterno").Value)

            vvcp = vvcp + traerDatos2(vsql, "nombre", pathDBMySQL)
            
            '----------------
            
            
            
            'vsql = "select nombre from clientes where codigo ='"++'""
            
            'vvCP = traerDatos2(vsql, "codigo", pathDBMySQL)
            
            vvalores = vvalores + "'" + vvcp + "'," ' TipoMovimiento
            
            
            KlexMovimientos.TextMatrix(i, 4) = Format(EsNulo(.Fields("Debito").Value), "###,###,##0.00")
            vvalores = vvalores + EsNulo(.Fields("Debito").Value) + "," ' Debito
            
            
            KlexMovimientos.TextMatrix(i, 5) = Format(EsNulo(.Fields("Credito").Value), "###,###,##0.00")
            vvalores = vvalores + EsNulo(.Fields("Credito").Value) + "," ' Credito
            
            KlexMovimientos.TextMatrix(i, 7) = EsNulo(.Fields("Comentario").Value)
            vvalores = vvalores + "'" + EsNulo(.Fields("Comentario").Value) + "'," ' Comentario
            
            vvalores = vvalores + "'" + EsNulo(.Fields("NroCheque").Value) + "'," ' nrocheque
            
            KlexMovimientos.TextMatrix(i, 6) = Format(Str(vSaldoParcial), "###,###,##0.00")
            vvalores = vvalores + Str(vSaldoParcial) ' Comentario

             '--------------
            vvcp = ""
            
            vsql = "select nombre from cuentascorrientes where nrointerno=" + EsNulo(.Fields("NroInterno").Value)

            vvcp = traerDatos2(vsql, "nombre", pathDBMySQL)
            
            vsql = "select nombre from pcuentascorrientes where nrointerno=" + EsNulo(.Fields("NroInterno").Value)

            vvcp = vvcp + traerDatos2(vsql, "nombre", pathDBMySQL)
            
            '----------------
            
            vvalores = vvalores + ",'" + vvcp + "'"
            
            
            
            
            'KlexMovimientos.TextMatrix(i, 8) = 0
            
            '.Fields("Saldo").Value = Val(Format(vSaldoParcial, "########0.00"))
            
           ' ---graba en el temporarl de mdb -------------
            vsql = "insert into MovimientosCaja (" + vCampoMovimientosCaja + ") values (" + vvalores + ")"
            Call EjecutarScript(vsql, PathDBListados)
            '--------------------------------------------------------
            
            .MoveNext
        
            i = i + 1
        Loop
        
        KlexMovimientos.TopRow = Val(KlexMovimientos.Rows - 1)
        .Fields.Refresh
        
        If Not .EOF = True Then .MoveLast
    
    
        lblSaldo(1).Caption = Format(vSaldoParcial, "############0.00")
        
        PBFiltrar.Enabled = True
    
    End With
    
    If Err Then
        GrabarLog "CalcularSaldo", Err.Number & " " & Err.Description, Me.Name
    End If
End Sub
Private Sub cmdVerDetalle_Click()
    On Error Resume Next

    If Not IsNull(bCajaMovimientos.Recordset("ncheque").Value) = True And Not Val(bCajaMovimientos.Recordset("ncheque").Value) = 0 Then
        
        With frmCheques
            .opModo(4).Value = True
            .cvnombre.Text = bCajaMovimientos.Recordset("codigo").Value
            .cvnombre_KeyPress 13
            .cvncheque.Text = bCajaMovimientos.Recordset("ncheque").Value
            .cmdBuscar_Click
        End With
    
    Else

        If bCajaMovimientos.Recordset("remito").Value > 0 Then
            frmBuscarCompra.vViene = "pctacte"
            frmBuscarCompra.vremito = (bCajaMovimientos.Recordset("remito").Value)
            frmBuscarCompra.Show
        Else
            MsgBox "El movimiento no fue realizado con documentos ", vbInformation, "Información..."
        End If
    End If

    If Err Then GrabarLog "cmdVerDetalle_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdNuevo_Click(Index As Integer)
On Error Resume Next

    If Index = 0 Then
        txtCaja(0).Text = ""
        txtCaja(1).Text = ""
        'txtImporte.Text = ""
        'txtComentario.Text = ""
        lblSaldo(0).Caption = ""
        lblSaldo(1).Caption = ""
        txtCaja(0).SetFocus
        
        With Me
            .Top = 300
            .Left = 300
            .Width = 10260
            .Height = 2500
        End With
    
    Else
    
        'txtImporte.Text = ""
        'txtComentario.Text = ""
        'txtImporte.SetFocus
        'If Not TabProveedor.Tab = 0 Then TabProveedor.Tab = 0

    End If

If Err Then GrabarLog "cmdNuevo_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdBorrar_Click()
    On Error Resume Next

    'If Not TabProveedor.Tab = 1 Then
    '    Exit Sub
    'End If
    
    If MsgBox("Confirma la baja del movimiento de Cuenta Corriente del Proveedor ? ", vbYesNo) = vbNo Then
        Exit Sub
    End If

    Dim vArreglo As Double, vSaldoProveedor As Double

    With bCajaMovimientos
        If Not (.Recordset.EOF = True) And Not (.Recordset.BOF = True) Then
            vArreglo = Val(Format(.Recordset("debito").Value, "#######0.00")) - Val(Format(.Recordset("credito").Value, "#######0.00"))
            lblSaldo(0).Caption = Trim(Val(lblSaldo(0).Caption) + vArreglo)
        
            .Recordset.Delete
            '.Refresh
        Else
            MsgBox "No tiene seleccionado ningun Movimiento...", vbExclamation, "Mensaje ...."
        End If
    
    End With

    CalcularSaldo (0)
    FormatoGrilla (0)
    
    If Err Then GrabarLog "cmdBorrar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Imprimir()
    On Error Resume Next
    
    Dim vsql As String
    
    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox "Prepare la impresora ...", vbInformation, "Mensaje ..."
      
    
     With drBancosMovimientos2
        .Sections("TituloEmpresa").Controls("lblTitulo").Caption = "Detalle de Movimientos de Caja"
        .Sections("TituloEmpresa").Controls("lblFechaDesde").Caption = txtFecha(0).Value
        .Sections("TituloEmpresa").Controls("lblFechaHasta").Caption = txtFecha(1).Value
        .Sections("TituloEmpresa").Controls("lblBanco").Caption = Trim(txtCaja(0).Text) & " - " & Trim(txtCaja(1).Text)
        .Sections("TituloEmpresa").Controls("lblSaldoAnterior").Caption = "$ " & Format(lblSaldo(0).Caption, "#######0.000")
        
        .Sections("PieInforme").Controls("lblSaldo").Caption = Format(lblSaldo(1).Caption, "#######0.000")
    
        .Show
    End With
    
    
    Exit Sub
    
    
    
    
    

    With Mantenimiento.rsBancosMovimientos
        If Not .State = 0 Then .Close
        
        If vSQLBusqueda = "" Then
            '.Source = "SELECT BM.idBancosMovimientos,bm.NroCheque, B.Descripcion, B.EsCaja, BC.idBancosCuentas, BC.Cuenta, BC.Descripcion, BM.Fecha, BM.Debito, BM.Credito, BM.Saldo,BM.Comentario,BM.TipoMovimiento, BM.NroInterno, BM.TipoMovimiento FROM BancosMovimientos BM INNER JOIN Bancos B ON BM.idBancos=B.idBancos LEFT JOIN BancosCuentas BC ON BM.idBancosCuentas=BC.idBancosCuentas WHERE (B.idBancos = '" & Trim(txtCaja(0).Text) & "') AND (Fecha >= '" & strfechaMySQL(txtFecha(0).Value) & "' and fecha <= '" & strfechaMySQL(txtFecha(1).Value) & "') ORDER BY fecha ASC, idBancosMovimientos ASC"
            '.Source = "SELECT BM.idBancosMovimientos,bm.NroCheque, B.Descripcion, B.EsCaja, BC.idBancosCuentas, BC.Cuenta, BC.Descripcion, BM.Fecha, BM.Debito, BM.Credito, BM.Saldo,BM.Comentario,BM.TipoMovimiento, BM.NroInterno, concat (aa.`CodigoProveedor`,aa.`CodigoCliente` ) as ClienteProveedor FROM BancosMovimientos BM INNER JOIN Bancos B ON BM.idBancos=B.idBancos LEFT JOIN BancosCuentas BC ON BM.idBancosCuentas=BC.idBancosCuentas  left join asientos aa on  bm.NroAsiento = aa.Numero WHERE (B.idBancos = '" & Trim(txtCaja(0).Text) & "') AND (bm.Fecha >= '" & strfechaMySQL(txtFecha(0).Value) & "' and bm.fecha <= '" & strfechaMySQL(txtFecha(1).Value) & "') ORDER BY bm.fecha ASC, bm.idBancosMovimientos ASC"
            
             If vDatosEmpresa.Alias = "Wgestion" Then
                    .Source = "SELECT DISTINCT  BM.idBancosMovimientos,bm.NroCheque, B.Descripcion, B.EsCaja, BC.idBancosCuentas, BC.Cuenta, BC.Descripcion, BM.Fecha, BM.Debito, BM.Credito, BM.Saldo,BM.Comentario,aa.TipoMovimiento, BM.NroInterno, concat (aa.`CodigoProveedor`,aa.`CodigoCliente` ) as ClienteProveedor FROM BancosMovimientos BM INNER JOIN Bancos B ON BM.idBancos=B.idBancos LEFT JOIN BancosCuentas BC ON BM.idBancosCuentas=BC.idBancosCuentas  left join asientos aa on  bm.NroInterno = aa.NroInterno WHERE (B.idBancos = '" & Trim(txtCaja(0).Text) & "') AND (bm.Fecha >= '" & strfechaMySQL(txtFecha(0).Value) & "' and bm.fecha <= '" & strfechaMySQL(txtFecha(1).Value) & "') ORDER BY bm.fecha ASC, bm.idBancosMovimientos ASC"
            Else
                    .Source = "SELECT DISTINCT BM.idBancosMovimientos,bm.NroCheque, B.Descripcion, B.EsCaja, BC.idBancosCuentas, BC.Cuenta, BC.Descripcion, BM.Fecha, BM.Debito, BM.Credito, BM.Saldo,BM.Comentario,aa.TipoMovimiento, BM.NroInterno, concat (aa.`CodigoProveedor`,aa.`CodigoCliente` ) as ClienteProveedor FROM BancosMovimientos BM INNER JOIN Bancos B ON BM.idBancos=B.idBancos LEFT JOIN BancosCuentas BC ON BM.idBancosCuentas=BC.idBancosCuentas  left join asientos aa on  bm.NroAsiento = aa.Numero WHERE (B.idBancos = '" & Trim(txtCaja(0).Text) & "') AND (bm.Fecha >= '" & strfechaMySQL(txtFecha(0).Value) & "' and bm.fecha <= '" & strfechaMySQL(txtFecha(1).Value) & "') ORDER BY bm.fecha ASC, bm.idBancosMovimientos ASC"
            End If

        Else
            If vDatosEmpresa.Alias = "Wgestion" Then ' esto es para que arneri encuentre con los nro interno los cp y leyenda de los asientos para bancosmovi
               .Source = "SELECT DISTINCT  BM.idBancosMovimientos,bm.NroCheque, B.idBancos, B.Descripcion, B.EsCaja, BC.idBancosCuentas, BC.Cuenta, BC.Descripcion, BM.Fecha, BM.Debito, BM.Credito, BM.Saldo,BM.Comentario,aa.TipoMovimiento, BM.NroInterno, concat (aa.`CodigoProveedor`,aa.`CodigoCliente` ) as ClienteProveedor FROM BancosMovimientos BM INNER JOIN Bancos B ON BM.idBancos=B.idBancos LEFT JOIN BancosCuentas BC ON BM.idBancosCuentas=BC.idBancosCuentas  left join asientos aa on  bm.NroInterno = aa.NroInterno    WHERE (" & Mid(vSQLBusqueda, 5, Len(vSQLBusqueda)) & ") ORDER BY bm.fecha ASC, bm.idBancosMovimientos ASC"
            Else
                .Source = "SELECT DISTINCT  BM.idBancosMovimientos,bm.NroCheque, B.idBancos, B.Descripcion, B.EsCaja, BC.idBancosCuentas, BC.Cuenta, BC.Descripcion, BM.Fecha, BM.Debito, BM.Credito, BM.Saldo,BM.Comentario,aa.TipoMovimiento, BM.NroInterno, concat (aa.`CodigoProveedor`,aa.`CodigoCliente` ) as ClienteProveedor FROM BancosMovimientos BM INNER JOIN Bancos B ON BM.idBancos=B.idBancos LEFT JOIN BancosCuentas BC ON BM.idBancosCuentas=BC.idBancosCuentas  left join asientos aa on  bm.Nroasiento = aa.numero    WHERE (" & Mid(vSQLBusqueda, 5, Len(vSQLBusqueda)) & ") ORDER BY bm.fecha ASC, bm.idBancosMovimientos ASC"
            End If
        
        End If
        
        If Not .State = 1 Then .Open
        .Close
        .Open

    End With

    With drBancosMovimientos
        .Sections("TituloEmpresa").Controls("lblTitulo").Caption = "Detalle de Movimientos de Caja"
        .Sections("TituloEmpresa").Controls("lblFechaDesde").Caption = txtFecha(0).Value
        .Sections("TituloEmpresa").Controls("lblFechaHasta").Caption = txtFecha(1).Value
        .Sections("TituloEmpresa").Controls("lblBanco").Caption = Trim(txtCaja(0).Text) & " - " & Trim(txtCaja(1).Text)
        .Sections("TituloEmpresa").Controls("lblSaldoAnterior").Caption = "$ " & Format(lblSaldo(0).Caption, "#######0.000")
        
        .Sections("PieInforme").Controls("lblSaldo").Caption = Format(lblSaldo(1).Caption, "#######0.000")
    
        .Show
    End With
    
    If Err Then GrabarLog "cmdImprimir_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, _
                       Shift As Integer)
On Error Resume Next

    If KeyCode = vbKeyF1 Then cmdNuevo_Click (0)
    If KeyCode = vbKeyF3 Then PbAcciones_Click (2)
    'If KeyCode = vbKeyF11 Then o1.Value = True
    'If KeyCode = vbKeyF12 Then o2.Value = True

If Err Then GrabarLog "Form_KeyUp", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Load()
    On Error Resume Next

    With Me
   '     .Show
   '     .Top = 0
   '     .Left = 0
   '     .Width = 11000
   '     .Height = 7440
        .KeyPreview = True
    End With
    
    'dtpAltaMovimiento.Value = Date
    txtFecha(0).Value = Date
    txtFecha(1).Value = Date
    lblSaldo(0).Caption = ""
    lblSaldo(1).Caption = ""

    init
    
    
    
    If Err Then GrabarLog "Form_load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub init()
If vConfigGral.vIncluyeContabilidad Then
    Me.PbAcciones(3).Enabled = False
Else
    Me.PbAcciones(3).Enabled = True
End If
Me.PbAcciones(3).Enabled = True
End Sub

Private Sub FormatoGrilla(vCantidadRenglones As Integer)
On Error Resume Next

    Dim i As Integer
    
    With KlexMovimientos
        .FixedRows = 1
        .FixedCols = 1
    
        .Cols = 9
        .Rows = vCantidadRenglones + 1
        
        If vCantidadRenglones = 1 Then
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
                .ColWidth(i) = 0
            Next
        End If
        
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 400
        
        .TextMatrix(0, 1) = "idBancosMovimientos"
        .ColWidth(1) = 0
               
        .TextMatrix(0, 2) = "Fecha"
        .ColWidth(2) = 1150
        
        .TextMatrix(0, 3) = "Nro Interno"
        .ColWidth(3) = 1000
        
        .TextMatrix(0, 4) = "Debito"
        .ColWidth(4) = 1250
        '.ColDisplayFormat(4) = "###,##0.00"
        .ColAlignment(4) = 9
        
        .TextMatrix(0, 5) = "Credito"
        .ColWidth(5) = 1250
        '.ColDisplayFormat(5) = "###,##0.00"
        .ColAlignment(5) = 9
        
        .TextMatrix(0, 6) = "Saldo"
        .ColWidth(6) = 1250
        '.ColDisplayFormat(6) = "###,##0.00"
        .ColAlignment(6) = 9
        
        .TextMatrix(0, 7) = "Observaciones"
        .ColWidth(7) = 3750

        .TextMatrix(0, 8) = ""
        .ColWidth(8) = 0

        '.BackColorAlternate = 14737632
    End With
    
If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub pbCarga_Click(Index As Integer)
 On Error Resume Next

    vVuelveBusqueda = Me.Name
    vVieneBusqueda = pbCarga(Index).Tag

    Select Case Index

        Case 0 To pbCarga.Count - 1
            frmBusqueda.Show
        
    End Select

    
    If Err Then GrabarLog "pbCarga_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub PBFiltrar_Click()
On Error Resume Next
    
    Me.lblSaldo(0).Cation = ""
    Me.lblSaldo(1).Caption = ""
   
MousePointer = vbHourglass
    FiltrarMovimientos
MousePointer = vbDefault

    Me.TabBancos.SelectedItem = 2
    
    vSQLBusqueda = ""
    vColor = vColorInicial
   ' Me.Height = 7440

If Err Then GrabarLog "PBFiltrar_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub PbAcciones_Click(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 0
            Imprimir
            
        Case 1
            Unload Me
            
        Case 2
            GBBusqueda.Visible = True
            txtBusqueda.SetFocus
       
        Case 3
        
        frmTransaccionMantenimiento.vnrointerno = Val(KlexMovimientos.TextMatrix(KlexMovimientos.RowSel, 3))
        frmTransaccionMantenimiento.Show

        Unload Me
        
        Case 4
    
    End Select
    
If Err Then GrabarLog "PBAcciones_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub MostrarCoincidencias(vBusqueda As String)
On Error Resume Next

    Dim sqlCaja As String

    Set rsCaja = New ADODB.Recordset

    If Trim(vBusqueda) = "" Then
        sqlCaja = "SELECT * FROM Caja WHERE 1=2"
    Else
        sqlCaja = "SELECT * FROM Caja WHERE (idBancos LIKE '%" & Trim(vBusqueda) & "%') OR (Descripcion LIKE '%" & Trim(vBusqueda) & "%')"
    End If

    With rsCaja
        If .State = 1 Then .Close

        .CursorLocation = adUseClient
    
        Call .Open(sqlCaja, ConnDDBB, adOpenStatic, adLockReadOnly)
    
        'dgBancos.Visible = Not .EOF
    
        If Not .EOF = True Then
            'Set dgBancos.DataSource = rsBancos
            Call FormatoGrilla(1)
        Else
            'Set dgBancos.DataSource = Nothing
        End If
    
    End With

    sqlCaja = ""

If Err Then GrabarLog "MostrarCoincidencias", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub PusArreglarSaldo_Click()

On Error Resume Next

Dim vdebito, vcredito As Double
Dim vsql, vvalores As String


vdebito = Val(InputBox("Debito para ajuste:", "Ajuste de Saldo"))
vcredito = Val(InputBox("Credito para ajuste:", "Ajuste de Saldo"))


If vdebito + vcredito = 0 Then Exit Sub


vvalores = " ('" + Trim(Me.txtCaja(0)) + "',0," + "'2011-01-01'," + Str(vdebito) + "," + Str(vcredito) + ",'Ajuste de Saldo')"

vsql = "insert into bancosmovimientos (idBancos,idBancosCuentas,fecha,debito,credito,comentario) Values " + vvalores

Call EjecutarScript(vsql, pathDBMySQL)

Call PBFiltrar_Click

If Err Then
    MsgBox "Error al intentar modificar saldo", vbCritical
End If


End Sub

Private Sub txtBusqueda_GotFocus()
On Error Resume Next

    With txtBusqueda
        .SelStart = Len(txtBusqueda.Text)
        .SelLength = Len(txtBusqueda.Text)
    End With

If Err Then GrabarLog "txtBusqueda_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
On Error Resume Next

    Dim i As Integer, j As Integer

    If KeyAscii = 13 Then
        
        If Not Trim(txtBusqueda.Text) = "" Then
        
            With KlexMovimientos
                .Row = 1
                For i = 1 To Val(.Rows - 1)
                    If (Val(.TextMatrix(i, 3)) = Val(txtBusqueda.Text)) Or (InStr(1, LCase(.TextMatrix(i, 7)), LCase(txtBusqueda.Text)) > 0) Then
                        .Row = i
                        
                        For j = 1 To Val(.Cols - 1)
                            .Col = j
                            .CellBackColor = vColor
                        Next
                        
                        vSQLBusqueda = vSQLBusqueda & " OR (idBancosMovimientos = " & Val(.TextMatrix(i, 1)) & ")"
                        
                    Else
                        If chkOcultarNoBuscados.Value = xtpChecked Then
                            .RowHeight(i) = 0
                        End If
                    End If
                
                Next
            
            
                Randomize
                vColor = Val(Rnd * vColorInicial)

            End With
            
        
        End If
        
        GBBusqueda.Visible = False
            
        
    End If

If Err Then GrabarLog "txtBusqueda_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtCaja_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        Call SeleccionarCaja(Index, txtCaja(Index).Text)
    End If

If Err Then GrabarLog "txtCaja_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Function SeleccionarCaja(Index As Integer, vValor As Variant)
On Error Resume Next

    vValor = Trim(vValor)
    
    Dim rsBanco As New ADODB.Recordset, sqlBanco As String
    
    Select Case Index
    
        Case 0
            sqlBanco = "SELECT * FROM Bancos WHERE (idBancos = '" & vValor & "')"
        
        Case 2
            sqlBanco = "SELECT * FROM Bancos WHERE (Cuenta = '" & vValor & "')"
    
    End Select
    
    With rsBanco
        Call .Open(sqlBanco, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then
            If Not vValor = "" Then
                Select Case Index
    
                    Case 0
                        txtCaja(0).Text = EsNulo(.Fields("idBancos").Value)
                        txtCaja(1).Text = EsNulo(.Fields("Descripcion").Value)
                        txtFecha(0).SetFocus
                    
                    Case 2
                        'txtCaja(2).Tag = EsNulo(.Fields("idBancosCuentas").Value)
                        'txtCaja(2).Text = EsNulo(.Fields("Cuenta").Value)
                        'txtCaja(3).Text = EsNulo(.Fields("Descripcion").Value)
                        '
    
                End Select

            End If
        End If
    End With
    
    sqlBanco = ""
    
    If rsBanco.State = 1 Then
        rsBanco.Close
        Set rsBanco = Nothing
    End If
    
If Err Then GrabarLog "SeleccionarCaja", Err.Number & " " & Err.Description, Me.Caption
End Function
Private Sub txtCaja_LostFocus(Index As Integer)
On Error Resume Next

'    Call SeleccionarCaja(Index, txtCaja(Index).Text)

If Err Then GrabarLog "txtCaja_LostFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        If Index = 0 Then txtFecha(Index + 1).SetFocus
        If Index = 1 Then PBFiltrar.SetFocus
        
    End If

If Err Then GrabarLog "txtFecha_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
