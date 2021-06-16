VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmImprimir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario de Impresion de Listados"
   ClientHeight    =   6585
   ClientLeft      =   3615
   ClientTop       =   -13410
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   9300
   Begin XtremeSuiteControls.TabControl TabImpresion 
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   9135
      _Version        =   851968
      _ExtentX        =   16113
      _ExtentY        =   7858
      _StockProps     =   68
      ItemCount       =   6
      Item(0).Caption =   "Intervalos"
      Item(0).ControlCount=   54
      Item(0).Control(0)=   "lblHasta(9)"
      Item(0).Control(1)=   "lblArticulos(10)"
      Item(0).Control(2)=   "txtArticulos(0)"
      Item(0).Control(3)=   "txtArticulos(1)"
      Item(0).Control(4)=   "txtArticulos(2)"
      Item(0).Control(5)=   "lblArticulos(0)"
      Item(0).Control(6)=   "lblHasta(10)"
      Item(0).Control(7)=   "txtArticulos(3)"
      Item(0).Control(8)=   "txtArticulos(4)"
      Item(0).Control(9)=   "lblArticulos(1)"
      Item(0).Control(10)=   "txtArticulos(5)"
      Item(0).Control(11)=   "lblHasta(11)"
      Item(0).Control(12)=   "txtArticulos(6)"
      Item(0).Control(13)=   "txtArticulos(7)"
      Item(0).Control(14)=   "chkActivarFecha(2)"
      Item(0).Control(15)=   "dtpAlta(2)"
      Item(0).Control(16)=   "dtpAlta(3)"
      Item(0).Control(17)=   "lblHasta(12)"
      Item(0).Control(18)=   "lblIntervalos(10)"
      Item(0).Control(19)=   "txtArticulos(8)"
      Item(0).Control(20)=   "lblArticulos(2)"
      Item(0).Control(21)=   "txtArticulos(9)"
      Item(0).Control(22)=   "lblHasta(13)"
      Item(0).Control(23)=   "txtArticulos(10)"
      Item(0).Control(24)=   "txtArticulos(11)"
      Item(0).Control(25)=   "txtArticulos(12)"
      Item(0).Control(26)=   "lblArticulos(3)"
      Item(0).Control(27)=   "txtArticulos(13)"
      Item(0).Control(28)=   "lblHasta(14)"
      Item(0).Control(29)=   "txtArticulos(14)"
      Item(0).Control(30)=   "txtArticulos(15)"
      Item(0).Control(31)=   "pbCargaArticulos(12)"
      Item(0).Control(32)=   "pbCargaArticulos(11)"
      Item(0).Control(33)=   "pbCargaArticulos(13)"
      Item(0).Control(34)=   "pbCargaArticulos(14)"
      Item(0).Control(35)=   "pbCargaArticulos(15)"
      Item(0).Control(36)=   "pbCargaArticulos(16)"
      Item(0).Control(37)=   "pbCargaArticulos(17)"
      Item(0).Control(38)=   "pbCargaArticulos(18)"
      Item(0).Control(39)=   "txtArticulos(16)"
      Item(0).Control(40)=   "pbCargaArticulos(0)"
      Item(0).Control(41)=   "txtArticulos(17)"
      Item(0).Control(42)=   "txtArticulos(18)"
      Item(0).Control(43)=   "pbCargaArticulos(1)"
      Item(0).Control(44)=   "txtArticulos(19)"
      Item(0).Control(45)=   "lblHasta(15)"
      Item(0).Control(46)=   "lblArticulos(4)"
      Item(0).Control(47)=   "chkArticulosPrecioConIva"
      Item(0).Control(48)=   "lblTipoIva(5)"
      Item(0).Control(49)=   "vpventa"
      Item(0).Control(50)=   "vorden"
      Item(0).Control(51)=   "lblTipoIva(0)"
      Item(0).Control(52)=   "GroTipoDe"
      Item(0).Control(53)=   "chkmd"
      Item(1).Caption =   "Intervalos"
      Item(1).ControlCount=   60
      Item(1).Control(0)=   "lblIntervalos(0)"
      Item(1).Control(1)=   "lblIntervalos(1)"
      Item(1).Control(2)=   "lblIntervalos(2)"
      Item(1).Control(3)=   "lblIntervalos(3)"
      Item(1).Control(4)=   "lblIntervalos(4)"
      Item(1).Control(5)=   "txtIntervalos(0)"
      Item(1).Control(6)=   "pbCarga(2)"
      Item(1).Control(7)=   "txtIntervalos(1)"
      Item(1).Control(8)=   "pbCarga(0)"
      Item(1).Control(9)=   "lblHasta(0)"
      Item(1).Control(10)=   "txtIntervalos(2)"
      Item(1).Control(11)=   "txtIntervalos(3)"
      Item(1).Control(12)=   "lblHasta(1)"
      Item(1).Control(13)=   "txtIntervalos(4)"
      Item(1).Control(14)=   "txtIntervalos(5)"
      Item(1).Control(15)=   "lblHasta(2)"
      Item(1).Control(16)=   "txtIntervalos(6)"
      Item(1).Control(17)=   "pbCarga(1)"
      Item(1).Control(18)=   "txtIntervalos(7)"
      Item(1).Control(19)=   "lblHasta(3)"
      Item(1).Control(20)=   "txtIntervalos(8)"
      Item(1).Control(21)=   "pbCarga(3)"
      Item(1).Control(22)=   "txtIntervalos(9)"
      Item(1).Control(23)=   "txtIntervalos(10)"
      Item(1).Control(24)=   "pbCarga(4)"
      Item(1).Control(25)=   "txtIntervalos(11)"
      Item(1).Control(26)=   "lblIntervalos(5)"
      Item(1).Control(27)=   "lblHasta(4)"
      Item(1).Control(28)=   "dtpAlta(0)"
      Item(1).Control(29)=   "dtpAlta(1)"
      Item(1).Control(30)=   "lblIntervalos(6)"
      Item(1).Control(31)=   "dtpNacimiento(0)"
      Item(1).Control(32)=   "lblHasta(5)"
      Item(1).Control(33)=   "dtpNacimiento(1)"
      Item(1).Control(34)=   "txtIntervalos(12)"
      Item(1).Control(35)=   "pbCarga(5)"
      Item(1).Control(36)=   "txtIntervalos(13)"
      Item(1).Control(37)=   "lblHasta(6)"
      Item(1).Control(38)=   "txtIntervalos(14)"
      Item(1).Control(39)=   "pbCarga(6)"
      Item(1).Control(40)=   "txtIntervalos(15)"
      Item(1).Control(41)=   "lblIntervalos(7)"
      Item(1).Control(42)=   "txtIntervalos(16)"
      Item(1).Control(43)=   "pbCarga(7)"
      Item(1).Control(44)=   "txtIntervalos(17)"
      Item(1).Control(45)=   "lblHasta(7)"
      Item(1).Control(46)=   "txtIntervalos(18)"
      Item(1).Control(47)=   "pbCarga(8)"
      Item(1).Control(48)=   "txtIntervalos(19)"
      Item(1).Control(49)=   "lblIntervalos(8)"
      Item(1).Control(50)=   "txtIntervalos(20)"
      Item(1).Control(51)=   "pbCarga(9)"
      Item(1).Control(52)=   "txtIntervalos(21)"
      Item(1).Control(53)=   "lblHasta(8)"
      Item(1).Control(54)=   "txtIntervalos(22)"
      Item(1).Control(55)=   "pbCarga(10)"
      Item(1).Control(56)=   "txtIntervalos(23)"
      Item(1).Control(57)=   "lblIntervalos(9)"
      Item(1).Control(58)=   "chkActivarFecha(0)"
      Item(1).Control(59)=   "chkActivarFecha(1)"
      Item(2).Caption =   "Intervalos"
      Item(2).ControlCount=   18
      Item(2).Control(0)=   "lblBancoCajaDetalle(0)"
      Item(2).Control(1)=   "pbCarga(11)"
      Item(2).Control(2)=   "txtBancoCajaDetalle(0)"
      Item(2).Control(3)=   "txtBancoCajaDetalle(1)"
      Item(2).Control(4)=   "lblHasta(18)"
      Item(2).Control(5)=   "txtBancoCajaDetalle(2)"
      Item(2).Control(6)=   "pbCarga(12)"
      Item(2).Control(7)=   "txtBancoCajaDetalle(3)"
      Item(2).Control(8)=   "lblBancoCajaDetalle(1)"
      Item(2).Control(9)=   "txtBancoCajaDetalle(4)"
      Item(2).Control(10)=   "pbCarga(13)"
      Item(2).Control(11)=   "txtBancoCajaDetalle(5)"
      Item(2).Control(12)=   "lblHasta(19)"
      Item(2).Control(13)=   "txtBancoCajaDetalle(6)"
      Item(2).Control(14)=   "pbCarga(14)"
      Item(2).Control(15)=   "txtBancoCajaDetalle(7)"
      Item(2).Control(16)=   "chkDiferido"
      Item(2).Control(17)=   "Frame1"
      Item(3).Caption =   "Intervalos"
      Item(3).ControlCount=   0
      Item(4).Caption =   "Plan de Cuentas"
      Item(4).ControlCount=   8
      Item(4).Control(0)=   "lblContabilidad(0)"
      Item(4).Control(1)=   "txtContabilidad(0)"
      Item(4).Control(2)=   "lblHasta(16)"
      Item(4).Control(3)=   "txtContabilidad(1)"
      Item(4).Control(4)=   "pbContabilidad(0)"
      Item(4).Control(5)=   "txtContabilidad(2)"
      Item(4).Control(6)=   "pbContabilidad(1)"
      Item(4).Control(7)=   "txtContabilidad(3)"
      Item(5).Caption =   "Opciones"
      Item(5).ControlCount=   4
      Item(5).Control(0)=   "gbAcciones"
      Item(5).Control(1)=   "gbClasificacion"
      Item(5).Control(2)=   "GBOrdenacion"
      Item(5).Control(3)=   "gbStock"
      Begin XtremeSuiteControls.CheckBox chkmd 
         Height          =   240
         Left            =   4590
         TabIndex        =   181
         Top             =   4005
         Width           =   2535
         _Version        =   851968
         _ExtentX        =   4471
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Mostrar Dirección"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroTipoDe 
         Height          =   615
         Left            =   4560
         TabIndex        =   177
         Top             =   3180
         Width           =   4395
         _Version        =   851968
         _ExtentX        =   7752
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Tipo de Interesado de este listado:"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton rbClientes 
            Height          =   255
            Left            =   180
            TabIndex        =   178
            Top             =   240
            Width           =   855
            _Version        =   851968
            _ExtentX        =   1508
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Clientes"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rdinterno 
            Height          =   255
            Left            =   1380
            TabIndex        =   179
            Top             =   240
            Width           =   885
            _Version        =   851968
            _ExtentX        =   1561
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Interno"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbRubro 
            Height          =   255
            Left            =   2520
            TabIndex        =   180
            Top             =   240
            Width           =   1755
            _Version        =   851968
            _ExtentX        =   3096
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Rubros - Articulos"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin VB.ComboBox vorden 
         Height          =   315
         ItemData        =   "frmImprimir.frx":0000
         Left            =   1680
         List            =   "frmImprimir.frx":0010
         TabIndex        =   175
         Text            =   "Codigo"
         Top             =   3990
         Width           =   2340
      End
      Begin VB.ComboBox vpventa 
         Height          =   315
         ItemData        =   "frmImprimir.frx":0035
         Left            =   1680
         List            =   "frmImprimir.frx":0048
         TabIndex        =   174
         Text            =   "Pventa1"
         Top             =   3180
         Width           =   2445
      End
      Begin VB.Frame Frame1 
         Height          =   675
         Left            =   -69700
         TabIndex        =   166
         Top             =   1710
         Visible         =   0   'False
         Width           =   8655
         Begin XtremeSuiteControls.CheckBox chkActivarFecha 
            Height          =   255
            Index           =   4
            Left            =   5820
            TabIndex        =   167
            Top             =   270
            Width           =   2655
            _Version        =   851968
            _ExtentX        =   4683
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Activar Intervalo por fecha de Alta"
            UseVisualStyle  =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtpBancoCajaMovimiento 
            Height          =   315
            Index           =   0
            Left            =   1650
            TabIndex        =   168
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   71106561
            CurrentDate     =   40291
         End
         Begin MSComCtl2.DTPicker dtpBancoCajaMovimiento 
            Height          =   315
            Index           =   1
            Left            =   3780
            TabIndex        =   169
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   71106561
            CurrentDate     =   40291
         End
         Begin XtremeSuiteControls.Label lblBancoCajaDetalle 
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   171
            Top             =   270
            Width           =   1350
            _Version        =   851968
            _ExtentX        =   2381
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "F. Movimiento:"
         End
         Begin XtremeSuiteControls.Label lblHasta 
            Height          =   195
            Index           =   20
            Left            =   3060
            TabIndex        =   170
            Top             =   300
            Width           =   495
            _Version        =   851968
            _ExtentX        =   882
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "a:"
            Alignment       =   2
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.CheckBox chkDiferido 
         Height          =   255
         Left            =   -69940
         TabIndex        =   165
         Top             =   4140
         Visible         =   0   'False
         Width           =   8955
         _Version        =   851968
         _ExtentX        =   15796
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Filtrar solamente valores diferidos"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkArticulosPrecioConIva 
         Height          =   375
         Left            =   1680
         TabIndex        =   148
         Top             =   3570
         Width           =   6705
         _Version        =   851968
         _ExtentX        =   11827
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Precio con IVA (Precio Final)"
         UseVisualStyle  =   -1  'True
         Value           =   1
      End
      Begin XtremeSuiteControls.PushButton pbContabilidad 
         Height          =   315
         Index           =   0
         Left            =   -67420
         TabIndex        =   142
         Tag             =   "CodigoCuentaD"
         Top             =   690
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtContabilidad 
         Height          =   315
         Index           =   0
         Left            =   -68320
         TabIndex        =   141
         Top             =   675
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.GroupBox gbStock 
         Height          =   1935
         Left            =   -63760
         TabIndex        =   87
         Top             =   600
         Visible         =   0   'False
         Width           =   2505
         _Version        =   851968
         _ExtentX        =   4419
         _ExtentY        =   3413
         _StockProps     =   79
         Caption         =   "Stock"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton rbStock 
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   88
            Top             =   360
            Width           =   2055
            _Version        =   851968
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Todos Los Articulos"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbStock 
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   89
            Top             =   600
            Width           =   2055
            _Version        =   851968
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Articulos Con Stock"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbStock 
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   90
            Top             =   840
            Width           =   2055
            _Version        =   851968
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Articulos Sin Stock"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbStock 
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   91
            Top             =   1080
            Width           =   2055
            _Version        =   851968
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Articulos Bajo Minimo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbStock 
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   92
            Top             =   1320
            Width           =   2055
            _Version        =   851968
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Articulos Sobre Maximo"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.CheckBox chkActivarFecha 
         Height          =   255
         Index           =   0
         Left            =   -64960
         TabIndex        =   68
         Top             =   2400
         Visible         =   0   'False
         Width           =   4000
         _Version        =   851968
         _ExtentX        =   7056
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Activar Intervalo por fecha de Alta"
         UseVisualStyle  =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtpAlta 
         Height          =   315
         Index           =   0
         Left            =   -68320
         TabIndex        =   31
         Top             =   2400
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   71106561
         CurrentDate     =   40291
      End
      Begin XtremeSuiteControls.FlatEdit txtIntervalos 
         Height          =   315
         Index           =   0
         Left            =   -68320
         TabIndex        =   9
         Top             =   675
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   1
         Left            =   -65500
         TabIndex        =   10
         Tag             =   "CodigoClienteH"
         Top             =   675
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtIntervalos 
         Height          =   315
         Index           =   1
         Left            =   -66400
         TabIndex        =   11
         Top             =   675
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   0
         Left            =   -67400
         TabIndex        =   12
         Tag             =   "CodigoClienteD"
         Top             =   675
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtIntervalos 
         Height          =   315
         Index           =   2
         Left            =   -68320
         TabIndex        =   14
         Top             =   1005
         Visible         =   0   'False
         Width           =   3135
         _Version        =   851968
         _ExtentX        =   5530
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtIntervalos 
         Height          =   315
         Index           =   3
         Left            =   -64240
         TabIndex        =   15
         Top             =   1005
         Visible         =   0   'False
         Width           =   3135
         _Version        =   851968
         _ExtentX        =   5530
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtIntervalos 
         Height          =   315
         Index           =   4
         Left            =   -68320
         TabIndex        =   17
         Top             =   1335
         Visible         =   0   'False
         Width           =   3135
         _Version        =   851968
         _ExtentX        =   5530
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtIntervalos 
         Height          =   315
         Index           =   5
         Left            =   -64240
         TabIndex        =   18
         Top             =   1330
         Visible         =   0   'False
         Width           =   3135
         _Version        =   851968
         _ExtentX        =   5530
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtIntervalos 
         Height          =   315
         Index           =   6
         Left            =   -68320
         TabIndex        =   20
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   2
         Left            =   -67405
         TabIndex        =   21
         Tag             =   "CodigoPostalD"
         Top             =   1680
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtIntervalos 
         Height          =   315
         Index           =   7
         Left            =   -67000
         TabIndex        =   22
         Top             =   1680
         Visible         =   0   'False
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtIntervalos 
         Height          =   315
         Index           =   8
         Left            =   -64240
         TabIndex        =   24
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   3
         Left            =   -63325
         TabIndex        =   25
         Tag             =   "CodigoPostalH"
         Top             =   1680
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtIntervalos 
         Height          =   315
         Index           =   9
         Left            =   -62920
         TabIndex        =   26
         Top             =   1680
         Visible         =   0   'False
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtIntervalos 
         Height          =   315
         Index           =   10
         Left            =   -68320
         TabIndex        =   27
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   4
         Left            =   -67405
         TabIndex        =   28
         Tag             =   "EstadoCliente"
         Top             =   2040
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtIntervalos 
         Height          =   315
         Index           =   11
         Left            =   -67000
         TabIndex        =   29
         Top             =   2040
         Visible         =   0   'False
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtpAlta 
         Height          =   315
         Index           =   1
         Left            =   -66520
         TabIndex        =   33
         Top             =   2400
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   71106561
         CurrentDate     =   40291
      End
      Begin MSComCtl2.DTPicker dtpNacimiento 
         Height          =   315
         Index           =   0
         Left            =   -68320
         TabIndex        =   35
         Top             =   2760
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   71106561
         CurrentDate     =   40291
      End
      Begin MSComCtl2.DTPicker dtpNacimiento 
         Height          =   315
         Index           =   1
         Left            =   -66520
         TabIndex        =   37
         Top             =   2760
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   71106561
         CurrentDate     =   40291
      End
      Begin XtremeSuiteControls.FlatEdit txtIntervalos 
         Height          =   315
         Index           =   12
         Left            =   -68320
         TabIndex        =   38
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   5
         Left            =   -67405
         TabIndex        =   39
         Tag             =   "TipoClienteD"
         Top             =   3120
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtIntervalos 
         Height          =   315
         Index           =   13
         Left            =   -67000
         TabIndex        =   40
         Top             =   3120
         Visible         =   0   'False
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtIntervalos 
         Height          =   315
         Index           =   14
         Left            =   -64240
         TabIndex        =   42
         Top             =   3120
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   6
         Left            =   -63325
         TabIndex        =   43
         Tag             =   "TipoClienteH"
         Top             =   3120
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtIntervalos 
         Height          =   315
         Index           =   15
         Left            =   -62920
         TabIndex        =   44
         Top             =   3120
         Visible         =   0   'False
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtIntervalos 
         Height          =   315
         Index           =   16
         Left            =   -68320
         TabIndex        =   46
         Top             =   3480
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   7
         Left            =   -67405
         TabIndex        =   47
         Tag             =   "ActividadD"
         Top             =   3480
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtIntervalos 
         Height          =   315
         Index           =   17
         Left            =   -67000
         TabIndex        =   48
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtIntervalos 
         Height          =   315
         Index           =   18
         Left            =   -64240
         TabIndex        =   50
         Top             =   3480
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   8
         Left            =   -63325
         TabIndex        =   51
         Tag             =   "ActividadH"
         Top             =   3480
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtIntervalos 
         Height          =   315
         Index           =   19
         Left            =   -62920
         TabIndex        =   52
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtIntervalos 
         Height          =   315
         Index           =   20
         Left            =   -68320
         TabIndex        =   54
         Top             =   3840
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   9
         Left            =   -67405
         TabIndex        =   55
         Tag             =   "VendedorD"
         Top             =   3840
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtIntervalos 
         Height          =   315
         Index           =   21
         Left            =   -67000
         TabIndex        =   56
         Top             =   3840
         Visible         =   0   'False
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtIntervalos 
         Height          =   315
         Index           =   22
         Left            =   -64240
         TabIndex        =   58
         Top             =   3840
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   10
         Left            =   -63325
         TabIndex        =   59
         Tag             =   "VendedorH"
         Top             =   3840
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtIntervalos 
         Height          =   315
         Index           =   23
         Left            =   -62920
         TabIndex        =   60
         Top             =   3840
         Visible         =   0   'False
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkActivarFecha 
         Height          =   255
         Index           =   1
         Left            =   -64960
         TabIndex        =   69
         Top             =   2760
         Visible         =   0   'False
         Width           =   4000
         _Version        =   851968
         _ExtentX        =   7056
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Activar Intervalo por fecha de Nacimiento"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox gbAcciones 
         Height          =   1215
         Left            =   -66760
         TabIndex        =   71
         Top             =   2640
         Visible         =   0   'False
         Width           =   2505
         _Version        =   851968
         _ExtentX        =   4410
         _ExtentY        =   2143
         _StockProps     =   79
         Caption         =   "Acciones"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.ComboBox cbTipoReporte 
            Height          =   315
            Left            =   120
            TabIndex        =   72
            Top             =   600
            Width           =   2235
            _Version        =   851968
            _ExtentX        =   3942
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Enabled         =   0   'False
         End
         Begin XtremeSuiteControls.Label lblAcciones 
            Height          =   195
            Left            =   120
            TabIndex        =   73
            Top             =   360
            Width           =   810
            _Version        =   851968
            _ExtentX        =   1429
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Imprimir:"
            Enabled         =   0   'False
            AutoEllipsis    =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox gbClasificacion 
         Height          =   1935
         Left            =   -66760
         TabIndex        =   74
         Top             =   600
         Visible         =   0   'False
         Width           =   2505
         _Version        =   851968
         _ExtentX        =   4410
         _ExtentY        =   3413
         _StockProps     =   79
         Caption         =   "Clasificacion"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton rbClasificacion 
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   75
            Top             =   360
            Width           =   2000
            _Version        =   851968
            _ExtentX        =   3528
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Sin Clasificar"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbClasificacion 
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   76
            Top             =   720
            Width           =   2000
            _Version        =   851968
            _ExtentX        =   3528
            _ExtentY        =   450
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbClasificacion 
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   77
            Top             =   1080
            Width           =   2000
            _Version        =   851968
            _ExtentX        =   3528
            _ExtentY        =   450
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbClasificacion 
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   78
            Top             =   1440
            Width           =   2000
            _Version        =   851968
            _ExtentX        =   3528
            _ExtentY        =   450
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GBOrdenacion 
         Height          =   3255
         Left            =   -69760
         TabIndex        =   79
         Top             =   600
         Visible         =   0   'False
         Width           =   2500
         _Version        =   851968
         _ExtentX        =   4410
         _ExtentY        =   5741
         _StockProps     =   79
         Caption         =   "Ordenacion:"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.CheckBox chkTipoOrden 
            Height          =   255
            Left            =   120
            TabIndex        =   80
            Top             =   2640
            Width           =   1995
            _Version        =   851968
            _ExtentX        =   3528
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Orden Descendente"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbOrden 
            Height          =   285
            Index           =   0
            Left            =   360
            TabIndex        =   81
            Top             =   360
            Width           =   2000
            _Version        =   851968
            _ExtentX        =   3528
            _ExtentY        =   503
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbOrden 
            Height          =   285
            Index           =   1
            Left            =   360
            TabIndex        =   82
            Top             =   720
            Width           =   2000
            _Version        =   851968
            _ExtentX        =   3528
            _ExtentY        =   503
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbOrden 
            Height          =   285
            Index           =   2
            Left            =   360
            TabIndex        =   83
            Top             =   1080
            Width           =   2000
            _Version        =   851968
            _ExtentX        =   3528
            _ExtentY        =   503
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbOrden 
            Height          =   285
            Index           =   3
            Left            =   360
            TabIndex        =   84
            Top             =   1440
            Width           =   2000
            _Version        =   851968
            _ExtentX        =   3528
            _ExtentY        =   503
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbOrden 
            Height          =   285
            Index           =   4
            Left            =   360
            TabIndex        =   85
            Top             =   1800
            Width           =   2000
            _Version        =   851968
            _ExtentX        =   3528
            _ExtentY        =   503
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbOrden 
            Height          =   285
            Index           =   5
            Left            =   360
            TabIndex        =   86
            Top             =   2160
            Width           =   2000
            _Version        =   851968
            _ExtentX        =   3528
            _ExtentY        =   503
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtArticulos 
         Height          =   315
         Index           =   0
         Left            =   1680
         TabIndex        =   93
         Top             =   675
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCargaArticulos 
         Height          =   315
         Index           =   12
         Left            =   4500
         TabIndex        =   94
         Tag             =   "CodigoArticuloH"
         Top             =   675
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtArticulos 
         Height          =   315
         Index           =   1
         Left            =   3600
         TabIndex        =   95
         Top             =   675
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCargaArticulos 
         Height          =   315
         Index           =   11
         Left            =   2600
         TabIndex        =   96
         Tag             =   "CodigoArticuloD"
         Top             =   675
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtArticulos 
         Height          =   315
         Index           =   2
         Left            =   1680
         TabIndex        =   99
         Top             =   1005
         Width           =   3135
         _Version        =   851968
         _ExtentX        =   5530
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtArticulos 
         Height          =   315
         Index           =   3
         Left            =   5760
         TabIndex        =   102
         Top             =   1005
         Width           =   3135
         _Version        =   851968
         _ExtentX        =   5530
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtArticulos 
         Height          =   315
         Index           =   4
         Left            =   1680
         TabIndex        =   103
         Top             =   1335
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCargaArticulos 
         Height          =   315
         Index           =   13
         Left            =   2600
         TabIndex        =   104
         Tag             =   "CodigoProveedorD"
         Top             =   1335
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtArticulos 
         Height          =   315
         Index           =   5
         Left            =   3000
         TabIndex        =   106
         Top             =   1335
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtArticulos 
         Height          =   315
         Index           =   6
         Left            =   5760
         TabIndex        =   108
         Top             =   1335
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCargaArticulos 
         Height          =   315
         Index           =   14
         Left            =   6675
         TabIndex        =   109
         Tag             =   "CodigoProveedorH"
         Top             =   1335
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtArticulos 
         Height          =   315
         Index           =   7
         Left            =   7080
         TabIndex        =   110
         Top             =   1335
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.CheckBox chkActivarFecha 
         Height          =   255
         Index           =   2
         Left            =   5040
         TabIndex        =   111
         Top             =   2415
         Width           =   4005
         _Version        =   851968
         _ExtentX        =   7056
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Activar Intervalo por fecha de Alta"
         UseVisualStyle  =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtpAlta 
         Height          =   315
         Index           =   2
         Left            =   1680
         TabIndex        =   112
         Top             =   2400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   71106561
         CurrentDate     =   40291
      End
      Begin MSComCtl2.DTPicker dtpAlta 
         Height          =   315
         Index           =   3
         Left            =   3480
         TabIndex        =   113
         Top             =   2400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   71106561
         CurrentDate     =   40291
      End
      Begin XtremeSuiteControls.FlatEdit txtArticulos 
         Height          =   315
         Index           =   8
         Left            =   1680
         TabIndex        =   116
         Top             =   1680
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCargaArticulos 
         Height          =   315
         Index           =   15
         Left            =   2600
         TabIndex        =   117
         Tag             =   "RubroD"
         Top             =   1680
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtArticulos 
         Height          =   315
         Index           =   9
         Left            =   3000
         TabIndex        =   119
         Top             =   1680
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtArticulos 
         Height          =   315
         Index           =   10
         Left            =   5760
         TabIndex        =   121
         Top             =   1680
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCargaArticulos 
         Height          =   315
         Index           =   16
         Left            =   6675
         TabIndex        =   122
         Tag             =   "RubroH"
         Top             =   1680
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtArticulos 
         Height          =   315
         Index           =   11
         Left            =   7080
         TabIndex        =   123
         Top             =   1680
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtArticulos 
         Height          =   315
         Index           =   12
         Left            =   1680
         TabIndex        =   124
         Top             =   2040
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCargaArticulos 
         Height          =   315
         Index           =   17
         Left            =   2595
         TabIndex        =   125
         Tag             =   "SubRubroD"
         Top             =   2040
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtArticulos 
         Height          =   315
         Index           =   13
         Left            =   3000
         TabIndex        =   127
         Top             =   2040
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtArticulos 
         Height          =   315
         Index           =   14
         Left            =   5760
         TabIndex        =   129
         Top             =   2040
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCargaArticulos 
         Height          =   315
         Index           =   18
         Left            =   6675
         TabIndex        =   130
         Tag             =   "SubRubroH"
         Top             =   2040
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtArticulos 
         Height          =   315
         Index           =   15
         Left            =   7080
         TabIndex        =   131
         Top             =   2040
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtArticulos 
         Height          =   315
         Index           =   16
         Left            =   1680
         TabIndex        =   132
         Top             =   2760
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCargaArticulos 
         Height          =   315
         Index           =   0
         Left            =   2595
         TabIndex        =   133
         Tag             =   "PorcentajeIvaD"
         Top             =   2760
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtArticulos 
         Height          =   315
         Index           =   17
         Left            =   3000
         TabIndex        =   134
         Top             =   2760
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtArticulos 
         Height          =   315
         Index           =   18
         Left            =   5760
         TabIndex        =   135
         Top             =   2760
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCargaArticulos 
         Height          =   315
         Index           =   1
         Left            =   6675
         TabIndex        =   136
         Tag             =   "PorcentajeIvaH"
         Top             =   2760
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtArticulos 
         Height          =   315
         Index           =   19
         Left            =   7080
         TabIndex        =   137
         Top             =   2760
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtContabilidad 
         Height          =   315
         Index           =   1
         Left            =   -67000
         TabIndex        =   144
         Top             =   675
         Visible         =   0   'False
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtContabilidad 
         Height          =   315
         Index           =   2
         Left            =   -64240
         TabIndex        =   145
         Top             =   675
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbContabilidad 
         Height          =   315
         Index           =   1
         Left            =   -63325
         TabIndex        =   146
         Tag             =   "CodigoCuentaH"
         Top             =   675
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtContabilidad 
         Height          =   315
         Index           =   3
         Left            =   -62920
         TabIndex        =   147
         Top             =   675
         Visible         =   0   'False
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtBancoCajaDetalle 
         Height          =   315
         Index           =   0
         Left            =   -68320
         TabIndex        =   150
         Top             =   675
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   11
         Left            =   -67400
         TabIndex        =   151
         Tag             =   "BancoD"
         Top             =   675
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtBancoCajaDetalle 
         Height          =   315
         Index           =   1
         Left            =   -67000
         TabIndex        =   152
         Top             =   675
         Visible         =   0   'False
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtBancoCajaDetalle 
         Height          =   315
         Index           =   2
         Left            =   -64240
         TabIndex        =   154
         Top             =   675
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   12
         Left            =   -63325
         TabIndex        =   155
         Tag             =   "BancoH"
         Top             =   675
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtBancoCajaDetalle 
         Height          =   315
         Index           =   3
         Left            =   -62920
         TabIndex        =   156
         Top             =   675
         Visible         =   0   'False
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtBancoCajaDetalle 
         Height          =   315
         Index           =   4
         Left            =   -68320
         TabIndex        =   158
         Top             =   1275
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   13
         Left            =   -67405
         TabIndex        =   159
         Tag             =   "BancoCuentaD"
         Top             =   1275
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtBancoCajaDetalle 
         Height          =   315
         Index           =   5
         Left            =   -67000
         TabIndex        =   160
         Top             =   1275
         Visible         =   0   'False
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtBancoCajaDetalle 
         Height          =   315
         Index           =   6
         Left            =   -64240
         TabIndex        =   162
         Top             =   1275
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   14
         Left            =   -63325
         TabIndex        =   163
         Tag             =   "BancoCuentaH"
         Top             =   1275
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtBancoCajaDetalle 
         Height          =   315
         Index           =   7
         Left            =   -62920
         TabIndex        =   164
         Top             =   1275
         Visible         =   0   'False
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.Label lblTipoIva 
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   176
         Top             =   4050
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Ordenar listado por: "
         Transparent     =   -1  'True
         AutoEllipsis    =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblTipoIva 
         Height          =   195
         Index           =   5
         Left            =   210
         TabIndex        =   173
         Top             =   3240
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Lista de Precio:"
         Transparent     =   -1  'True
         AutoEllipsis    =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblHasta 
         Height          =   195
         Index           =   19
         Left            =   -64960
         TabIndex        =   161
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
         _Version        =   851968
         _ExtentX        =   882
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "a:"
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblBancoCajaDetalle 
         Height          =   195
         Index           =   1
         Left            =   -69820
         TabIndex        =   157
         Top             =   1320
         Visible         =   0   'False
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Cuenta"
      End
      Begin XtremeSuiteControls.Label lblHasta 
         Height          =   195
         Index           =   18
         Left            =   -64960
         TabIndex        =   153
         Top             =   720
         Visible         =   0   'False
         Width           =   495
         _Version        =   851968
         _ExtentX        =   882
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "a:"
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblBancoCajaDetalle 
         Height          =   195
         Index           =   0
         Left            =   -69820
         TabIndex        =   149
         Top             =   720
         Visible         =   0   'False
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Banco"
      End
      Begin XtremeSuiteControls.Label lblHasta 
         Height          =   195
         Index           =   16
         Left            =   -64960
         TabIndex        =   143
         Top             =   750
         Visible         =   0   'False
         Width           =   495
         _Version        =   851968
         _ExtentX        =   882
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "hasta:"
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblContabilidad 
         Height          =   195
         Index           =   0
         Left            =   -69820
         TabIndex        =   140
         Top             =   720
         Visible         =   0   'False
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Cuenta Desde :"
      End
      Begin XtremeSuiteControls.Label lblArticulos 
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   139
         Top             =   2805
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Tipo Iva :"
         Transparent     =   -1  'True
         AutoEllipsis    =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblHasta 
         Height          =   195
         Index           =   15
         Left            =   5040
         TabIndex        =   138
         Top             =   2805
         Width           =   495
         _Version        =   851968
         _ExtentX        =   882
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "a:"
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblHasta 
         Height          =   195
         Index           =   14
         Left            =   5040
         TabIndex        =   128
         Top             =   2085
         Width           =   495
         _Version        =   851968
         _ExtentX        =   882
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "a:"
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblArticulos 
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   126
         Top             =   2085
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "SubRubro :"
         Transparent     =   -1  'True
         AutoEllipsis    =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblHasta 
         Height          =   195
         Index           =   13
         Left            =   5040
         TabIndex        =   120
         Top             =   1720
         Width           =   495
         _Version        =   851968
         _ExtentX        =   882
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "a:"
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblArticulos 
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   118
         Top             =   1720
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Rubro :"
         Transparent     =   -1  'True
         AutoEllipsis    =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblIntervalos 
         Height          =   195
         Index           =   10
         Left            =   180
         TabIndex        =   115
         Top             =   2445
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "F. de Alta:"
         Transparent     =   -1  'True
         AutoEllipsis    =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblHasta 
         Height          =   195
         Index           =   12
         Left            =   3000
         TabIndex        =   114
         Top             =   2445
         Width           =   495
         _Version        =   851968
         _ExtentX        =   882
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "a:"
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblHasta 
         Height          =   195
         Index           =   11
         Left            =   5040
         TabIndex        =   107
         Top             =   1370
         Width           =   495
         _Version        =   851968
         _ExtentX        =   882
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "a:"
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblArticulos 
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   105
         Top             =   1370
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Proveedor :"
         Transparent     =   -1  'True
         AutoEllipsis    =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblHasta 
         Height          =   195
         Index           =   10
         Left            =   5040
         TabIndex        =   101
         Top             =   1045
         Width           =   495
         _Version        =   851968
         _ExtentX        =   882
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "a:"
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblArticulos 
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   100
         Top             =   1045
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Descripcion :"
         Transparent     =   -1  'True
         AutoEllipsis    =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblArticulos 
         Height          =   195
         Index           =   10
         Left            =   180
         TabIndex        =   98
         Top             =   720
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Codigo :"
         Transparent     =   -1  'True
         AutoEllipsis    =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblHasta 
         Height          =   195
         Index           =   9
         Left            =   3000
         TabIndex        =   97
         Top             =   720
         Width           =   495
         _Version        =   851968
         _ExtentX        =   882
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "a:"
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblIntervalos 
         Height          =   195
         Index           =   9
         Left            =   -69880
         TabIndex        =   61
         Top             =   3885
         Visible         =   0   'False
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Vendedor:"
         Transparent     =   -1  'True
         AutoEllipsis    =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblHasta 
         Height          =   195
         Index           =   8
         Left            =   -64960
         TabIndex        =   57
         Top             =   3885
         Visible         =   0   'False
         Width           =   495
         _Version        =   851968
         _ExtentX        =   882
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "a:"
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblIntervalos 
         Height          =   195
         Index           =   8
         Left            =   -69880
         TabIndex        =   53
         Top             =   3525
         Visible         =   0   'False
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Actividad:"
         Transparent     =   -1  'True
         AutoEllipsis    =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblHasta 
         Height          =   195
         Index           =   7
         Left            =   -64960
         TabIndex        =   49
         Top             =   3525
         Visible         =   0   'False
         Width           =   495
         _Version        =   851968
         _ExtentX        =   882
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "a:"
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblIntervalos 
         Height          =   195
         Index           =   7
         Left            =   -69880
         TabIndex        =   45
         Top             =   3160
         Visible         =   0   'False
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Tipo:"
         Transparent     =   -1  'True
         AutoEllipsis    =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblHasta 
         Height          =   195
         Index           =   6
         Left            =   -64960
         TabIndex        =   41
         Top             =   3165
         Visible         =   0   'False
         Width           =   495
         _Version        =   851968
         _ExtentX        =   882
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "a:"
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblHasta 
         Height          =   195
         Index           =   5
         Left            =   -67000
         TabIndex        =   36
         Top             =   2805
         Visible         =   0   'False
         Width           =   495
         _Version        =   851968
         _ExtentX        =   882
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "a:"
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblIntervalos 
         Height          =   195
         Index           =   6
         Left            =   -69880
         TabIndex        =   34
         Top             =   2800
         Visible         =   0   'False
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "F. de Nacimiento:"
         Transparent     =   -1  'True
         AutoEllipsis    =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblHasta 
         Height          =   195
         Index           =   4
         Left            =   -67000
         TabIndex        =   32
         Top             =   2440
         Visible         =   0   'False
         Width           =   495
         _Version        =   851968
         _ExtentX        =   882
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "a:"
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblIntervalos 
         Height          =   195
         Index           =   5
         Left            =   -69850
         TabIndex        =   30
         Top             =   2445
         Visible         =   0   'False
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "F. de Alta:"
         Transparent     =   -1  'True
         AutoEllipsis    =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblHasta 
         Height          =   195
         Index           =   3
         Left            =   -64960
         TabIndex        =   23
         Top             =   1720
         Visible         =   0   'False
         Width           =   495
         _Version        =   851968
         _ExtentX        =   882
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "a:"
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblHasta 
         Height          =   195
         Index           =   2
         Left            =   -64960
         TabIndex        =   19
         Top             =   1370
         Visible         =   0   'False
         Width           =   495
         _Version        =   851968
         _ExtentX        =   882
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "a:"
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblHasta 
         Height          =   195
         Index           =   1
         Left            =   -64960
         TabIndex        =   16
         Top             =   1045
         Visible         =   0   'False
         Width           =   495
         _Version        =   851968
         _ExtentX        =   882
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "a:"
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblHasta 
         Height          =   195
         Index           =   0
         Left            =   -67000
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   500
         _Version        =   851968
         _ExtentX        =   882
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "a:"
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblIntervalos 
         Height          =   195
         Index           =   4
         Left            =   -69820
         TabIndex        =   6
         Top             =   2070
         Visible         =   0   'False
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Estado del Cliente:"
         Transparent     =   -1  'True
         AutoEllipsis    =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblIntervalos 
         Height          =   195
         Index           =   3
         Left            =   -69850
         TabIndex        =   5
         Top             =   1680
         Visible         =   0   'False
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Codigo Postal:"
         Transparent     =   -1  'True
         AutoEllipsis    =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblIntervalos 
         Height          =   195
         Index           =   2
         Left            =   -69820
         TabIndex        =   4
         Top             =   1350
         Visible         =   0   'False
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Localidad :"
         Transparent     =   -1  'True
         AutoEllipsis    =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblIntervalos 
         Height          =   195
         Index           =   1
         Left            =   -69820
         TabIndex        =   3
         Top             =   1050
         Visible         =   0   'False
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Nombre :"
         Transparent     =   -1  'True
         AutoEllipsis    =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblIntervalos 
         Height          =   195
         Index           =   0
         Left            =   -69820
         TabIndex        =   2
         Top             =   720
         Visible         =   0   'False
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Codigo :"
         Transparent     =   -1  'True
         AutoEllipsis    =   -1  'True
      End
   End
   Begin MSDataGridLib.DataGrid dgFiltro 
      Height          =   855
      Left            =   2400
      TabIndex        =   70
      Top             =   360
      Visible         =   0   'False
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   1508
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   13
      RowDividerStyle =   4
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicInferior 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      Picture         =   "frmImprimir.frx":0079
      ScaleHeight     =   555
      ScaleWidth      =   9405
      TabIndex        =   62
      Top             =   6000
      Width           =   9400
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   0
         Left            =   4890
         TabIndex        =   63
         Top             =   120
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2558
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Vista Previa"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmImprimir.frx":512C
         BorderGap       =   10
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   2
         Left            =   7800
         TabIndex        =   64
         Top             =   120
         Width           =   1450
         _Version        =   851968
         _ExtentX        =   2558
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Cerrar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmImprimir.frx":B98E
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   1
         Left            =   6360
         TabIndex        =   67
         Top             =   120
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2558
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmImprimir.frx":BD8E
         BorderGap       =   10
      End
      Begin VB.Label lblWGestion 
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
         TabIndex        =   65
         Top             =   150
         Width           =   1770
      End
      Begin VB.Label lblWGestion 
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
         TabIndex        =   66
         Top             =   170
         Width           =   1770
      End
   End
   Begin XtremeSuiteControls.GroupBox gbImpresora 
      Height          =   1395
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   9255
      _Version        =   851968
      _ExtentX        =   16325
      _ExtentY        =   2461
      _StockProps     =   79
      Caption         =   "Tipo Salida del Informe"
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.RadioButton RbTipoSalida 
         Height          =   195
         Index           =   0
         Left            =   510
         TabIndex        =   7
         Top             =   330
         Width           =   3135
         _Version        =   851968
         _ExtentX        =   5530
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Impresora"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RbTipoSalida 
         Height          =   225
         Index           =   1
         Left            =   510
         TabIndex        =   8
         Top             =   930
         Width           =   1605
         _Version        =   851968
         _ExtentX        =   2831
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Archivo"
         Enabled         =   0   'False
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RbTipoSalida 
         Height          =   255
         Index           =   2
         Left            =   510
         TabIndex        =   172
         Top             =   600
         Width           =   3135
         _Version        =   851968
         _ExtentX        =   5530
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Pantalla"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
   End
End
Attribute VB_Name = "frmImprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsFiltro As New ADODB.Recordset, sqlFiltro As String
Dim vsql As String
Dim vSQLEncabezado As String, vSQLDetalle As String
Private Sub cbTipoReporte_GotFocus()
On Error Resume Next
    

If Err Then GrabarLog "cbTipoReporte_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub chkActivarFecha_Click(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 0
            dtpAlta(0).Enabled = CBool(chkActivarFecha(Index).Value)
            dtpAlta(1).Enabled = CBool(chkActivarFecha(Index).Value)
            
            If chkActivarFecha(Index).Value = xtpChecked Then
                dtpAlta(0).Value = Date
                dtpAlta(1).Value = Date
            Else
                dtpAlta(0).Value = dtpAlta(0).MinDate
                dtpAlta(1).Value = dtpAlta(1).MaxDate
            End If
    
        Case 1
            dtpNacimiento(0).Enabled = CBool(chkActivarFecha(Index).Value)
            dtpNacimiento(1).Enabled = CBool(chkActivarFecha(Index).Value)
            
            If chkActivarFecha(Index).Value = xtpChecked Then
                dtpNacimiento(0).Value = Date
                dtpNacimiento(1).Value = Date
            Else
                dtpNacimiento(0).Value = dtpNacimiento(0).MinDate
                dtpNacimiento(1).Value = dtpNacimiento(1).MaxDate
            End If
        
        Case 2
            dtpAlta(2).Enabled = CBool(chkActivarFecha(Index).Value)
            dtpAlta(3).Enabled = CBool(chkActivarFecha(Index).Value)
            
            If chkActivarFecha(Index).Value = xtpChecked Then
                dtpAlta(2).Value = Date
                dtpAlta(3).Value = Date
            Else
                dtpAlta(2).Value = dtpAlta(2).MinDate
                dtpAlta(3).Value = dtpAlta(3).MaxDate
            End If
    
        Case 3
            'dtpContabilidad(0).Enabled = CBool(chkActivarFecha(Index).Value)
            'dtpContabilidad(1).Enabled = CBool(chkActivarFecha(Index).Value)
        
           ' If chkActivarFecha(Index).Value = xtpChecked Then
           '     dtpContabilidad(0).Value = Date
           '     dtpContabilidad(1).Value = Date
           ' Else
           '     dtpContabilidad(0).Value = dtpAlta(2).MinDate
           '     dtpContabilidad(1).Value = dtpAlta(3).MaxDate
           ' End If
    
    
        Case 4
            dtpBancoCajaMovimiento(0).Enabled = CBool(chkActivarFecha(Index).Value)
            dtpBancoCajaMovimiento(1).Enabled = CBool(chkActivarFecha(Index).Value)
        
            If chkActivarFecha(Index).Value = xtpChecked Then
                dtpBancoCajaMovimiento(0).Value = Date
                dtpBancoCajaMovimiento(1).Value = Date
            Else
                dtpBancoCajaMovimiento(0).Value = dtpAlta(2).MinDate
                dtpBancoCajaMovimiento(1).Value = dtpAlta(3).MaxDate
            End If
        
    End Select

If Err Then GrabarLog "chkActivarFecha_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Load()
On Error Resume Next

    Me.Show

    chkActivarFecha_Click (0)
    chkActivarFecha_Click (1)
    chkActivarFecha_Click (2)
    chkActivarFecha_Click (3)
    chkActivarFecha_Click (4)

    SeleccionarModelo
    
    Me.vorden.Text = "Descrip"
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub PbAcciones_Click(Index As Integer)
On Error Resume Next
    
    Select Case Index
    
        Case 0, 1
            
            If Me.rbRubro Then
                drLARubro.Show
                Exit Sub
            End If
            
            
            Call ArmarSQL
            
            Call CargarGrilla
            
            If Index = 0 Then
                Call Imprimir(False)
            Else
                Call Imprimir(True)
            End If
        
        Case 2
            Unload Me
    
    End Select

If Err Then GrabarLog "PbAcciones_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub pbCarga_Click(Index As Integer)
On Error Resume Next

    vVuelveBusqueda = Me.Name
    vVieneBusqueda = pbCarga(Index).Tag
    
    frmBusqueda.Show

If Err Then GrabarLog "pbCarga_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub pbCargaArticulos_Click(Index As Integer)
On Error Resume Next

    vVuelveBusqueda = Me.Name
    vVieneBusqueda = pbCargaArticulos(Index).Tag
    
    frmBusqueda.Show

If Err Then GrabarLog "pbCargaArticulos_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub ArmarSQL()
On Error Resume Next
    
    vsql = ""
    vSQLEncabezado = ""
    vSQLDetalle = ""
    
    Select Case vVieneImpresion
        
        Case "frmCheques"
            If Not Trim(txtBancoCajaDetalle(0).Text) = "" And Not Trim(txtBancoCajaDetalle(2).Text) = "" Then
                vSQLEncabezado = vSQLEncabezado + " WHERE (B.idBancos >= '" & Trim(txtBancoCajaDetalle(0).Text) & "' AND B.idBancos  <= '" & Trim(txtBancoCajaDetalle(2).Text) & "')"
                
                If Not Trim(txtBancoCajaDetalle(4).Text) = "" And Not Trim(txtBancoCajaDetalle(6).Text) = "" Then
                    vSQLEncabezado = vSQLEncabezado + " AND (BC.idBancosCuentas >= " & Trim(txtBancoCajaDetalle(4).Text) & " AND BC.idBancosCuentas  <= " & Trim(txtBancoCajaDetalle(6).Text) & ")"
                End If
            
            End If
        
            
            If Not chkDiferido.Value = xtpChecked Then
                If chkActivarFecha(4).Value = xtpChecked Then
                    vSQLDetalle = " WHERE (Fecha >= '" & strfechaMySQL(dtpBancoCajaMovimiento(0).Value) & "' AND Fecha <= '" & strfechaMySQL(dtpBancoCajaMovimiento(1).Value) & "')"
                End If
            Else
                vSQLDetalle = " WHERE (FechaAcreditacion >= '" & strfechaMySQL(Date) & "')"
            End If
            
        Case "frmClientes", "frmProveedores"
            If Not Trim(txtIntervalos(0).Text) = "" And Not Trim(txtIntervalos(1).Text) = "" Then
                vsql = vsql + " AND (Codigo >= '" & Trim(txtIntervalos(0).Text) & "' AND Codigo <= '" & Trim(txtIntervalos(1).Text) & "')"
            End If

            If Not Trim(txtIntervalos(2).Text) = "" And Not Trim(txtIntervalos(3).Text) = "" Then
                vsql = vsql + " AND (Nombre >= '" & Trim(txtIntervalos(2).Text) & "%' AND Nombre  <= '" & Trim(txtIntervalos(3).Text) & "%')"
            End If
            
            If Not Trim(txtIntervalos(4).Text) = "" And Not Trim(txtIntervalos(5).Text) = "" Then
                vsql = vsql + " AND (Localidad >= '" & Trim(txtIntervalos(4).Text) & "%' AND Localidad  <=  '" & Trim(txtIntervalos(5).Text) & "%')"
            End If
        
            If Not Trim(txtIntervalos(6).Text) = "" And Not Trim(txtIntervalos(8).Text) = "" Then
                vsql = vsql + " AND (CodigoPostal >= '" & Trim(txtIntervalos(6).Text) & "' AND CodigoPostal <= '" & Trim(txtIntervalos(8).Text) & "')"
            End If
                
            If Not Trim(txtIntervalos(10).Text) = "" Then
                vsql = vsql + " AND (idEstados = '" & Trim(txtIntervalos(10).Text) & "')" ' OR IS NULL idEstados OR idEstados IS NULL OR idEstados = ''))"
            End If
                        
            If chkActivarFecha(0).Value = xtpChecked Then
                vsql = vsql + " AND (Fecha_Alta >= '" & strfechaMySQL(dtpAlta(0).Value) & "' AND Fecha_Alta <= '" & strfechaMySQL(dtpAlta(1).Value) & "')"
            End If
            
            If chkActivarFecha(1).Value = xtpChecked Then
                vsql = vsql + " AND (Fecha_Nacimiento >= '" & strfechaMySQL(dtpNacimiento(0).Value) & "' AND Fecha_Nacimiento <= '" & strfechaMySQL(dtpNacimiento(1).Value) & "')"
            End If
            
            If Not Trim(txtIntervalos(12).Text) = "" And Not Trim(txtIntervalos(14).Text) = "" Then
                vsql = vsql + " AND (idTipoCliente >= '" & Trim(txtIntervalos(12).Text) & "' AND idTipoCliente <= '" & Trim(txtIntervalos(14).Text) & "')"
            End If
                
            If Not Trim(txtIntervalos(16).Text) = "" And Not Trim(txtIntervalos(18).Text) = "" Then
                vsql = vsql + " AND (idActividad >= '" & Trim(txtIntervalos(16).Text) & "' AND idActividad <= '" & Trim(txtIntervalos(18).Text) & "')"
            End If
            
            If Not Trim(txtIntervalos(20).Text) = "" And Not Trim(txtIntervalos(22).Text) = "" Then
                vsql = vsql + " AND (idVendedor >= '" & Trim(txtIntervalos(20).Text) & "' AND idVendedor <= '" & Trim(txtIntervalos(22).Text) & "')"
            End If
    
        Case "frmArticulos"
            If Not Trim(txtArticulos(0).Text) = "" And Not Trim(txtArticulos(1).Text) = "" Then
                vsql = vsql + " AND (Codigo >= '" & Trim(txtArticulos(0).Text) & "' AND Codigo <= '" & Trim(txtArticulos(1).Text) & "')"
            End If

            If Not Trim(txtArticulos(2).Text) = "" And Not Trim(txtArticulos(3).Text) = "" Then
                vsql = vsql + " AND (Descrip >= '" & Trim(txtArticulos(2).Text) & "%' AND Descrip  <= '" & Trim(txtArticulos(3).Text) & "%')"
            End If
            
            If Not Trim(txtArticulos(4).Text) = "" And Not Trim(txtArticulos(6).Text) = "" Then
                vsql = vsql + " AND (CProveedor >= '" & Trim(txtArticulos(4).Text) & "' AND CProveedor  <=  '" & Trim(txtArticulos(6).Text) & "')"
            End If

            If Not Trim(txtArticulos(8).Text) = "" And Not Trim(txtArticulos(9).Text) = "" Then
                vsql = vsql + " AND (idRubros >= '" & Trim(txtArticulos(8).Text) & "' AND idRubros  <=  '" & Trim(txtArticulos(8).Text) & "')"
            End If
            
            If Not Trim(txtArticulos(12).Text) = "" And Not Trim(txtArticulos(14).Text) = "" Then
                vsql = vsql + " AND (idSubRubros >= '" & Trim(txtArticulos(12).Text) & "' AND idSubRubros  <=  '" & Trim(txtArticulos(12).Text) & "')"
            End If
            If chkActivarFecha(2).Value = xtpChecked Then
                vsql = vsql + " AND (Fecha_Alta >= '" & strfechaMySQL(dtpAlta(2).Value) & "' AND Fecha_Alta <= '" & strfechaMySQL(dtpAlta(3).Value) & "')"
            End If
            If Not Trim(txtArticulos(16).Text) = "" And Not Trim(txtArticulos(18).Text) = "" Then
                vsql = vsql + " AND (idPorcentajeIva >= '" & Trim(txtArticulos(16).Text) & "' AND idPorcentajeIva  <=  '" & Trim(txtArticulos(18).Text) & "')"
            End If
    
        Case "frmMovimientosCuentas"
            If Not Trim(txtContabilidad(0).Text) = "" And Not Trim(txtContabilidad(2).Text) = "" Then
                vsql = vsql + " AND (CodigoCuenta >= '" & Trim(txtContabilidad(0).Text) & "' AND CodigoCuenta <= '" & Trim(txtContabilidad(2).Text) & "')"
            End If

            If chkActivarFecha(3).Value = xtpChecked Then
                'vSQL = vSQL + " AND (Fecha >= '" & strfechaMySQL(dtpContabilidad(0).Value) & "' AND Fecha <= '" & strfechaMySQL(dtpContabilidad(1).Value) & "')"
            End If
                   
        Case "frmCuentas"
        
    End Select
    
If Err Then GrabarLog "ArmaSQL", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub CargarGrilla()
On Error Resume Next
    
    Select Case vVieneImpresion
        
        Case "frmClientes"
            sqlFiltro = "SELECT * FROM VistaClientes WHERE 1=1 " & vsql & " ORDER BY " & VerOrden
            'sqlFiltro = "CALL SP_Clientes('" & txtIntervalos(0).Text & "','" & txtIntervalos(1).Text & "','" & txtIntervalos(2).Text & "','" & txtIntervalos(3).Text & "','" & txtIntervalos(4).Text & "','" & txtIntervalos(5).Text & "', '" & txtIntervalos(6).Text & "','" & txtIntervalos(8).Text & "','" & txtIntervalos(10).Text & "','" & strfechaMySQL(dtpAlta(0).Value) & "','" & strfechaMySQL(dtpAlta(1).Value) & "','" & strfechaMySQL(dtpNacimiento(0).Value) & "','" & strfechaMySQL(dtpNacimiento(1).Value) & "','" & txtIntervalos(12).Text & "','" & txtIntervalos(14).Text & "','" & txtIntervalos(16).Text & "','" & txtIntervalos(18).Text & "','" & txtIntervalos(20).Text & "','" & txtIntervalos(22).Text & "'," & VerOrden & ")"

        Case "frmProveedores"
            sqlFiltro = "SELECT * FROM VistaProveedores WHERE 1=1 " & vsql & " ORDER BY " & VerOrden
    
        Case "frmArticulos"
            sqlFiltro = "SELECT * FROM VistaArticulos WHERE 1=1 " & vsql & " ORDER BY " & Trim(vorden.Text)
    
        Case "frmMovimientosCuentas"
            sqlFiltro = "SELECT * FROM Cuentas WHERE 1=1 " & vsql & " ORDER BY " & VerOrden
        
        Case "frmCheques"
            sqlFiltro = "SELECT * FROM VistaCheques WHERE 1=1 " & vsql & ""
    
        Case "frmCuentas"
            sqlFiltro = "SELECT * FROM Cuentas WHERE 1=1 " & vsql & ""
    
    End Select

    With rsFiltro
        If .State = 1 Then .Close
        Call .Open(sqlFiltro, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If Not .EOF = True Then .MoveFirst
        
        Set dgFiltro.DataSource = rsFiltro
    
        dgFiltro.Visible = True
    End With

If Err Then GrabarLog "CargarGrilla", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Function VerOrden()
On Error Resume Next

    Dim vTipoOrden As String, i As Integer

    For i = 0 To rbOrden.Count - 1
        If rbOrden(i).Value = True Then
            vTipoOrden = rbOrden(i).Tag
            Exit For
        End If
    Next
    
    If chkTipoOrden.Value = xtpChecked Then
        vTipoOrden = vTipoOrden & " DESC"
    Else
        vTipoOrden = vTipoOrden & " ASC"
    End If
    
    VerOrden = vTipoOrden
    
If Err Then GrabarLog "VerOrden", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub SeleccionarModelo()
On Error Resume Next
    
    Dim i As Integer

    For i = 0 To Val(TabImpresion.ItemCount - 1)
        TabImpresion.Item(i).Visible = False
    Next
    
   ' TabImpresion.Visible = False
    
    Select Case vVieneImpresion
    
        Case "frmClientes", "frmProveedores"
            TabImpresion.Item(1).Visible = True
            TabImpresion.Item(4).Visible = True
            rbClasificacion(1).Caption = "Por Rubros"
            rbClasificacion(2).Caption = "Por Actividad"
            rbClasificacion(3).Caption = "Por Tipo de Cliente"
            
            With rbOrden
                .Item(0).Caption = "Codigo"
                .Item(0).Tag = "Codigo"
                .Item(1).Caption = "Nombre"
                .Item(1).Tag = "Nombre"
                .Item(2).Caption = "Nombre Comercial"
                .Item(2).Tag = "RazonSocial"
                .Item(3).Caption = "Localidad"
                .Item(3).Tag = "Localidad"
                .Item(4).Caption = "Direccion"
                .Item(4).Tag = "Direccion"
                .Item(5).Caption = "Telefono"
                .Item(5).Tag = "Telefono"
            End With
            
        Case "frmArticulos"
            TabImpresion.Item(0).Visible = True
            TabImpresion.Item(4).Visible = True
            rbClasificacion(1).Caption = "Por Proveedor"
            rbClasificacion(2).Caption = "Por Rubros"
            rbClasificacion(3).Caption = "Por Sub-Rubros"

            With rbOrden
                .Item(0).Caption = "Codigo"
                .Item(0).Tag = "Codigo"
                .Item(1).Caption = "Descripcion"
                .Item(1).Tag = "Descrip"
                .Item(2).Caption = "Proveedor"
                .Item(2).Tag = "CProveedor"
                .Item(3).Caption = "Rubro"
                .Item(3).Tag = "idRubros"
                .Item(4).Caption = "SubRubros"
                .Item(4).Tag = "idSubRubros"
                .Item(5).Caption = "P. Venta"
                .Item(5).Tag = "PVenta1"
            End With

            
        Case "frmEmpleados"
        Case "frmBuscarFactura"
        Case "frmCuentas"
            TabImpresion.Visible = True
            TabImpresion.Item(2).Visible = True
            TabImpresion.SelectedItem = 4
            
            
        Case "frmCheques"
            TabImpresion.Visible = True
            TabImpresion.Item(2).Visible = True
            TabImpresion.SelectedItem = 2
        
            
        Case "frmMovimientosCuentas"
            TabImpresion.Item(3).Visible = True
            TabImpresion.Item(4).Visible = True
    
            With rbOrden
                .Item(0).Caption = "Codigo"
                .Item(0).Tag = "CodigoCuenta"
                .Item(1).Caption = "Cuenta"
                .Item(1).Tag = "Cuenta"
                .Item(2).Caption = ""
                .Item(2).Tag = ""
                .Item(3).Caption = ""
                .Item(3).Tag = ""
                .Item(4).Caption = ""
                .Item(4).Tag = ""
                .Item(5).Caption = ""
                .Item(5).Tag = ""
            End With
    
    End Select

If Err Then GrabarLog "SeleccionarModelo", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub pbContabilidad_Click(Index As Integer)
On Error Resume Next

    vVuelveBusqueda = Me.Name
    vVieneBusqueda = pbContabilidad(Index).Tag
    
    frmBusqueda.Show

If Err Then GrabarLog "pbContabilidad_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub TabImpresion_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error Resume Next

    Select Case Item.Index
    
        Case 0
        
        Case 1
        
        Case 2
        
        Case 3
        
        Case 4
        
        Case 5
        
        Case 6
    
    End Select
    
If Err Then GrabarLog "TabImpresion_SelectedChanged", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtArticulos_Change(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 4, 6, 8, 10, 12, 14, 16, 18
            txtArticulos(Index + 1).Text = ""

    End Select

If Err Then GrabarLog "txtArticulos_Change", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtBancoCajaDetalle_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
    
    If KeyAscii = 13 Then
        Select Case Index
    
    
            Case 0
                txtBancoCajaDetalle(Index + 1).Text = TraerDato("Bancos", "idBancos = '" & Trim(txtBancoCajaDetalle(Index).Text) & "'", "Descripcion")
                txtBancoCajaDetalle(Index + 2).SetFocus
            
            Case 2
                txtBancoCajaDetalle(Index + 1).Text = TraerDato("Bancos", "idBancos = '" & Trim(txtBancoCajaDetalle(Index).Text) & "'", "Descripcion")
                txtBancoCajaDetalle(Index + 2).SetFocus
        
        End Select
    
    End If
    
If Err Then GrabarLog "txtBancoCajaDetalle_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtIntervalos_Change(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 6, 8, 10, 12, 14, 16, 18, 20, 22
            txtIntervalos(Index + 1).Text = ""

    End Select

If Err Then GrabarLog "txtIntervalos_Change", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Imprimir(vDirectoAImpresora As Boolean)
On Error Resume Next

    Dim i As Integer
    
   
' seteo de consulta en el DE

     Select Case vVieneImpresion
        Case "frmClientes", "frmProveedores"
            Mantenimiento.rsReporte.Source = rsFiltro.Source
            
        Case "frmArticulos"
            For i = 0 To Val(rbClasificacion.Count - 1)
                
                If rbClasificacion(i).Value = True Then
                    Select Case i
                    
                        Case 0
                            With Mantenimiento.rsReporte
                                If .State = 1 Then .Close
                                .Source = rsFiltro.Source
            
                                If .State = 0 Then .Open
                                .Close
                                .Open
                                
           
                            End With
                            
                        Case 1
                            With Mantenimiento.rsReporteShape
                                If .State = 1 Then .Close
                                .Source = "SHAPE {SELECT * FROM Proveedores} AS ReporteShape APPEND ({" & rsFiltro.Source & "}  AS ReporteShapeHijo RELATE 'idProveedores' TO 'CProveedor') AS ReporteShapeHijo"
                                If .State = 0 Then .Open
                                .Close
                                .Open
                            End With
                        Case 2
                            With Mantenimiento.rsReporteShape
                                If .State = 1 Then .Close
                                .Source = " SHAPE {SELECT * FROM Rubros} AS ReporteShape APPEND ({" & rsFiltro.Source & "}  AS ReporteShapeHijo RELATE 'idRubros' TO 'idRubros') AS ReporteShapeHijo"
                                If .State = 0 Then .Open
                                .Close
                                .Open
                            End With
                        Case 3
                            With Mantenimiento.rsReporteShape
                                If .State = 1 Then .Close
                                .Source = " SHAPE {SELECT * FROM SubRubros} AS ReporteShape APPEND ({" & rsFiltro.Source & "}  AS ReporteShapeHijo RELATE 'idSubRubros' TO 'idSubRubros') AS ReporteShapeHijo"
                                If .State = 0 Then .Open
                                .Close
                                .Open
                            End With
                    End Select
                    
                End If
            Next
            
        Case "frmCheques"
            With Mantenimiento.rsValoresDiferidosEncabezado
                If .State = 1 Then .Close
                
                '.Source = rsFiltro.Source
                .Source = " SHAPE {SELECT B.idBancos, B.Descripcion, BC.idBancosCuentas, BC.Cuenta, BC.Descripcion, BC.CuentaContableAsociada, BC.idTipoCuentaBanco FROM Bancos B INNER JOIN BancosCuentas BC ON B.idBancos = BC.idBancos " & vSQLEncabezado & ";}  AS ValoresDiferidosEncabezado APPEND ({SELECT * FROM VistaCheques  " & vSQLDetalle & ";}  AS ValoresDiferidosDetalle RELATE 'idBancos' TO 'idBancos','idBancosCuentas' TO 'idBancosCuentas') AS ValoresDiferidosDetalle"
                
                If .State = 0 Then .Open
                .Close
                .Open
            End With

    End Select






                               ' Unload Mantenimiento
                                'Load Mantenimiento
                                
                                'Set drReporte.DataSource = Mantenimiento.rsReporte
                                
                                




' panic:Modificación. Esto hay que pasarlo



    
    Select Case vVieneImpresion
    
        Case "frmClientes"
        
                frmListadoSaldos.instanciaCP = "Clientes"
                frmListadoSaldos.RBTipoSaldos(0).Value = True
                
                frmListadoSaldos.Show
            
        
        
        
        Case "frmProveedores"
            
                frmListadoSaldos.instanciaCP = "Proveedores"
                frmListadoSaldos.RBTipoSaldos(0).Value = True
                
                frmListadoSaldos.Show
                
                
                Exit Sub
            
            
            With drReporte
                '.Sections("EncabezadoInforme")
                .Sections("TituloEmpresa").Controls("lblTitulo").Caption = "Listado de Clientes"
                '.Sections ("DetalleInforme")
                '.Sections ("PiePagina")
                '.Sections ("PieInforme")
                
                .TopMargin = 250
                .BottomMargin = 0
                .LeftMargin = 500
                .RightMargin = 250
                If vDirectoAImpresora = True Then
                    .Hide
                    Call .PrintReport(False, rptRangeAllPages)
                Else
                    Call .Show
                End If
            End With
    
        Case "frmArticulos"
            
            For i = 0 To Val(rbClasificacion.Count - 1)
                
                If rbClasificacion(i).Value = True Then
                
                    Select Case i
            
                        Case 0
                            With drReporte
                            
                            
                                .Sections("DetalleInforme").Controls("txtCampo01").DataField = rsFiltro.Fields("codigo").Name
                                
                                .Sections("DetalleInforme").Controls("txtCampo02").DataField = "Descrip"
                                .Sections("TituloEmpresa").Controls("lblCampo02").Caption = "Descripcion"
                    
                                .Sections("DetalleInforme").Controls("txtCampo03").DataField = "Rubro"
                                .Sections("TituloEmpresa").Controls("lblCampo03").Caption = "Rubro"
                                
                    
                                .Sections("DetalleInforme").Controls("txtCampo04").DataField = "Proveedor"
                                .Sections("TituloEmpresa").Controls("lblCampo04").Caption = "Proveedor"
                                
                                .Sections("DetalleInforme").Controls("txtCampo05").DataField = "Stock"
                                .Sections("DetalleInforme").Controls("txtCampo05").Alignment = 1
                                .Sections("TituloEmpresa").Controls("lblCampo05").Caption = "Stock"
                                
                                
                                If chkArticulosPrecioConIva.Value = xtpUnchecked Then
                                    .Sections("DetalleInforme").Controls("txtCampo06").DataField = Trim((Me.vpventa))
                                    .Sections("DetalleInforme").Controls("txtCampo06").DataFormat = "Moneda"
                                Else
                                    .Sections("DetalleInforme").Controls("txtCampo06").DataField = "PVentaConIva"
                                    '.Sections("DetalleInforme").Controls("txtCampo06").DataField = Trim(Val(Me.vpventa))
                                    '.Sections("DetalleInforme").Controls("txtCampo06").DataFormat = "Moneda"
                                    
                                End If
                                
                                
                                '.Sections("DetalleInforme").Controls("txtCampo07").DataField = "Valorizacion"
                                .Sections("DetalleInforme").Controls("txtCampo07").Alignment = 1
                                
                                .Sections("TituloEmpresa").Controls("lblCampo07").Caption = "Valorizacion"

                                 .Sections("DetalleInforme").Controls("txtCampo07").DataField = "Valoralizacion"
                                
                                .Sections("DetalleInforme").Controls("total07").DataFormat = "Moneda"

                                .Sections("PieInforme").Controls("total07").DataField = "Valoralizacion"

                                .Sections("DetalleInforme").Controls("txtCampo08").DataField = "pcosto"
                                .Sections("DetalleInforme").Controls("txtCampo08").Alignment = 1
                                .Sections("TituloEmpresa").Controls("lblCampo08").Caption = "P.Costo"

                                .Sections("TituloEmpresa").Controls("lblCampo06").Caption = "P. Venta"
                                .Sections("TituloEmpresa").Controls("lblTitulo").Caption = "Listado de Articulos"

                                '.Sections("EncabezadoInforme")
                                '.Sections ("DetalleInforme")
                                '.Sections ("PiePagina")
                                '.Sections ("PieInforme")
                                .TopMargin = 250
                                .BottomMargin = 0
                                .LeftMargin = 500
                                .RightMargin = 250
                                
                            If LeerXml("MostrarSaldoEnDoc") = "SI" Then
                            
                                    .Sections("DetalleInforme").Controls("txtCampo03").Visible = False
                                    .Sections("TituloEmpresa").Controls("lblCampo03").Visible = False
                                    
                                    .Sections("DetalleInforme").Controls("txtCampo04").Visible = False
                                    .Sections("TituloEmpresa").Controls("lblCampo04").Visible = False
                                    
                                   ' .Sections("DetalleInforme").Controls("txtCampo05").Visible = False
                                   ' .Sections("DetalleInforme").Controls("lblCampo05").Visible = False
                                    
                                    .Sections("DetalleInforme").Controls("txtCampo07").Visible = False
                                    .Sections("TituloEmpresa").Controls("lblCampo07").Visible = False
                                    
                                    .Sections("DetalleInforme").Controls("txtCampo08").Visible = False
                                    .Sections("TituloEmpresa").Controls("lblCampo08").Visible = False
                                    
                            End If
                             
                            
                            
                            
                            If Me.rbClientes Then
                            
                            
                               .Sections("DetalleInforme").Controls("txtCampo03").DataField = "Rubro"
                                .Sections("TituloEmpresa").Controls("lblCampo03").Caption = "Rubro"
                    
                                .Sections("DetalleInforme").Controls("txtCampo04").DataField = "SubRubro"
                                .Sections("TituloEmpresa").Controls("lblCampo04").Caption = "Sub Rubro"
                                
                                .Sections("DetalleInforme").Controls("txtCampo05").DataField = "pventa5"
                                .Sections("DetalleInforme").Controls("txtCampo05").Alignment = 1
                                .Sections("TituloEmpresa").Controls("lblCampo05").Caption = ""

                                .Sections("DetalleInforme").Controls("txtCampo07").DataField = "pventa5"
                                .Sections("DetalleInforme").Controls("txtCampo07").Alignment = 1
                                .Sections("TituloEmpresa").Controls("lblCampo07").Caption = ""
                                
                                .Sections("PieInforme").Controls("total07").DataField = ""
                                
                                
                                
                                
    
                                
                                If chkArticulosPrecioConIva.Value = xtpUnchecked Then
                                    .Sections("DetalleInforme").Controls("txtCampo08").DataField = Trim((Me.vpventa))
                                    .Sections("DetalleInforme").Controls("txtCampo08").DataFormat = "Moneda"
                                    .Sections("DetalleInforme").Controls("txtCampo08").Visible = False
                                    .Sections("DetalleInforme").Controls("txtCampo07").Visible = False


                                Else
                                    .Sections("DetalleInforme").Controls("txtCampo08").DataField = "PVentaConIva"
                                    .Sections("DetalleInforme").Controls("txtCampo08").Alignment = 1
                                    .Sections("TituloEmpresa").Controls("lblCampo08").Caption = "C/IVA"
                               
                                End If
                                
                                
                                
                                
                                
                                
                                
                                 .Sections("DetalleInforme").Controls("txtCampo06").DataField = "PVenta1"
                                .Sections("DetalleInforme").Controls("txtCampo06").Alignment = 1
                                .Sections("TituloEmpresa").Controls("lblCampo06").Caption = "Público"
                        
                                .Sections("DetalleInforme").Controls("txtCampo02").DataField = "Valorizacion"
                                .Sections("TituloEmpresa").Controls("lblCampo02").Caption = "Valorizacion"
                    
                                .Sections("DetalleInforme").Controls("txtCampo03").DataField = "pventa5"
                                .Sections("TituloEmpresa").Controls("lblCampo03").Caption = ""
                    
                                .Sections("DetalleInforme").Controls("txtCampo04").DataField = "pventa5"
                                .Sections("TituloEmpresa").Controls("lblCampo04").Caption = ""
                                
                               .Sections("DetalleInforme").Controls("txtCampo05").DataField = "pventa5"
                                .Sections("DetalleInforme").Controls("txtCampo05").Alignment = 1
                                .Sections("TituloEmpresa").Controls("lblCampo05").Caption = ""
                                
                                .Sections("DetalleInforme").Controls("txtCampo02").DataField = "Descrip"
                                .Sections("TituloEmpresa").Controls("lblCampo02").Caption = "Descripcion"
                             
                               
                                
                               .Sections("PieInforme").Controls("total07").DataField = "pcosto"
                                
                                 
                               .Sections("DetalleInforme").Controls("txtCampo03").DataField = "Rubro"
                                .Sections("TituloEmpresa").Controls("lblCampo03").Caption = "Rubro"
                    
                                .Sections("DetalleInforme").Controls("txtCampo05").DataField = "SubRubro"
                                .Sections("TituloEmpresa").Controls("lblCampo05").Caption = "Sub Rubro"
                                
                                
                            
                                                        
                            End If
                                
                                

                                
                                If vDirectoAImpresora = True Then
                                    .Hide
                                    Call .PrintReport(False, rptRangeAllPages)
                                Else
                                    Call .Show
                                End If
                                
                                
                                
                            End With
                    
                        Case 1
                            With drReporteShape
                                .Sections("EncabezadoGrupo").Controls("txtCampoShape01").DataField = "idProveedores"
                                .Sections("EncabezadoGrupo").Controls("txtCampoShape02").DataField = "Nombre"
                                
                                .Sections("DetalleInforme").Controls("txtCampo01").DataField = "Codigo"
                                .Sections("DetalleInforme").Controls("txtCampo02").DataField = "Descrip"
                                .Sections("EncabezadoPagina").Controls("lblCampo02").Caption = "Descripcion"
                    
                                .Sections("DetalleInforme").Controls("txtCampo03").DataField = "Rubro"
                                .Sections("EncabezadoPagina").Controls("lblCampo03").Caption = "Rubro"
                    
                                .Sections("DetalleInforme").Controls("txtCampo04").DataField = "Proveedor"
                                .Sections("EncabezadoGrupo").Controls("lblCampo04").Caption = "CProveedor"
                                
                                .Sections("DetalleInforme").Controls("txtCampo05").DataField = "Fabricante"
                                .Sections("EncabezadoGrupo").Controls("lblCampo05").Caption = "idFabricante"
                                
                                If chkArticulosPrecioConIva.Value = xtpUnchecked Then
                                    .Sections("DetalleInforme").Controls("txtCampo06").DataField = "PVenta1"
                                Else
                                    .Sections("DetalleInforme").Controls("txtCampo06").DataField = "PVentaConIva"
                                End If
                                
                                .Sections("EncabezadoGrupo").Controls("lblCampo06").Caption = "P. Venta"
                                .Sections("EncabezadoPagina").Controls("lblTitulo").Caption = "Listado de Articulos"
                                
                                .TopMargin = 250
                                .BottomMargin = 0
                                .LeftMargin = 500
                                .RightMargin = 250
                                If vDirectoAImpresora = True Then
                                    .Hide
                                    Call .PrintReport(False, rptRangeAllPages)
                                Else
                                    Call .Show
                                End If
                            End With
                        Case 2
                
                        Case 3
            
            
                    End Select
                End If
            Next
            
            
        Case "frmEmpleados"
        Case "frmBuscarFactura"
        Case "frmBuscarCompra"
        Case "frmCuentas"
        
        
        Case "frmCheques"
            With drCheques
                If chkDiferido.Value = xtpUnchecked Then
                    .Sections("TituloEmpresa").Controls("lblTitulo").Caption = "Listado de Cheques"
                Else
                    .Sections("TituloEmpresa").Controls("lblTitulo").Caption = "Listado de Valores Diferidos"
                End If
                If vDirectoAImpresora = True Then
                    .Hide
                    Call .PrintReport(False, rptRangeAllPages)
                Else
                    Call .Refresh
                    Call .Show
                End If
            End With
    
    End Select

If Err Then GrabarLog "Imprimir", Err.Number & " " & Err.Description, Me.Caption
End Sub

