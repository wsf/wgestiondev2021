VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{52463EDA-D668-43B6-8D47-4FA8035EF04A}#1.0#0"; "PhotoWSF.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "Copia de KlexGrid.ocx"
Begin VB.Form frmArticulosAlta 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   9390
   Begin XtremeSuiteControls.FlatEdit pventa1 
      Height          =   330
      Left            =   7950
      TabIndex        =   97
      Top             =   3420
      Width           =   1185
      _Version        =   851968
      _ExtentX        =   2090
      _ExtentY        =   582
      _StockProps     =   77
      BackColor       =   14737632
      Text            =   "40"
      BackColor       =   14737632
   End
   Begin XtremeSuiteControls.PushButton pbCarga 
      Height          =   315
      Index           =   0
      Left            =   9000
      TabIndex        =   14
      Top             =   120
      Width           =   315
      _Version        =   851968
      _ExtentX        =   556
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "..."
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton pbCierraFoto 
      Height          =   315
      Left            =   9000
      TabIndex        =   15
      Top             =   420
      Visible         =   0   'False
      Width           =   315
      _Version        =   851968
      _ExtentX        =   556
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "X"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtAlta 
      Height          =   315
      Index           =   1
      Left            =   2760
      TabIndex        =   1
      Top             =   390
      Width           =   4215
      _Version        =   851968
      _ExtentX        =   7435
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.TabControl TabAlta 
      Height          =   5355
      Left            =   120
      TabIndex        =   18
      Top             =   1890
      Width           =   9135
      _Version        =   851968
      _ExtentX        =   16113
      _ExtentY        =   9446
      _StockProps     =   68
      Color           =   8
      ItemCount       =   5
      Item(0).Caption =   "Ficha"
      Item(0).ControlCount=   44
      Item(0).Control(0)=   "lblFicha(1)"
      Item(0).Control(1)=   "lblFicha(2)"
      Item(0).Control(2)=   "txtFicha(2)"
      Item(0).Control(3)=   "txtFicha(3)"
      Item(0).Control(4)=   "lblFicha(3)"
      Item(0).Control(5)=   "txtFicha(4)"
      Item(0).Control(6)=   "lblFicha(5)"
      Item(0).Control(7)=   "txtFicha(6)"
      Item(0).Control(8)=   "pbCarga(3)"
      Item(0).Control(9)=   "txtFicha(1)"
      Item(0).Control(10)=   "txtFicha(0)"
      Item(0).Control(11)=   "txtFicha(5)"
      Item(0).Control(12)=   "pbCarga(4)"
      Item(0).Control(13)=   "lblFicha(0)"
      Item(0).Control(14)=   "txtFicha(7)"
      Item(0).Control(15)=   "klexPrecios"
      Item(0).Control(16)=   "lblFicha(4)"
      Item(0).Control(17)=   "txtFicha(8)"
      Item(0).Control(18)=   "pbCarga(5)"
      Item(0).Control(19)=   "txtFicha(9)"
      Item(0).Control(20)=   "lblFicha(6)"
      Item(0).Control(21)=   "txtFicha(10)"
      Item(0).Control(22)=   "lblFicha(7)"
      Item(0).Control(23)=   "txtFicha(11)"
      Item(0).Control(24)=   "lblFicha(8)"
      Item(0).Control(25)=   "txtFicha(12)"
      Item(0).Control(26)=   "lblFicha(9)"
      Item(0).Control(27)=   "cpcosto"
      Item(0).Control(28)=   "pventa2"
      Item(0).Control(29)=   "pventa3"
      Item(0).Control(30)=   "pventa4"
      Item(0).Control(31)=   "pventa5"
      Item(0).Control(32)=   "a1"
      Item(0).Control(33)=   "a2"
      Item(0).Control(34)=   "a3"
      Item(0).Control(35)=   "a4"
      Item(0).Control(36)=   "a5"
      Item(0).Control(37)=   "a7"
      Item(0).Control(38)=   "vstocka"
      Item(0).Control(39)=   "lblStock2"
      Item(0).Control(40)=   "lblAumentarStock"
      Item(0).Control(41)=   "lblStockActual"
      Item(0).Control(42)=   "PusCambiarStock"
      Item(0).Control(43)=   "lblAumento"
      Item(1).Caption =   "Técnica"
      Item(1).ControlCount=   21
      Item(1).Control(0)=   "dtpFecha(0)"
      Item(1).Control(1)=   "dtpFecha(1)"
      Item(1).Control(2)=   "txtTecnica(2)"
      Item(1).Control(3)=   "txtTecnica(3)"
      Item(1).Control(4)=   "txtTecnica(4)"
      Item(1).Control(5)=   "txtTecnica(0)"
      Item(1).Control(6)=   "txtTecnica(1)"
      Item(1).Control(7)=   "lblTecnica(7)"
      Item(1).Control(8)=   "lblTecnica(3)"
      Item(1).Control(9)=   "lblTecnica(4)"
      Item(1).Control(10)=   "lblTecnica(6)"
      Item(1).Control(11)=   "lblTecnica(5)"
      Item(1).Control(12)=   "lblTecnica(2)"
      Item(1).Control(13)=   "lblTecnica(1)"
      Item(1).Control(14)=   "lblTecnica(0)"
      Item(1).Control(15)=   "txtTecnica(7)"
      Item(1).Control(16)=   "txtTecnica(5)"
      Item(1).Control(17)=   "lblTecnica(8)"
      Item(1).Control(18)=   "txtTecnica(6)"
      Item(1).Control(19)=   "pbCarga(7)"
      Item(1).Control(20)=   "chkActualizacionDePrecio"
      Item(2).Caption =   "Stock"
      Item(2).ControlCount=   12
      Item(2).Control(0)=   "txtStock(2)"
      Item(2).Control(1)=   "txtStock(0)"
      Item(2).Control(2)=   "txtStock(1)"
      Item(2).Control(3)=   "lblStock(0)"
      Item(2).Control(4)=   "lblStock(3)"
      Item(2).Control(5)=   "lblStock(2)"
      Item(2).Control(6)=   "lblStock(1)"
      Item(2).Control(7)=   "cboDepositos"
      Item(2).Control(8)=   "KlexStock"
      Item(2).Control(9)=   "PusVerMovimientos(2)"
      Item(2).Control(10)=   "PushButton1"
      Item(2).Control(11)=   "PushButton2"
      Item(3).Caption =   "Cond. Especiales de Venta"
      Item(3).ControlCount=   4
      Item(3).Control(0)=   "KlexArticulosClientes"
      Item(3).Control(1)=   "cmdCondicionesEspeciales(0)"
      Item(3).Control(2)=   "cmdCondicionesEspeciales(1)"
      Item(3).Control(3)=   "cmdCondicionesEspeciales(2)"
      Item(4).Caption =   "Proveedores"
      Item(4).ControlCount=   11
      Item(4).Control(0)=   "cmdProveedoresPrecio(0)"
      Item(4).Control(1)=   "cmdProveedoresPrecio(1)"
      Item(4).Control(2)=   "cmdProveedoresPrecio(2)"
      Item(4).Control(3)=   "txtProveedores(0)"
      Item(4).Control(4)=   "txtProveedores(1)"
      Item(4).Control(5)=   "txtProveedores(2)"
      Item(4).Control(6)=   "cmdProveedoresPrecio(3)"
      Item(4).Control(7)=   "lblProveedores(0)"
      Item(4).Control(8)=   "lblProveedores(1)"
      Item(4).Control(9)=   "KlexProveedores"
      Item(4).Control(10)=   "pbCarga(6)"
      Begin XtremeSuiteControls.PushButton PusCambiarStock 
         Height          =   285
         Left            =   1575
         TabIndex        =   112
         Top             =   2745
         Width           =   2265
         _Version        =   851968
         _ExtentX        =   3995
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Cambiar Stock"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit vstocka 
         Height          =   330
         Left            =   1575
         TabIndex        =   108
         Top             =   2385
         Width           =   2265
         _Version        =   851968
         _ExtentX        =   3995
         _ExtentY        =   582
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton a1 
         Height          =   285
         Left            =   3285
         TabIndex        =   102
         Top             =   1890
         Width           =   780
         _Version        =   851968
         _ExtentX        =   1376
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "+"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit cpcosto 
         Height          =   330
         Left            =   1935
         TabIndex        =   96
         Top             =   1845
         Width           =   1140
         _Version        =   851968
         _ExtentX        =   2011
         _ExtentY        =   582
         _StockProps     =   77
         BackColor       =   14737632
         BackColor       =   14737632
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   345
         Left            =   -63820
         TabIndex        =   84
         Top             =   2910
         Visible         =   0   'False
         Width           =   2865
         _Version        =   851968
         _ExtentX        =   5054
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Borrar Historial de movimiento de stock"
         Appearance      =   4
      End
      Begin XtremeSuiteControls.CheckBox chkActualizacionDePrecio 
         Height          =   255
         Left            =   -62560
         TabIndex        =   82
         ToolTipText     =   "Setea manualmente la fecha de la ultima actualizacion de Precio"
         Top             =   750
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cambiar"
         BackColor       =   -2147483639
         UseVisualStyle  =   -1  'True
      End
      Begin Grid.KlexGrid KlexArticulosClientes 
         Height          =   3135
         Left            =   -69880
         TabIndex        =   56
         Top             =   480
         Visible         =   0   'False
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   5530
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
         MouseIcon       =   "frmArticulosAlta.frx":0000
         Rows            =   10
      End
      Begin Grid.KlexGrid klexPrecios 
         Height          =   1815
         Left            =   120
         TabIndex        =   55
         Top             =   3420
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   3201
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
         MouseIcon       =   "frmArticulosAlta.frx":001C
         Rows            =   7
      End
      Begin XtremeSuiteControls.ComboBox cboDepositos 
         Height          =   315
         Left            =   -68170
         TabIndex        =   47
         Top             =   3390
         Visible         =   0   'False
         Width           =   2535
         _Version        =   851968
         _ExtentX        =   4471
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Index           =   0
         Left            =   -67600
         TabIndex        =   42
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57344001
         CurrentDate     =   40290
      End
      Begin XtremeSuiteControls.FlatEdit txtTecnica 
         Height          =   315
         Index           =   2
         Left            =   -67600
         TabIndex        =   10
         Top             =   1440
         Visible         =   0   'False
         Width           =   1575
         _Version        =   851968
         _ExtentX        =   2778
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   1
         Left            =   3360
         TabIndex        =   5
         Top             =   420
         Width           =   2175
         _Version        =   851968
         _ExtentX        =   3836
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   0
         Left            =   1920
         TabIndex        =   39
         Top             =   420
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   2
         Left            =   6720
         TabIndex        =   6
         Top             =   420
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   5
         Left            =   1920
         TabIndex        =   7
         Top             =   1140
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtTecnica 
         Height          =   315
         Index           =   3
         Left            =   -64720
         TabIndex        =   11
         Top             =   1440
         Visible         =   0   'False
         Width           =   3015
         _Version        =   851968
         _ExtentX        =   5318
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         MaxLength       =   3
      End
      Begin XtremeSuiteControls.FlatEdit txtTecnica 
         Height          =   315
         Index           =   4
         Left            =   -67600
         TabIndex        =   12
         Top             =   1800
         Visible         =   0   'False
         Width           =   5895
         _Version        =   851968
         _ExtentX        =   10398
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtTecnica 
         Height          =   1395
         Index           =   7
         Left            =   -67600
         TabIndex        =   13
         Top             =   2520
         Visible         =   0   'False
         Width           =   5895
         _Version        =   851968
         _ExtentX        =   10398
         _ExtentY        =   2461
         _StockProps     =   77
         BackColor       =   -2147483643
         MaxLength       =   255
      End
      Begin XtremeSuiteControls.FlatEdit txtTecnica 
         Height          =   315
         Index           =   0
         Left            =   -67600
         TabIndex        =   8
         Top             =   1080
         Visible         =   0   'False
         Width           =   2100
         _Version        =   851968
         _ExtentX        =   3704
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         MaxLength       =   3
      End
      Begin XtremeSuiteControls.FlatEdit txtTecnica 
         Height          =   315
         Index           =   1
         Left            =   -64045
         TabIndex        =   9
         Top             =   1080
         Visible         =   0   'False
         Width           =   2340
         _Version        =   851968
         _ExtentX        =   4128
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   3
         Left            =   1920
         TabIndex        =   30
         Top             =   780
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   4
         Left            =   3360
         TabIndex        =   31
         Top             =   780
         Width           =   4455
         _Version        =   851968
         _ExtentX        =   7858
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   3
         Left            =   2880
         TabIndex        =   36
         Tag             =   "PorcentajeIva"
         Top             =   420
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   4
         Left            =   2880
         TabIndex        =   37
         Tag             =   "Proveedor"
         Top             =   780
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   5
         Left            =   2880
         TabIndex        =   40
         Tag             =   "Fabricante"
         Top             =   1140
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   6
         Left            =   3360
         TabIndex        =   41
         Top             =   1140
         Width           =   4455
         _Version        =   851968
         _ExtentX        =   7858
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Index           =   1
         Left            =   -64045
         TabIndex        =   43
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   187695105
         CurrentDate     =   40290
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   7
         Left            =   1920
         TabIndex        =   2
         Top             =   1500
         Width           =   2175
         _Version        =   851968
         _ExtentX        =   3836
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtStock 
         Height          =   315
         Index           =   2
         Left            =   -63730
         TabIndex        =   49
         Top             =   3390
         Visible         =   0   'False
         Width           =   2505
         _Version        =   851968
         _ExtentX        =   4410
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtStock 
         Height          =   315
         Index           =   0
         Left            =   -68170
         TabIndex        =   48
         Top             =   3750
         Visible         =   0   'False
         Width           =   2505
         _Version        =   851968
         _ExtentX        =   4410
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtStock 
         Height          =   315
         Index           =   1
         Left            =   -63730
         TabIndex        =   50
         Top             =   3750
         Visible         =   0   'False
         Width           =   2505
         _Version        =   851968
         _ExtentX        =   4410
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtTecnica 
         Height          =   315
         Index           =   5
         Left            =   -67600
         TabIndex        =   57
         Top             =   2160
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
         Left            =   -66640
         TabIndex        =   58
         Top             =   2160
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtTecnica 
         Height          =   315
         Index           =   6
         Left            =   -66160
         TabIndex        =   60
         Top             =   2160
         Visible         =   0   'False
         Width           =   4455
         _Version        =   851968
         _ExtentX        =   7858
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin Grid.KlexGrid KlexProveedores 
         Height          =   2655
         Left            =   -69880
         TabIndex        =   62
         Top             =   480
         Visible         =   0   'False
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   4683
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
         MouseIcon       =   "frmArticulosAlta.frx":0038
         Rows            =   10
      End
      Begin XtremeSuiteControls.PushButton cmdProveedoresPrecio 
         Height          =   375
         Index           =   0
         Left            =   -69880
         TabIndex        =   63
         Top             =   3650
         Visible         =   0   'False
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "&Nuevo"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdProveedoresPrecio 
         Height          =   375
         Index           =   1
         Left            =   -68560
         TabIndex        =   64
         Top             =   3650
         Visible         =   0   'False
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "&Modificar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdProveedoresPrecio 
         Height          =   375
         Index           =   2
         Left            =   -67240
         TabIndex        =   65
         Top             =   3650
         Visible         =   0   'False
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "&Borrar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtProveedores 
         Height          =   315
         Index           =   0
         Left            =   -68440
         TabIndex        =   66
         Top             =   3240
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtProveedores 
         Height          =   315
         Index           =   1
         Left            =   -67000
         TabIndex        =   67
         Top             =   3240
         Visible         =   0   'False
         Width           =   3495
         _Version        =   851968
         _ExtentX        =   6165
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   6
         Left            =   -67480
         TabIndex        =   68
         Tag             =   "ProveedorPrecio"
         Top             =   3240
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtProveedores 
         Height          =   315
         Index           =   2
         Left            =   -62440
         TabIndex        =   69
         Top             =   3240
         Visible         =   0   'False
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdProveedoresPrecio 
         Height          =   375
         Index           =   3
         Left            =   -65920
         TabIndex        =   70
         Top             =   3650
         Visible         =   0   'False
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "&Guardar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdCondicionesEspeciales 
         Height          =   375
         Index           =   0
         Left            =   -69880
         TabIndex        =   73
         Top             =   3650
         Visible         =   0   'False
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "&Nuevo"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdCondicionesEspeciales 
         Height          =   375
         Index           =   1
         Left            =   -68560
         TabIndex        =   74
         Top             =   3650
         Visible         =   0   'False
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "&Modificar"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdCondicionesEspeciales 
         Height          =   375
         Index           =   2
         Left            =   -67240
         TabIndex        =   75
         Top             =   3650
         Visible         =   0   'False
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "&Borrar"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin Grid.KlexGrid KlexStock 
         Height          =   2415
         Left            =   -69880
         TabIndex        =   76
         Top             =   480
         Visible         =   0   'False
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   4260
         EnterKeyBehaviour=   0
         BackColorAlternate=   0
         GridLinesFixed  =   2
         BackColorFixed  =   -2147483626
         Cols            =   8
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
         MouseIcon       =   "frmArticulosAlta.frx":0054
         Rows            =   5
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   8
         Left            =   5640
         TabIndex        =   3
         Top             =   1500
         Width           =   1725
         _Version        =   851968
         _ExtentX        =   3043
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.PushButton PusVerMovimientos 
         Height          =   345
         Index           =   2
         Left            =   -69880
         TabIndex        =   83
         Top             =   2910
         Visible         =   0   'False
         Width           =   2805
         _Version        =   851968
         _ExtentX        =   4948
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Ver movimientos de stock"
         Enabled         =   0   'False
         Appearance      =   6
         Picture         =   "frmArticulosAlta.frx":0070
         BorderGap       =   10
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   345
         Left            =   -65890
         TabIndex        =   85
         Top             =   2910
         Visible         =   0   'False
         Width           =   2055
         _Version        =   851968
         _ExtentX        =   3625
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Borrar linea"
         Appearance      =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   9
         Left            =   5640
         TabIndex        =   88
         Top             =   1860
         Width           =   1725
         _Version        =   851968
         _ExtentX        =   3043
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   10
         Left            =   5640
         TabIndex        =   90
         Top             =   2220
         Width           =   1725
         _Version        =   851968
         _ExtentX        =   3043
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   11
         Left            =   5640
         TabIndex        =   92
         Top             =   2580
         Width           =   1725
         _Version        =   851968
         _ExtentX        =   3043
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   12
         Left            =   5640
         TabIndex        =   94
         Top             =   2940
         Width           =   1725
         _Version        =   851968
         _ExtentX        =   3043
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit pventa2 
         Height          =   330
         Left            =   7830
         TabIndex        =   98
         Top             =   1890
         Width           =   1185
         _Version        =   851968
         _ExtentX        =   2090
         _ExtentY        =   582
         _StockProps     =   77
         BackColor       =   14737632
         Text            =   "55"
         BackColor       =   14737632
      End
      Begin XtremeSuiteControls.FlatEdit pventa3 
         Height          =   330
         Left            =   7830
         TabIndex        =   99
         Top             =   2250
         Width           =   1185
         _Version        =   851968
         _ExtentX        =   2090
         _ExtentY        =   582
         _StockProps     =   77
         BackColor       =   14737632
         BackColor       =   14737632
      End
      Begin XtremeSuiteControls.FlatEdit pventa4 
         Height          =   330
         Left            =   7830
         TabIndex        =   100
         Top             =   2610
         Width           =   1185
         _Version        =   851968
         _ExtentX        =   2090
         _ExtentY        =   582
         _StockProps     =   77
         BackColor       =   14737632
         BackColor       =   14737632
      End
      Begin XtremeSuiteControls.FlatEdit pventa5 
         Height          =   330
         Left            =   7830
         TabIndex        =   101
         Top             =   2970
         Width           =   1185
         _Version        =   851968
         _ExtentX        =   2090
         _ExtentY        =   582
         _StockProps     =   77
         BackColor       =   14737632
         BackColor       =   14737632
      End
      Begin XtremeSuiteControls.PushButton a2 
         Height          =   285
         Left            =   7380
         TabIndex        =   103
         Top             =   1530
         Width           =   420
         _Version        =   851968
         _ExtentX        =   741
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "+"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton a3 
         Height          =   285
         Left            =   7380
         TabIndex        =   104
         Top             =   1890
         Width           =   420
         _Version        =   851968
         _ExtentX        =   741
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "+"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton a4 
         Height          =   285
         Left            =   7380
         TabIndex        =   105
         Top             =   2250
         Width           =   420
         _Version        =   851968
         _ExtentX        =   741
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "+"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton a5 
         Height          =   285
         Left            =   7380
         TabIndex        =   106
         Top             =   2610
         Width           =   420
         _Version        =   851968
         _ExtentX        =   741
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "+"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton a7 
         Height          =   285
         Left            =   7380
         TabIndex        =   107
         Top             =   2970
         Width           =   420
         _Version        =   851968
         _ExtentX        =   741
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "+"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblAumento 
         Height          =   315
         Left            =   8040
         TabIndex        =   113
         Top             =   1170
         Width           =   885
         _Version        =   851968
         _ExtentX        =   1561
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "% Aumento"
      End
      Begin XtremeSuiteControls.Label lblStockActual 
         Height          =   240
         Left            =   315
         TabIndex        =   111
         Top             =   3105
         Width           =   1140
         _Version        =   851968
         _ExtentX        =   2011
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Stock Actual:"
      End
      Begin VB.Label lblAumentarStock 
         Caption         =   "Aumentar Stock:"
         Height          =   375
         Left            =   270
         TabIndex        =   110
         Top             =   2430
         Width           =   1320
      End
      Begin XtremeSuiteControls.Label lblStock2 
         Height          =   375
         Left            =   1575
         TabIndex        =   109
         Top             =   3015
         Width           =   2265
         _Version        =   851968
         _ExtentX        =   3995
         _ExtentY        =   661
         _StockProps     =   79
         BackColor       =   -2147483636
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio de Venta 5:"
         Height          =   195
         Index           =   9
         Left            =   4260
         TabIndex        =   95
         Top             =   2985
         Width           =   1755
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio de Venta 4:"
         Height          =   195
         Index           =   8
         Left            =   4260
         TabIndex        =   93
         Top             =   2625
         Width           =   1755
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio de Venta 3:"
         Height          =   195
         Index           =   7
         Left            =   4260
         TabIndex        =   91
         Top             =   2265
         Width           =   1755
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio de Venta 2:"
         Height          =   195
         Index           =   6
         Left            =   4260
         TabIndex        =   89
         Top             =   1905
         Width           =   1755
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio de Venta 1:"
         Height          =   195
         Index           =   5
         Left            =   4260
         TabIndex        =   77
         Top             =   1545
         Width           =   1755
      End
      Begin VB.Label lblProveedores 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio:"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   -63280
         TabIndex        =   72
         Top             =   3285
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblProveedores 
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor:"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   -69880
         TabIndex        =   71
         Top             =   3285
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblTecnica 
         BackStyle       =   0  'Transparent
         Caption         =   "Asociar Concepto Caja:"
         Height          =   195
         Index           =   7
         Left            =   -69520
         TabIndex        =   59
         Top             =   2190
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblStock 
         BackStyle       =   0  'Transparent
         Caption         =   "Depositos:"
         Height          =   195
         Index           =   0
         Left            =   -69850
         TabIndex        =   54
         Top             =   3420
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblStock 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Maximo :"
         Height          =   195
         Index           =   3
         Left            =   -65440
         TabIndex        =   53
         Top             =   3780
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblStock 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Minimo :"
         Height          =   195
         Index           =   2
         Left            =   -65440
         TabIndex        =   52
         Top             =   3420
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblStock 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Actual :"
         Height          =   195
         Index           =   1
         Left            =   -69850
         TabIndex        =   51
         Top             =   3780
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio de Costo:"
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   46
         Top             =   1545
         Width           =   1755
      End
      Begin VB.Label lblTecnica 
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones:"
         Height          =   195
         Index           =   8
         Left            =   -69520
         TabIndex        =   45
         Top             =   2570
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblTecnica 
         BackStyle       =   0  'Transparent
         Caption         =   "Peso Total :"
         Height          =   195
         Index           =   3
         Left            =   -65320
         TabIndex        =   44
         Top             =   1125
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Porcentaje :"
         Height          =   195
         Index           =   1
         Left            =   5640
         TabIndex        =   38
         Top             =   465
         Width           =   1095
      End
      Begin VB.Label lblTecnica 
         BackStyle       =   0  'Transparent
         Caption         =   "Unidades por Bulto :"
         Height          =   195
         Index           =   4
         Left            =   -69520
         TabIndex        =   29
         Top             =   1480
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblTecnica 
         BackStyle       =   0  'Transparent
         Caption         =   "Mensaje Emergente :"
         Height          =   195
         Index           =   6
         Left            =   -69520
         TabIndex        =   28
         Top             =   1840
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblTecnica 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dimensiones :"
         Height          =   195
         Index           =   5
         Left            =   -65920
         TabIndex        =   27
         Top             =   1480
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblTecnica 
         BackStyle       =   0  'Transparent
         Caption         =   "Peso por Unidad :"
         Height          =   195
         Index           =   2
         Left            =   -69520
         TabIndex        =   26
         Top             =   1120
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblTecnica 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Ult. Modificacion :"
         Height          =   195
         Index           =   1
         Left            =   -66040
         TabIndex        =   25
         Top             =   765
         Visible         =   0   'False
         Width           =   2715
      End
      Begin VB.Label lblTecnica 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha De Alta :"
         Height          =   195
         Index           =   0
         Left            =   -69520
         TabIndex        =   22
         Top             =   760
         Visible         =   0   'False
         Width           =   1750
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Fabricante:"
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   21
         Top             =   1185
         Width           =   1755
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor:"
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   20
         Top             =   825
         Width           =   1755
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "IVA:"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   19
         Top             =   465
         Width           =   1755
      End
   End
   Begin VB.PictureBox PicInferior 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      Picture         =   "frmArticulosAlta.frx":0477
      ScaleHeight     =   555
      ScaleWidth      =   9405
      TabIndex        =   23
      Top             =   7290
      Width           =   9400
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   0
         Left            =   6900
         TabIndex        =   4
         Top             =   120
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Grabar"
         Appearance      =   6
         Picture         =   "frmArticulosAlta.frx":552A
         BorderGap       =   10
      End
      Begin XtremeSuiteControls.PushButton PusVerMovimientos 
         Height          =   345
         Index           =   1
         Left            =   8040
         TabIndex        =   24
         Top             =   105
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Cerrar"
         Appearance      =   6
         Picture         =   "frmArticulosAlta.frx":5931
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtAlta 
      Height          =   315
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Top             =   60
      Width           =   4215
      _Version        =   851968
      _ExtentX        =   7435
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Alignment       =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtAlta 
      Height          =   315
      Index           =   2
      Left            =   2760
      TabIndex        =   32
      Top             =   750
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
      Left            =   3720
      TabIndex        =   33
      Tag             =   "SubRubro"
      Top             =   750
      Width           =   315
      _Version        =   851968
      _ExtentX        =   556
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "..."
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtAlta 
      Height          =   315
      Index           =   3
      Left            =   4200
      TabIndex        =   34
      Top             =   750
      Width           =   2775
      _Version        =   851968
      _ExtentX        =   4895
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
   End
   Begin PhotoWSF.Photo phtArticulo 
      Height          =   1305
      Left            =   7320
      TabIndex        =   61
      Top             =   30
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2302
      BorderColor     =   8421504
      BorderColor     =   8421504
      BackStyle       =   0
   End
   Begin XtremeSuiteControls.FlatEdit txtAlta 
      Height          =   315
      Index           =   4
      Left            =   2760
      TabIndex        =   78
      Top             =   1110
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
      Left            =   3720
      TabIndex        =   79
      Tag             =   "Rubro"
      Top             =   1110
      Width           =   315
      _Version        =   851968
      _ExtentX        =   556
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "..."
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtAlta 
      Height          =   315
      Index           =   5
      Left            =   4200
      TabIndex        =   80
      Top             =   1110
      Width           =   2775
      _Version        =   851968
      _ExtentX        =   4895
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit txtAlta 
      Height          =   315
      Index           =   6
      Left            =   2760
      TabIndex        =   86
      Top             =   1470
      Width           =   4215
      _Version        =   851968
      _ExtentX        =   7435
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Alignment       =   1
   End
   Begin VB.Label lblCódigoDe 
      BackStyle       =   0  'Transparent
      Caption         =   "Código de Barra:"
      Height          =   195
      Index           =   4
      Left            =   480
      TabIndex        =   87
      Top             =   1500
      Width           =   2250
   End
   Begin VB.Label lblAlta 
      BackStyle       =   0  'Transparent
      Caption         =   "Rubro / Familia :"
      Height          =   195
      Index           =   3
      Left            =   480
      TabIndex        =   81
      Top             =   1185
      Width           =   2250
   End
   Begin VB.Label lblAlta 
      BackStyle       =   0  'Transparent
      Caption         =   "SubRubro :"
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   35
      Top             =   870
      Width           =   2250
   End
   Begin VB.Label lblAlta 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion del Articulo :"
      Height          =   195
      Index           =   1
      Left            =   495
      TabIndex        =   17
      Top             =   510
      Width           =   2250
   End
   Begin VB.Label lblAlta 
      BackStyle       =   0  'Transparent
      Caption         =   "Código  del Articulo:"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   16
      Top             =   150
      Width           =   2250
   End
   Begin VB.Shape shpSuperior 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   1845
      Left            =   0
      Top             =   30
      Width           =   9420
   End
End
Attribute VB_Name = "frmArticulosAlta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vaccion, vViene As String
Dim vCostoAnterior As Double, vPVentaAnterior As Double

Private Sub a1_Click()
Me.txtFicha(7).Text = Val(Me.txtFicha(7).Text) + Val(cpcosto.Text)
cpcosto.Text = ""
End Sub

Private Sub a2_Click()
Me.txtFicha(8).Text = Val(Me.txtFicha(7).Text) * (1 + Val(pventa1.Text) / 100)
pventa1.Text = ""
End Sub

Private Sub a3_Click()
Me.txtFicha(9).Text = Val(Me.txtFicha(7).Text) * (1 + Val(pventa2.Text) / 100)
pventa2.Text = ""

End Sub

Private Sub a4_Click()
Me.txtFicha(10).Text = Val(Me.txtFicha(7).Text) * (1 + Val(pventa3.Text) / 100)
pventa3.Text = ""
End Sub

Private Sub a5_Click()
Me.txtFicha(11).Text = Val(Me.txtFicha(7).Text) * (1 + Val(pventa4.Text) / 100)
pventa4.Text = ""
End Sub

Private Sub a7_Click()
Me.txtFicha(12).Text = Val(Me.txtFicha(7).Text) * (1 + Val(pventa5.Text) / Val(Me.txtFicha(7).Text))
pventa5.Text = ""
End Sub

Private Sub cboDepositos_Click()
On Error Resume Next

    cboDepositos.Tag = TraerDato("Depositos", "Deposito = '" & Trim(cboDepositos.Text) & "'", "idDepositos")

If Err Then GrabarLog "cboDepositos_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cboDepositos_GotFocus()
On Error Resume Next

    Call CargarComboNew("Depositos", "Deposito", cboDepositos, True)

If Err Then GrabarLog "cboDepositos_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cmdCondicionesEspeciales_Click(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 0
            'NuevoEspecial
            
        Case 1
            'ModificarEspecial
        
        Case 2
            'BorrarRegistro
        
    
    End Select

If Err Then GrabarLog "cmdCondicionesEspeciales_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cmdProveedoresPrecio_Click(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 0
            NuevoProveedorPrecio (True)

        Case 1
            ModificarProveedorPrecio
        
        Case 2
            BorrarProveedorPrecio
            NuevoProveedorPrecio (True)
            CargarGrillaArticulosProveedores (txtAlta(0).Text)
        
        Case 3
            GuardarProveedorPrecio
            NuevoProveedorPrecio (False)
            CargarGrillaArticulosProveedores (txtAlta(0).Text)
            
    End Select

If Err Then GrabarLog "cmdProveedoresPrecio_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub CargarGrillaArticulosProveedores(vCodigoArticulo As String)
On Error Resume Next
    
    Dim rsArticulosProveedorPrecio As New ADODB.Recordset, sqlArticulosProveedorPrecio As String
    
    With KlexProveedores
        .Cols = 10
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 1
        
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 250
        
        .TextMatrix(0, 1) = "id"
        .ColWidth(1) = 0
        
        .TextMatrix(0, 2) = "Codigo"
        .ColWidth(2) = 1000
        
        .TextMatrix(0, 3) = "Nombre Proveedor"
        .ColWidth(3) = 3500
        
        .TextMatrix(0, 4) = "Localidad"
        .ColWidth(4) = 1500
        
        .TextMatrix(0, 5) = "Telefono"
        .ColWidth(5) = 1500
        
        .TextMatrix(0, 6) = "Precio"
        .ColWidth(6) = 1000
        .ColDisplayFormat(6) = "#0.00"
        
        Dim i As Integer
        
        For i = 7 To .Cols - 1
            .TextMatrix(0, i) = ""
            .ColWidth(i) = 0
        Next

        .Editable = True
       
        '.EnterKeyBehaviour = klexEKMoveDown
        .EnterKeyBehaviour = klexEKNone
        .BackColorAlternate = &HE0E0E0
    End With
    
    sqlArticulosProveedorPrecio = "SELECT idarticulosproveedorprecio as Id, CodigoArticulo, Codigo, Nombre, Localidad, Telefono, PrecioDeCosto FROM articulosproveedorprecio AP INNER JOIN Proveedores P ON AP.CodigoProveedor=P.Codigo WHERE (CodigoArticulo = '" & Trim(vCodigoArticulo) & "');"

    With rsArticulosProveedorPrecio
        Call .Open(sqlArticulosProveedorPrecio, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not .EOF = True Then
            .MoveFirst
            KlexProveedores.Rows = .RecordCount + 1
        Else
            For i = 0 To KlexProveedores.Cols - 1
                KlexProveedores.TextMatrix(1, i) = ""
            Next
        End If
        Do Until .EOF = True
        
            KlexProveedores.TextMatrix(.AbsolutePosition, 1) = EsNulo(.Fields("id").Value)
            KlexProveedores.TextMatrix(.AbsolutePosition, 2) = EsNulo(.Fields("Codigo").Value)
            KlexProveedores.TextMatrix(.AbsolutePosition, 3) = EsNulo(.Fields("Nombre").Value)
            KlexProveedores.TextMatrix(.AbsolutePosition, 4) = EsNulo(.Fields("Localidad").Value)
            KlexProveedores.TextMatrix(.AbsolutePosition, 5) = EsNulo(.Fields("Telefono").Value)
            KlexProveedores.TextMatrix(.AbsolutePosition, 6) = EsNulo(.Fields("PrecioDeCosto").Value)
            KlexProveedores.TextMatrix(.AbsolutePosition, 7) = EsNulo(.Fields("CodigoArticulo").Value)
            'KlexProveedores.TextMatrix(.AbsolutePosition, 8) = EsNulo(.Fields("Descripcion").Value)
            'KlexProveedores.TextMatrix(.AbsolutePosition, 9) = EsNulo(.Fields("ConsumoMinimo").Value)
            
        
            .MoveNext
        Loop
    
    End With

    sqlArticulosProveedorPrecio = ""
    
    If rsArticulosProveedorPrecio.State = 1 Then
        rsArticulosProveedorPrecio.Close
        Set rsArticulosProveedorPrecio = Nothing
    End If
    
If Err Then GrabarLog "CargarArticulosProveedor", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub NuevoProveedorPrecio(vActivar As Boolean)
On Error Resume Next
    
    Dim i As Integer
    
    For i = 0 To txtProveedores.Count - 1
        If Not i = 2 Then lblProveedores(i).Enabled = vActivar
        txtProveedores(i).Enabled = vActivar
        txtProveedores(i).Locked = Not vActivar
        txtProveedores(i).Text = ""
        txtProveedores(i).Tag = ""
    Next

    pbCarga(5).Enabled = vActivar
    
If Err Then GrabarLog "NuevoProveedorPrecio", Err.Number & "  " & Err.Description, Me.Caption
End Sub
Private Sub ModificarProveedorPrecio()
On Error Resume Next
    
    With KlexProveedores
        If Not Trim(.TextMatrix(.RowSel, 1)) = "" Then
            NuevoProveedorPrecio (True)
            txtProveedores(0).Tag = .TextMatrix(.RowSel, 1)
            txtProveedores(0).Text = .TextMatrix(.RowSel, 2)
            txtProveedores(1).Text = .TextMatrix(.RowSel, 3)
            txtProveedores(2).Text = .TextMatrix(.RowSel, 6)
        Else
            MsgBox "Debe seleccionar un Registro para poder modificarlo!!", vbExclamation, "Mensaje ..."
        End If
    End With

    
If Err Then GrabarLog "ModificarProveedorPrecio", Err.Number & "  " & Err.Description, Me.Caption
End Sub
Private Sub BorrarProveedorPrecio()
On Error Resume Next
    
    Dim vIdProveedorArticulo As Long
        
    With KlexProveedores
        If Not Trim(.TextMatrix(.RowSel, 1)) = "" Then
            vIdProveedorArticulo = .Row
            Call BorrarBase("ArticulosProveedorPrecio WHERE (idArticulosProveedorPrecio = " & Val(.TextMatrix(vIdProveedorArticulo, 1)) & ")", pathDBMySQL)
            .RemoveItem (vIdProveedorArticulo)
            
        Else
            MsgBox "Debe seleccionar un Registro para poder Borrarlo!!", vbExclamation, "Mensaje ..."
        End If
    End With
    
If Err Then GrabarLog "BorrarProveedorPrecio", Err.Number & "  " & Err.Description, Me.Caption
End Sub
Private Sub GuardarProveedorPrecio()
On Error Resume Next
    
    With KlexProveedores
        If Not Trim(txtProveedores(0).Tag) = "" Then
            
            Call EjecutarScript("UPDATE articulosproveedorprecio SET PrecioDeCosto = " & Val(txtProveedores(2).Text) & " WHERE (CodigoArticulo = '" & Trim(txtAlta(0).Text) & "') AND (CodigoProveedor = '" & Trim(txtProveedores(0).Text) & "')")
        
        Else
            If ValidarProveedorPrecio(Trim(txtAlta(0).Text), Trim(txtProveedores(0).Text)) = True Then
                Call EjecutarScript("INSERT INTO articulosproveedorprecio (CodigoArticulo, CodigoProveedor, PrecioDeCosto) VALUES ('" & Trim(txtAlta(0).Text) & "','" & Trim(txtProveedores(0).Text) & "','" & Val(txtProveedores(2).Text) & "')")
            End If
        End If
    
    End With
    
If Err Then GrabarLog "GuardarProveedorPrecio", Err.Number & "  " & Err.Description, Me.Caption
End Sub
Private Function ValidarProveedorPrecio(vCodigoArticulo, vCodigoProveedor) As Boolean
On Error Resume Next

    ValidarProveedorPrecio = True
    
    If vCodigoArticulo = "" Then
        MsgBox "Ingrese un Articulo", vbExclamation, "Mensaje ..."
        ValidarProveedorPrecio = False
        Exit Function
    End If
    
    If vCodigoProveedor = "" Then
        MsgBox "Ingrese un proveedor", vbExclamation, "Mensaje ..."
        ValidarProveedorPrecio = False
        Exit Function
    End If

    If Not Val(TraerDato("ArticulosProveedorPrecio", "(CodigoArticulo = '" & vCodigoArticulo & "') AND (CodigoProveedor = '" & vCodigoProveedor & "')", "idArticulosProveedorPrecio")) = 0 Then
        MsgBox "El Articulo-Proveedor ya se encuentra cargado!! ", vbExclamation, "Mensaje ..."
        ValidarProveedorPrecio = False
        Exit Function
    End If

If Err Then GrabarLog "ValidarProveedorPrecio", Err.Number & " " & Err.Description, Me.Caption
End Function

Private Sub FlatEdit3_Change()

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    
    If KeyAscii = 13 Then SendKeys "{TAB}"

    If Err Then GrabarLog "Form_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    If KeyCode = vbKeyF1 Then
        VerAyuda (Me.Name)
    End If
    
    'If KeyCode = vbKeyF2 Then
        'Grabar
    'End If

If Err Then GrabarLog "Form_KeyUp", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Load()
    On Error Resume Next

    With Me
        .Show
        .Top = 0
        .Left = 0
    End With
    
    LimpiarCampos
    CargarTarifas
    
    
  '  CargarGrillaEspeciales ("")
  '  CargarGrillaArticulosProveedores ("")
    'CargarGrillaStock ("")
    
    CentrarFormulario (Me)
    
    Me.TabAlta.Selected = 0

  '  CargarTarifas
  
   lblStock2.Caption = txtStock(0).Text


    If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub CargarGrillaEspeciales(vCodigoArticulo As String)
On Error Resume Next
    
    Dim rsEspeciales As New ADODB.Recordset, sqlEspeciales As String
    
    With KlexArticulosClientes
        .Cols = 10
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 1
        
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 250
        
        .TextMatrix(0, 1) = "idArticulosClientes"
        .ColWidth(1) = 0
        
        .TextMatrix(0, 2) = "Cliente"
        .ColWidth(2) = 1000
        
        .TextMatrix(0, 3) = "Nombre"
        .ColWidth(3) = 3500
        
        .TextMatrix(0, 4) = "CodigoArticulo"
        .ColWidth(4) = 0
        
        .TextMatrix(0, 5) = "Descripcion"
        .ColWidth(5) = 0
        
        .TextMatrix(0, 6) = "Importe"
        .ColWidth(6) = 1250
        .ColDisplayFormat(6) = "#0.00"
        
        .TextMatrix(0, 7) = "idTipoCondicion"
        .ColWidth(7) = 0
                
        .TextMatrix(0, 8) = "Tipo Cond."
        .ColWidth(8) = 1250
                
        .TextMatrix(0, 9) = "ConsumoMinimo"
        .ColWidth(9) = 0
                
        
        .Editable = True
       
        '.EnterKeyBehaviour = klexEKMoveDown
        .EnterKeyBehaviour = klexEKNone
        .BackColorAlternate = &HE0E0E0

        sqlEspeciales = "SELECT * FROM ArticulosClientesEspeciales WHERE (CodigoArticulo = '" & vCodigoArticulo & "') ORDER BY 1;"
    
        Call rsEspeciales.Open(sqlEspeciales, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If Not rsEspeciales.EOF = True Then
            .Rows = rsEspeciales.RecordCount + 1
            rsEspeciales.MoveFirst
        Else
            '.Cols = 10
            '.Rows = 2
        End If
        Do Until rsEspeciales.EOF = True

            .TextMatrix(rsEspeciales.AbsolutePosition, 1) = EsNulo(rsEspeciales.Fields("IdArticulosClientes").Value)
            .TextMatrix(rsEspeciales.AbsolutePosition, 2) = EsNulo(rsEspeciales.Fields("CodigoCliente").Value)
            .TextMatrix(rsEspeciales.AbsolutePosition, 3) = EsNulo(rsEspeciales.Fields("Nombre").Value)
            .TextMatrix(rsEspeciales.AbsolutePosition, 4) = EsNulo(rsEspeciales.Fields("CodigoArticulo").Value)
            .TextMatrix(rsEspeciales.AbsolutePosition, 5) = EsNulo(rsEspeciales.Fields("Descrip").Value)
            .TextMatrix(rsEspeciales.AbsolutePosition, 6) = EsNulo(rsEspeciales.Fields("Precio").Value)
            .TextMatrix(rsEspeciales.AbsolutePosition, 7) = EsNulo(rsEspeciales.Fields("idTipoCondicion").Value)
            .TextMatrix(rsEspeciales.AbsolutePosition, 8) = EsNulo(rsEspeciales.Fields("Descripcion").Value)
            .TextMatrix(rsEspeciales.AbsolutePosition, 9) = EsNulo(rsEspeciales.Fields("ConsumoMinimo").Value)
            
            rsEspeciales.MoveNext
        Loop
            
        '.AutoSizeMode = klexAutoSizeColWidth
        '.AutoSize 2, 2
        '.AutoSize 3, 3
        '.AutoSize 6, 6
        '.AutoSize 8, 8
    
    End With
    
If Err Then GrabarLog "CargarGrillaEspeciales", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub KlexPrecios_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next

    Select Case Col
    
        Case 0
        
        Case 1
        
        Case 2
        
        Case 3
    '    klexPrecios.TextMatrix(Row, 7) = klexPrecios.TextMatrix(Row, 6) + klexPrecios.TextMatrix(Row, 6) * Val(klexPrecios.TextMatrix(Row, 3)) / 100
        
        Case 4
        
        Case 5
        
        Case 6
            klexPrecios.TextMatrix(Row, Col + 1) = klexPrecios.TextMatrix(Row, Col) + klexPrecios.TextMatrix(Row, Col) * Val(klexPrecios.TextMatrix(Row, 3)) / 100
            klexPrecios.ColDisplayFormat(Col) = "#0.00"
            klexPrecios.ColDisplayFormat(Col + 1) = "#0.00"
    End Select

If Err Then GrabarLog "KlexPrecios_AfterEdit", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub klexPrecios_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
On Error Resume Next
        
    Dim vIDTarifas As String
    
    If KeyAscii = 13 Then
        
        
        
        With klexPrecios
            
            vIDTarifas = Replace(Replace(.TextMatrix(Row, 1), "[", ""), "]", "")
            
            Select Case Col
            
                Case 2
                    Call EjecutarScript("UPDATE Tarifas SET Descripcion = '" & .TextMatrix(Row, Col) & "' WHERE (idTarifas = '" & vIDTarifas & "')")
                
                Case 3
                    Call EjecutarScript("UPDATE Tarifas SET Margen = " & .TextMatrix(Row, Col) & " WHERE (idTarifas = '" & vIDTarifas & "')")
                    actualizaPrecio
            End Select
    
        End With
    
    End If
    
If Err Then GrabarLog "KlexPrecios_AfterEdit", Err.Number & " " & Err.Description, Me.Caption
End Sub



Private Sub PbAcciones_Click(Index As Integer)
    On Error Resume Next

    Select Case Index
    
        Case 0
            Grabar
        
        Case 1
            Unload Me

    End Select

    If Err Then GrabarLog "PbAcciones_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub pbCarga_Click(Index As Integer)
    On Error Resume Next

    vVuelveBusqueda = Me.Name
    vVieneBusqueda = pbCarga(Index).Tag

    Select Case Index
    
        Case 0

            'Foto
            With phtArticulo
                .PhotoFileName = ""
                .OpenPhotoFile

                If Not .PhotoFileName = "" Then
                    pbCierraFoto.Visible = True
                Else
                    pbCierraFoto.Visible = Not True
                End If

            End With
        
        Case 1 To 6
            frmBusqueda.Show
    
        Case 7
            vVieneConcepto = Me.Name

            With frmEstructuraCaja
                .leido = False
                .vModo = Modo.Seleccion
                .Show
            End With
            
    End Select

    
    If Err Then GrabarLog "pbCarga_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Grabar()
    On Error Resume Next

    If Not ValidarCampos() = True Then
        Exit Sub
    End If
    
    Dim rsArticulos As New ADODB.Recordset, sqlArticulos As String, i As Integer

    Select Case vaccion
        Case "Nuevo"
            sqlArticulos = "SELECT * FROM Articulos WHERE 1=2"
        
        Case "Modificar"
            sqlArticulos = "SELECT * FROM Articulos WHERE (Codigo =  '" & Trim(txtAlta(0).Text) & "')"
        
        Case "Duplicar"
            
    End Select
        
    With rsArticulos
        .CursorLocation = adUseServer
        Call .Open(sqlArticulos, pathDBMySQL, adOpenDynamic, adLockOptimistic)
        
        If Not .State = 0 Then
        
            Select Case vaccion
            
                Case "Nuevo"
                    .AddNew
                    .Fields("Codigo").Value = Trim(txtAlta(0).Text)
                    .Fields("codigoNum").Value = Val(txtAlta(0).Text)
 
                    .Fields("Stock").Value = GuardarEnStock("Articulo-Nuevo", .Fields("Codigo").Value, Date, Val(txtStock(0).Text), "Stock Inicial", 0, 0)
            
                Case "Modificar"
                    
                      ' si modifica salgo  panic
                    
                    
                   ' If Not .EOF = True Then
                        'Tengo que ver que pasa con el otro
                   '     .Fields("Stock").Value = GuardarEnStock("Articulo-Modificar", .Fields("Codigo").Value, Date, Val(txtStock(0).Text), "Stock Inicial (M)", 0, 0)
                   ' Else
                   '     MsgBox "El registro se ha borrado!!!", vbExclamation, "Mensaje ..."
                   ' End If
                    
                Case "Duplicar"
                    .AddNew
                    .Fields("Codigo").Value = "" 'Tendria que traer el ultimo codigo
                    .Fields("codigo_num").Value = Val(txtAlta(0).Text)
                    .Fields("Stock").Value = 0
            
            End Select
            
            'No Opcional
            .Fields("Descrip").Value = Left(txtAlta(1).Text, 255)
            .Fields("idSubRubros").Value = Left(txtAlta(2).Text, 3)
            .Fields("idRubros").Value = Left(txtAlta(4).Text, 3)
            Call GuardarFoto(rsArticulos, phtArticulo.PhotoFileName)
        
            'Ficha
            .Fields("idPorcentajeIva").Value = Left(txtFicha(0).Text, 3)
            .Fields("idProveedor").Value = txtFicha(3).Text
            .Fields("idFabricantes").Value = EsNulo(Left(txtFicha(5).Text, 150))
            .Fields("PCosto").Value = Val(Format(txtFicha(7).Text, "#####0.00"))
            
            .Fields("CodigoBarra").Value = txtAlta(6).Text
            
            .Fields("PVenta1").Value = Val(txtFicha(8).Text)
            
            .Fields("PVenta2").Value = Val(txtFicha(9).Text)
            .Fields("PVenta3").Value = Val(txtFicha(10).Text)
            .Fields("PVenta4").Value = Val(txtFicha(11).Text)
            .Fields("PVenta5").Value = Val(txtFicha(12).Text)
            
    
            
            
           ' For i = 1 To klexPrecios.Rows - 2
           '     .Fields("PVenta" & i).Value = Val(Format(klexPrecios.TextMatrix(i, 7), "#####0.00"))
           ' Next
            
            'Tecnica
            .Fields("FechaAlta").Value = strfechaMySQL(dtpFecha(0).Value)
            
            If chkActualizacionDePrecio.Value = xtpUnchecked Then
                If Val(vCostoAnterior) <> Val(txtFicha(7).Text) Then
                    .Fields("FechaModificacion").Value = strfechaMySQL(Date)
                Else
                    If Val(vPVentaAnterior) <> Val(txtFicha(8).Text) Then
                        .Fields("FechaModificacion").Value = strfechaMySQL(Date)
                    End If
                End If
            Else
                .Fields("FechaModificacion").Value = strfechaMySQL(dtpFecha(1).Value)
            End If
            
            .Fields("Peso_U").Value = Val(txtTecnica(0).Text)
            .Fields("Peso_T").Value = Val(txtTecnica(1).Text)
            .Fields("UnidadesPorBulto").Value = Val(txtTecnica(2).Text)
            .Fields("Dimensiones").Value = EsNulo(txtTecnica(3).Text)
            .Fields("MensajeEmergente").Value = EsNulo(txtTecnica(4).Text)
            .Fields("CodigoConcepto").Value = Val(txtTecnica(5).Text)
            .Fields("Observaciones").Value = EsNulo(txtTecnica(7).Text)

            
            'Stock
            .Fields("Stock").Value = Val(txtStock(0).Text)
            .Fields("StockMin").Value = Val(txtStock(1).Text)
            .Fields("StockMax").Value = Val(txtStock(2).Text)
            .Fields("idDepositos").Value = EsNulo(cboDepositos.Tag)
            '.Fields("Faltante").Value = Val(txtStock(3).Text)

            .Update
        
        End If
        
    End With

    sqlArticulos = ""
    
    If rsArticulos.State = 1 Then
        rsArticulos.Close
        Set rsArticulos = Nothing
    End If
    
    If Err Then
        GrabarLog "Guardar", Err.Number & " " & Err.Description, Me.Name

    Else

    End If
    
    
    Select Case vViene
    
        Case "frmConsultas"
        
                        frmConsultas.vbuscando.Text = Me.txtAlta(1).Text
                        frmConsultas.WindowState = 2
                        Unload Me
    
    
        Case "frmRemito"
    
                        frmRemito.txtDetalle(1).Text = Me.txtAlta(1).Text
                        frmRemito.txtDetalle(1).Tag = Me.txtAlta(0).Text
                        frmRemito.WindowState = 2
                        Unload Me
    
        Case Else
        
                      frmArticulos.Buscar ("")
        
    
    End Select
    
       ' If vViene = "frmConsultas" Then
       '         frmConsultas.vbuscando.Text = Me.txtAlta(1).Text
       '         frmConsultas.WindowState = 2
       '         Unload Me
       ' Else
       '         frmArticulos.Buscar ("")
       ' End If
    
    LimpiarCampos
    Unload Me
    

End Sub
Private Function ValidarCampos() As Boolean
    On Error Resume Next

     Dim i As Integer
    
     ValidarCampos = True
    
     ValidarCampos = True
     Exit Function
    
    
    
  '  For i = 0 To Val(txtAlta.Count - 2)

  '      If Trim(txtAlta(i).Text) = "" Then
  '          MsgBox "Existen campos de ingreso obligatorio vacíos.", vbExclamation, "Mensaje ..."
  '          'ValidarCampos = Not True
  '          ValidarCampos = True
  '          Exit Function
  '      End If
'
 '   Next
    
    If Me.txtFicha(0).Text = "" Then
        MsgBox "El campo porcentaje IVA es de ingreso obligatorio.", vbExclamation, "Mensaje ..."
        ValidarCampos = Not True
        Exit Function
    End If
    
    If vaccion = "Nuevo" Then
        If Not Trim(TraerDato("Articulos", "Codigo = '" & Trim(txtAlta(0).Text) & "'", "Codigo")) = "" Then
            MsgBox "Ya existe un registro con ese codigo.", vbExclamation, "Mensaje ..."
            ValidarCampos = Not True
            Exit Function
        End If
    End If
    
    If Err Then GrabarLog "ValidarCampos", Err.Number & " " & Err.Description, Me.Caption
End Function
Private Function GuardarFoto(rsFoto As ADODB.Recordset, _
                             vFilename As String) As Boolean
    On Error Resume Next
    
    If Not vFilename = "" Then
    
        Dim stmFoto As New ADODB.Stream
    
        With stmFoto

            If Not .State = adStateClosed Then .Close
            
            .Type = adTypeBinary
            .Open

            .LoadFromFile (vFilename)
        End With
    
    End If
    
    With rsFoto

        If Not (.State = 0) And Not (.EOF = True) Then
            If Not vFilename = "" Then
                .Fields("Foto").Value = stmFoto.Read
            Else
                .Fields("Foto").Value = Null
            End If
        End If
        
    End With
    
    If Not stmFoto.State = adStateClosed Then
        stmFoto.Close
        Set stmFoto = Nothing
    End If
    
    If Err Then
        GuardarFoto = Not True
        GrabarLog "GuardarFoto", Err.Number & " " & Err.Description, "Global"
    Else
        GuardarFoto = True
    End If

End Function

Private Sub LimpiarCampos()
    On Error Resume Next
    
    Dim i As Integer
    
    For i = 0 To txtAlta.Count - 1
        txtAlta(i).Text = ""
    Next
    
    phtArticulo.Reset
    pbCierraFoto.Visible = Not True
    
    For i = 0 To txtFicha.Count - 1
        txtFicha(i).Text = ""
    Next
    
    For i = 0 To txtTecnica.Count - 1
        txtTecnica(i).Text = ""
    Next
    
    For i = 0 To txtStock.Count - 1
        txtStock(i).Text = ""
    Next
    For i = 0 To dtpFecha.Count - 1
        dtpFecha(i).Value = Date
    Next
    
    vaccion = "Nuevo"
    txtAlta(0).Locked = Not True
    txtAlta(0).Text = Val(GenerarDato("SELECT idArticulos, Codigo FROM Articulos ORDER BY idArticulos DESC", "Codigo")) + 1
    txtAlta(0).Text = FormatoUltimoCodigo(5, txtAlta(0).Text)
    
    KeyPreview = True

    If Err Then GrabarLog "Limpia", Err.Number & "-" & Err.Description, Me.Name
End Sub
Public Sub ModificarArticulo(vidArticulo As Long)
    On Error Resume Next

    Dim rsArticulo As New ADODB.Recordset, sqlArticulo As String
    
    sqlArticulo = "SELECT * FROM Articulos WHERE (idArticulos = " & vidArticulo & ")"
    
    vCostoAnterior = 0
    vPVentaAnterior = 0
            
    With rsArticulo
        Call .Open(sqlArticulo, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not (.EOF = True) And Not (.BOF = True) Then
        
            'No Opcionales
            txtAlta(0).Text = .Fields("codigo").Value
            txtAlta(0).Locked = True
        
            txtAlta(1).Text = EsNulo(.Fields("Descrip").Value)
            
            txtAlta(2).Text = EsNulo(.Fields("idSubRubros").Value)
            txtAlta(3).Text = EsNulo(TraerDato("SubRubros", "idSubRubros = '" & .Fields("idSubRubros").Value & "'", "SubRubro"))
            
            txtAlta(4).Text = EsNulo(.Fields("idRubros").Value)
            txtAlta(5).Text = EsNulo(TraerDato("Rubros", "idRubros = '" & .Fields("idRubros").Value & "'", "Rubro"))
            txtAlta(6).Text = EsNulo(.Fields("CodigoBarra").Value)
            
            
            
            If Not IsNull(.Fields("Foto").Value) = True And Not Trim(.Fields("Foto").Value) = "" Then
                BorrarArchivo (App.Path & "\" & .Fields("Codigo").Value & ".dat")
                phtArticulo.BlobToFile rsArticulo!Foto, App.Path & "\" & .Fields("Codigo").Value & ".dat"
                Call phtArticulo.AbrirFotoDesdeArchivo(App.Path & "\" & .Fields("Codigo").Value & ".dat")
                BorrarArchivo (App.Path & "\" & .Fields("Codigo").Value & ".dat")
                pbCierraFoto.Visible = True
            End If

            'Ficha
        
            txtFicha(0).Text = EsNulo(.Fields("idPorcentajeIva").Value)
            txtFicha(1).Text = EsNulo(TraerDato("PorcentajeIva", "idPorcentajeIva = '" & .Fields("idPorcentajeIva").Value & "'", "Descripcion"))
            txtFicha(2).Text = EsNulo(TraerDato("PorcentajeIva", "idPorcentajeIva = '" & .Fields("idPorcentajeIva").Value & "'", "Porcentaje"))
            txtFicha(3).Text = EsNulo(.Fields("idProveedor").Value)
            txtFicha(4).Text = EsNulo(TraerDato("Proveedores", "Codigo = " & .Fields("idProveedor").Value & "", "Nombre"))
            txtFicha(5).Text = EsNulo(.Fields("idFabricantes").Value)
            txtFicha(6).Text = EsNulo(TraerDato("Fabricantes", "idFabricantes = '" & .Fields("idFabricantes").Value & "'", "Nombre"))
            txtFicha(7).Text = EsNulo(.Fields("PCosto").Value)
            
            txtFicha(8).Text = EsNulo(.Fields("PVenta1").Value)
            
            txtFicha(9).Text = EsNulo(.Fields("PVenta2").Value)
            txtFicha(10).Text = EsNulo(.Fields("PVenta3").Value)
            txtFicha(11).Text = EsNulo(.Fields("PVenta4").Value)
            txtFicha(12).Text = EsNulo(.Fields("PVenta5").Value)
            
            
            vCostoAnterior = EsNulo(.Fields("PCosto").Value)
            vPVentaAnterior = EsNulo(.Fields("PVenta1").Value)
        
            
            'CargarTarifas
        
            'Tecnica
            dtpFecha(0).Value = EsNulo(strfechaMySQL(.Fields("FechaAlta").Value))
            dtpFecha(1).Value = EsNulo(strfechaMySQL(.Fields("FechaModificacion").Value))
            txtTecnica(0).Text = EsNulo(.Fields("Peso_u").Value)
            txtTecnica(1).Text = EsNulo(.Fields("Peso_t").Value)
            txtTecnica(2).Text = EsNulo(.Fields("UnidadesPorBulto").Value)
            txtTecnica(3).Text = EsNulo(.Fields("Dimensiones").Value)
            txtTecnica(4).Text = EsNulo(.Fields("MensajeEmergente").Value)
            txtTecnica(5).Text = EsNulo(.Fields("CodigoConcepto").Value)
            txtTecnica(6).Text = EsNulo(TraerDato("Concepto", "Codigo = " & Val(txtTecnica(6).Text) & "", "Descripcion"))
            txtTecnica(7).Text = EsNulo(.Fields("Observaciones").Value)
        
            
            'Cond. Especiales de Venta
            CargarGrillaEspeciales (.Fields("Codigo").Value)
            CargarGrillaArticulosProveedores (.Fields("Codigo").Value)
            CargarGrillaStock (.Fields("Codigo").Value)

            'Stock
            txtStock(0).Text = EsNulo(.Fields("Stock").Value)
            txtStock(1).Text = EsNulo(.Fields("StockMin").Value)
            txtStock(2).Text = EsNulo(.Fields("StockMax").Value)
            cboDepositos.Tag = EsNulo(.Fields("idDepositos").Value)
            cboDepositos.Text = EsNulo(TraerDato("Depositos", "idDepositos = '" & cboDepositos.Tag & "'", "Deposito"))
            'txtStock(3).Text = EsNulo(.Fields("Faltante").Value)
        
        End If

    End With
        
    sqlArticulo = ""
    
    If rsArticulo.State = 1 Then
        rsArticulo.Close
        Set rsArticulo = Nothing
    End If
    
    
    CargarTarifas
       
    lblStock2.Caption = txtStock(0).Text

    
    If Err Then GrabarLog "ModificarArticulo", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub pbCierraFoto_Click()
    On Error Resume Next

    phtArticulo.Reset
    pbCierraFoto.Visible = Not True

    If Err Then GrabarLog "pbCierraFoto_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CargarTarifas()
On Error Resume Next

Dim i As Integer
    
    Me.Show
    
    Dim rsTarifas As New ADODB.Recordset, sqlTarifas As String
    
    sqlTarifas = "SELECT * FROM Tarifas ORDER BY 1"
    
    With rsTarifas
        Call .Open(sqlTarifas, ConnDDBB, adOpenStatic, adLockReadOnly)
    
        If Not .EOF = True Then
            .MoveFirst
            FormatoGrillaPrecio (.RecordCount)
        End If
        
        i = 0
        Do Until .EOF = True
           i = i + 1
            klexPrecios.TextMatrix(.AbsolutePosition, 1) = "[" & .Fields("idTarifas").Value & "]"
            klexPrecios.TextMatrix(.AbsolutePosition, 2) = .Fields("Descripcion").Value
            klexPrecios.TextMatrix(.AbsolutePosition, 3) = .Fields("Margen").Value
            klexPrecios.TextMatrix(.AbsolutePosition, 4) = .Fields("IvaIncluido").Value
            klexPrecios.TextMatrix(.AbsolutePosition, 3) = "Lista " + Str(i)
            klexPrecios.TextMatrix(.AbsolutePosition, 7) = traerDatos2("select * from articulos where codigo ='" + Me.txtAlta(0) + "'", "Pventa" + Trim(Str(i)), pathDBMySQL)
        
        
            .MoveNext
        Loop
        
    End With

    sqlTarifas = ""

    If rsTarifas.State = 1 Then
        rsTarifas.Close
        Set rsTarifas = Nothing
    End If
    
    'KlexPrecios.Editable = True
If Err Then GrabarLog "CargarTarifas", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub CargarGrillaStock(vCodigoArticulo As String)
On Error Resume Next

    Dim rsStock As New ADODB.Recordset, sqlStock As String, vSaldoStock As Long
    
    sqlStock = "SELECT * FROM Stock WHERE (CodigoArticulo = '" & vCodigoArticulo & "') ORDER BY 1"
    
    With rsStock
        Call .Open(sqlStock, ConnDDBB, adOpenStatic, adLockReadOnly)
    
        If Not .EOF = True Then
            .MoveFirst
            FormatoGrillaStock (.RecordCount)
        Else
            FormatoGrillaStock (1)
        End If
        
        vSaldoStock = 0
        
        Do Until .EOF = True
        
            vSaldoStock = vSaldoStock + Val(.Fields("Entrada").Value) - Val(.Fields("Salida").Value)
            
            KlexStock.TextMatrix(.AbsolutePosition, 1) = .Fields("idStock").Value
            KlexStock.TextMatrix(.AbsolutePosition, 2) = .Fields("Fecha").Value
            KlexStock.TextMatrix(.AbsolutePosition, 3) = .Fields("CodigoArticulo").Value
            KlexStock.TextMatrix(.AbsolutePosition, 4) = .Fields("Entrada").Value
            KlexStock.TextMatrix(.AbsolutePosition, 5) = .Fields("Salida").Value
            KlexStock.TextMatrix(.AbsolutePosition, 6) = vSaldoStock
            KlexStock.TextMatrix(.AbsolutePosition, 7) = .Fields("Comentario").Value
            KlexStock.TextMatrix(.AbsolutePosition, 8) = .Fields("idFDetalle").Value
            KlexStock.TextMatrix(.AbsolutePosition, 9) = .Fields("idPFDetalle").Value
        
            .MoveNext
        Loop
        
    End With

    sqlStock = ""

    If rsStock.State = 1 Then
        rsStock.Close
        Set rsStock = Nothing
    End If
    
If Err Then GrabarLog "CargarGrillaStock", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub FormatoGrillaPrecio(vCantidadRenglones As Integer)
On Error Resume Next

    With klexPrecios
        .FixedRows = 1
        .FixedCols = 1
    
        .Cols = 8
        .Rows = vCantidadRenglones + 1
                
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 250
        
        .TextMatrix(0, 1) = "Codigo"
        .ColWidth(1) = 750
        
        .TextMatrix(0, 2) = "Descripcion"
        .ColWidth(2) = 2000
        
        .TextMatrix(0, 3) = "Margen"
        .ColWidth(3) = 1500
        .ColDisplayFormat(3) = "##0.00"
        
        .TextMatrix(0, 4) = "Incluye Iva"
        .ColWidth(4) = 1000
        
        .TextMatrix(0, 5) = "Descripcion"
        .ColWidth(5) = 0
        
        .TextMatrix(0, 6) = "P. Costo"
        .ColWidth(6) = 1250
        .ColDisplayFormat(6) = "#0.00"
        
        .TextMatrix(0, 7) = "P. Venta"
        .ColWidth(7) = 1250
        .ColDisplayFormat(7) = "#0.00"
                

        .Editable = True

        '.EnterKeyBehaviour = klexEKMoveDown
        .EnterKeyBehaviour = klexEKNone
        .BackColorAlternate = &HE0E0E0

    End With
    
If Err Then GrabarLog "FormatoGrillaPrecio", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub FormatoGrillaStock(vCantidadRenglones As Integer)
On Error Resume Next

    With KlexStock
        .FixedRows = 1
        .FixedCols = 1
    
        .Cols = 10
        .Rows = vCantidadRenglones + 1
                
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 250
        
        .TextMatrix(0, 1) = ""
        .ColWidth(1) = 0
        
        
        .TextMatrix(0, 2) = "Fecha"
        .ColWidth(2) = 1000
        
        .TextMatrix(0, 3) = "Codigo"
        .ColWidth(3) = 0
        
        .TextMatrix(0, 4) = "Entrada"
        .ColWidth(4) = 1250
        .ColDisplayFormat(4) = "##0.00"
        
        .TextMatrix(0, 5) = "Salida"
        .ColWidth(5) = 1250
        .ColDisplayFormat(5) = "##0.00"
        
        .TextMatrix(0, 6) = "Saldo"
        .ColWidth(6) = 1250
        .ColDisplayFormat(6) = "##0.00"
        
        .TextMatrix(0, 7) = "Comentario"
        .ColWidth(7) = 2500
        
        .TextMatrix(0, 8) = "idFDetalle"
        .ColWidth(8) = 0
        
        .TextMatrix(0, 9) = "idPFDetalle"
        .ColWidth(9) = 0
                
        '.Editable = True

        '.EnterKeyBehaviour = klexEKMoveDown
        .EnterKeyBehaviour = klexEKNone
        '.BackColorAlternate = &HE0E0E0

    End With
    
If Err Then GrabarLog "FormatoGrillaStock", Err.Number & " " & Err.Description, Me.Caption
End Sub


Private Sub PusCambiarStock_Click()
txtStock(0) = Val(txtStock(0)) + Val(vstocka)
lblStock2.Caption = txtStock(0).Text
End Sub

Private Sub PushButton1_Click()
Dim vsql As String

If MsgBox("Estás seguro ?", vbYesNo) = vbNo Then Exit Sub


vsql = "delete from stock where CodigoArticulo='" + txtAlta(0) + "'"

Call EjecutarScript(vsql, pathDBMySQL)

Call CargarGrillaStock(txtAlta(0))

End Sub

Private Sub PushButton2_Click()
On Error Resume Next

Dim vsql As String
Dim vrow As Integer

If Not MsgBox("Está seuro que desea borrar la linea ?", vbYesNo) = vbYes Then Exit Sub

vrow = Me.KlexStock.TextMatrix(Me.KlexStock.Row, 1)

If Not vrow > 0 Then Exit Sub


vsql = "delete from stock where idstock = " + Str(vrow)

Call EjecutarScript(vsql, pathDBMySQL)

Call CargarGrillaStock(txtAlta(0))

If Err Then Exit Sub

End Sub

Private Sub PusVerMovimientos_Click(Index As Integer)

Unload Me
End Sub

Private Sub txtAlta_Change(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 0
            If vaccion = "Modificar" Then
                
            End If
        Case 1
        
        Case 2
        
    End Select

If Err Then GrabarLog "txtAlta_Changes", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtFicha_Change(Index As Integer)
On Error Resume Next

    Dim i As Integer

    Select Case Index
    
        Case 7
            
            With klexPrecios
                For i = 1 To .Rows - 1
                    .TextMatrix(i, 6) = Val(txtFicha(Index).Text)
                    
                    If i = 1 Then
                        If Not Trim(txtFicha(8).Text) = "" Then
                            .TextMatrix(i, 7) = .TextMatrix(i, 7)
                        Else
                            .TextMatrix(i, 7) = Val(.TextMatrix(i, 6)) + Val(.TextMatrix(i, 6)) * Val(.TextMatrix(i, 3)) / 100
                        End If
                    Else
                        .TextMatrix(i, 7) = Val(.TextMatrix(i, 6)) + Val(.TextMatrix(i, 6)) * Val(.TextMatrix(i, 3)) / 100
                    End If
                    .ColDisplayFormat(6) = "#0.00"
                    .ColDisplayFormat(7) = "#0.00"
                Next
            End With
        Case 8
                actualizaPrecio
            
            'End With
    End Select

If Err Then GrabarLog "txtFicha_Change", Err.Number & " " & Err.Description, Me.Caption
End Sub


Private Sub actualizaPrecio()
Dim i As Integer
With klexPrecios
                For i = 1 To 1
                    .TextMatrix(i, 6) = Val(txtFicha(7).Text)
                    .TextMatrix(i, 7) = Val(txtFicha(8).Text)
                    .ColDisplayFormat(6) = "#0.00"
                    .ColDisplayFormat(7) = "#0.00"
                Next
            
            End With
End Sub

