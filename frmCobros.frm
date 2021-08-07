VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.ShortcutBar.v13.0.0.Demo.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "Copia de KlexGrid.ocx"
Object = "{9746E3DA-06E1-4D26-9CE4-D9F6411A9C70}#1.0#0"; "SMGA_OcxTxt2008.ocx"
Begin VB.Form frmCobros 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gestion de Cobro de Clientes - Proveedores - Empleados - Acopios "
   ClientHeight    =   8955
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   14280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   14280
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox vconcepto 
      Height          =   315
      Left            =   3120
      TabIndex        =   163
      Top             =   1470
      Width           =   11085
   End
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   495
      Left            =   60
      TabIndex        =   127
      Top             =   -90
      Width           =   14205
      _Version        =   851968
      _ExtentX        =   25056
      _ExtentY        =   873
      _StockProps     =   79
      BackColor       =   -2147483644
      Appearance      =   3
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit vbc1 
         Height          =   285
         Left            =   7080
         TabIndex        =   153
         Top             =   210
         Visible         =   0   'False
         Width           =   1305
         _Version        =   851968
         _ExtentX        =   2302
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "Vbc1"
      End
      Begin XtremeSuiteControls.FlatEdit vbc0 
         Height          =   285
         Left            =   5700
         TabIndex        =   152
         Top             =   210
         Visible         =   0   'False
         Width           =   1305
         _Version        =   851968
         _ExtentX        =   2302
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "Vbc0"
      End
      Begin XtremeSuiteControls.PushButton cmdCobrar 
         Height          =   375
         Left            =   60
         TabIndex        =   7
         Top             =   120
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Ejecutar pago <F2>"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCobros.frx":0000
      End
      Begin XtremeSuiteControls.PushButton PusGrabar 
         Height          =   375
         Index           =   0
         Left            =   30
         TabIndex        =   128
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Ejecutar"
         Enabled         =   0   'False
         Appearance      =   4
         Picture         =   "frmCobros.frx":059A
         BorderGap       =   10
      End
      Begin XtremeSuiteControls.PushButton PusCerrar 
         Height          =   375
         Index           =   1
         Left            =   8880
         TabIndex        =   129
         Top             =   90
         Visible         =   0   'False
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cerrar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCobros.frx":09A1
      End
      Begin XtremeSuiteControls.PushButton PusImprimirComprobante 
         Height          =   375
         Left            =   12090
         TabIndex        =   161
         Top             =   120
         Width           =   2115
         _Version        =   851968
         _ExtentX        =   3731
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Imprimir Comprobante"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCobros.frx":0F3B
         BorderGap       =   10
      End
   End
   Begin VB.Frame FraDocumentosImpagos 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   60
      TabIndex        =   112
      Top             =   990
      Width           =   14205
      Begin VB.TextBox vnroOrdenPago 
         Height          =   315
         Left            =   3930
         TabIndex        =   1
         Top             =   90
         Width           =   1815
      End
      Begin Aplisoft_CajasDeTexto.TxF vfechaCredito 
         Height          =   330
         Left            =   7440
         TabIndex        =   3
         Top             =   90
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   582
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
      Begin Grid.KlexGrid KlexDetalle 
         Height          =   195
         Left            =   3810
         TabIndex        =   113
         ToolTipText     =   "Documentos a cobrar"
         Top             =   600
         Visible         =   0   'False
         Width           =   8385
         _ExtentX        =   14790
         _ExtentY        =   344
         EnterKeyBehaviour=   0
         GridLinesFixed  =   2
         AllowUserResizing=   1
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
         MouseIcon       =   "frmCobros.frx":194D
         SelectionMode   =   1
      End
      Begin VB.Label Label20 
         Caption         =   "> Nro.Orden Pago asociada a un Doc. externo: "
         Height          =   225
         Left            =   120
         TabIndex        =   160
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label Label16 
         Caption         =   "para la cuenta corriente"
         Height          =   195
         Left            =   9540
         TabIndex        =   150
         Top             =   180
         Width           =   1755
      End
      Begin XtremeSuiteControls.Label Label9 
         Height          =   315
         Left            =   11640
         TabIndex        =   131
         Top             =   150
         Width           =   615
         _Version        =   851968
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "> Saldo:"
         ForeColor       =   0
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label vsaldo 
         Height          =   315
         Left            =   12540
         TabIndex        =   130
         Top             =   120
         Width           =   1515
         _Version        =   851968
         _ExtentX        =   2672
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "0.00"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin VB.Label lblIngFecha 
         Caption         =   "> Fecha operación:"
         Height          =   195
         Left            =   5940
         TabIndex        =   126
         Top             =   150
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   585
      Left            =   30
      TabIndex        =   103
      Top             =   8310
      Width           =   14175
      Begin VB.CommandButton Command2 
         Caption         =   "Imprimir recibo 2"
         Height          =   255
         Left            =   10050
         TabIndex        =   169
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Recibo2 - 2"
         Height          =   255
         Left            =   11700
         TabIndex        =   168
         Top             =   240
         Width           =   1065
      End
      Begin VB.CommandButton cmdRecibo2 
         Caption         =   "Recibo2 -1"
         Height          =   255
         Left            =   12930
         TabIndex        =   167
         Top             =   240
         Width           =   1065
      End
      Begin VB.TextBox TxtTotalAPagar 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1770
         TabIndex        =   105
         Top             =   210
         Width           =   1455
      End
      Begin VB.TextBox txtMontoTotalPendienteSeleccionado 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6840
         TabIndex        =   104
         Top             =   210
         Width           =   1455
      End
      Begin VB.Label lblTotalA 
         BackStyle       =   0  'Transparent
         Caption         =   "> Importe total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   107
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblMontoTotalPendente 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Doc. Seleccionados a pagar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   3540
         TabIndex        =   106
         Top             =   270
         Width           =   3225
      End
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      CausesValidation=   0   'False
      Height          =   5835
      Left            =   0
      TabIndex        =   50
      Top             =   1830
      Width           =   14205
      _Version        =   851968
      _ExtentX        =   25056
      _ExtentY        =   10292
      _StockProps     =   68
      PaintManager.BoldSelected=   -1  'True
      PaintManager.MultiRowFixedSelection=   -1  'True
      ItemCount       =   6
      Item(0).Caption =   "Importe Efectivo"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "GroupBox1"
      Item(1).Caption =   "Cheques Tercero"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "fraCheques"
      Item(2).Caption =   "Ingreso de Tarjetas"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "fraTarjeta"
      Item(3).Caption =   "Retenciones"
      Item(3).ControlCount=   6
      Item(3).Control(0)=   "GroupBox2"
      Item(3).Control(1)=   "PushButton4"
      Item(3).Control(2)=   "gRetencion"
      Item(3).Control(3)=   "PushButton3"
      Item(3).Control(4)=   "Label14"
      Item(3).Control(5)=   "vtimporteRetenciones"
      Item(4).Caption =   "Cheques propios"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "fraDepositos"
      Item(5).Caption =   "Documentos Selecionados"
      Item(5).ControlCount=   1
      Item(5).Control(0)=   "dgSeleccionados"
      Begin MSDataGridLib.DataGrid dgSeleccionados 
         Height          =   5235
         Left            =   -69910
         TabIndex        =   166
         Top             =   450
         Visible         =   0   'False
         Width           =   13995
         _ExtentX        =   24686
         _ExtentY        =   9234
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
               LCID            =   3082
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
               LCID            =   3082
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gRetencion 
         Height          =   3825
         Left            =   -69520
         TabIndex        =   141
         Top             =   1560
         Visible         =   0   'False
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   6747
         _Version        =   393216
         Cols            =   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   615
         Left            =   -69550
         TabIndex        =   135
         Top             =   450
         Visible         =   0   'False
         Width           =   13425
         _Version        =   851968
         _ExtentX        =   23680
         _ExtentY        =   1085
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.PushButton PushButton2 
            Height          =   315
            Left            =   3060
            TabIndex        =   136
            Tag             =   "Banco"
            Top             =   210
            Width           =   315
            _Version        =   851968
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit vdretencion 
            Height          =   315
            Left            =   3420
            TabIndex        =   137
            Top             =   210
            Width           =   6735
            _Version        =   851968
            _ExtentX        =   11880
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   255
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vcretencion 
            Height          =   315
            Left            =   1950
            TabIndex        =   138
            Top             =   210
            Width           =   1065
            _Version        =   851968
            _ExtentX        =   1879
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   255
            BackColor       =   -2147483643
            Alignment       =   1
         End
         Begin XtremeSuiteControls.FlatEdit vImporteRet 
            Height          =   315
            Left            =   11700
            TabIndex        =   143
            Top             =   210
            Width           =   1635
            _Version        =   851968
            _ExtentX        =   2884
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   255
            BackColor       =   -2147483643
         End
         Begin VB.Label Label13 
            Caption         =   "> Importe:"
            Height          =   195
            Left            =   10770
            TabIndex        =   144
            Top             =   270
            Width           =   855
         End
         Begin VB.Label Label12 
            Caption         =   "Seleccione retención:"
            Height          =   195
            Left            =   120
            TabIndex        =   139
            Top             =   270
            Width           =   1845
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   3225
         Left            =   60
         TabIndex        =   89
         Top             =   540
         Width           =   14085
         _Version        =   851968
         _ExtentX        =   24844
         _ExtentY        =   5689
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
         Begin XtremeSuiteControls.GroupBox GroupBox4 
            Height          =   780
            Left            =   6300
            TabIndex        =   170
            Top             =   2295
            Width           =   4380
            _Version        =   851968
            _ExtentX        =   7726
            _ExtentY        =   1376
            _StockProps     =   79
            Caption         =   "Tipos de Movimientos "
            UseVisualStyle  =   -1  'True
            Begin XtremeSuiteControls.RadioButton RadioButton1 
               Height          =   435
               Left            =   135
               TabIndex        =   171
               Top             =   270
               Width           =   975
               _Version        =   851968
               _ExtentX        =   1720
               _ExtentY        =   767
               _StockProps     =   79
               Caption         =   "Todos"
               Appearance      =   6
               Value           =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton radioTodoDoc 
               Height          =   405
               Left            =   1350
               TabIndex        =   172
               Top             =   270
               Width           =   1305
               _Version        =   851968
               _ExtentX        =   2302
               _ExtentY        =   714
               _StockProps     =   79
               Caption         =   "Sólo DOC"
               Appearance      =   6
            End
            Begin XtremeSuiteControls.RadioButton Radsolofact 
               Height          =   255
               Left            =   2925
               TabIndex        =   173
               Top             =   360
               Width           =   1305
               _Version        =   851968
               _ExtentX        =   2302
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Sólo FACT"
               Appearance      =   6
            End
         End
         Begin XtremeSuiteControls.PushButton PusCopiarTotal 
            Height          =   315
            Left            =   3570
            TabIndex        =   162
            Top             =   570
            Width           =   2745
            _Version        =   851968
            _ExtentX        =   4842
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Copiar Total Doc. Seleccionado"
            UseVisualStyle  =   -1  'True
         End
         Begin Aplisoft_CajasDeTexto.TxF dtpFecha 
            Height          =   330
            Left            =   8670
            TabIndex        =   5
            Top             =   600
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   582
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
         Begin XtremeSuiteControls.FlatEdit txtCotizacionDolar 
            Height          =   315
            Left            =   2130
            TabIndex        =   93
            Top             =   1350
            Width           =   1335
            _Version        =   851968
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Alignment       =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtImporteEfectivoDolar 
            Height          =   315
            Left            =   2130
            TabIndex        =   91
            Top             =   960
            Width           =   1335
            _Version        =   851968
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Alignment       =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtImporteEfectivoPesos 
            Height          =   315
            Left            =   2130
            TabIndex        =   4
            Top             =   570
            Width           =   1335
            _Version        =   851968
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   255
            BackColor       =   -2147483643
            Alignment       =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtNroInterno 
            Height          =   315
            Left            =   8670
            TabIndex        =   108
            Top             =   990
            Width           =   1995
            _Version        =   851968
            _ExtentX        =   3519
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Alignment       =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtBancoCheque 
            Height          =   315
            Index           =   6
            Left            =   2130
            TabIndex        =   94
            Top             =   1740
            Width           =   1335
            _Version        =   851968
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.PushButton pbCarga 
            Height          =   315
            Index           =   6
            Left            =   3570
            TabIndex        =   117
            Tag             =   "caja-importe-cobro"
            Top             =   1770
            Width           =   315
            _Version        =   851968
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtBancoCheque 
            Height          =   315
            Index           =   7
            Left            =   4020
            TabIndex        =   95
            Top             =   1770
            Width           =   6645
            _Version        =   851968
            _ExtentX        =   11721
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Caja:"
            Height          =   195
            Index           =   3
            Left            =   1560
            TabIndex        =   118
            Top             =   1800
            Width           =   435
         End
         Begin VB.Label lblCobros 
            BackStyle       =   0  'Transparent
            Caption         =   "Nro Interno:"
            Height          =   255
            Index           =   4
            Left            =   7620
            TabIndex        =   99
            Top             =   1020
            Width           =   975
         End
         Begin VB.Label lblCobros 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha :"
            Height          =   255
            Index           =   3
            Left            =   7920
            TabIndex        =   98
            Top             =   630
            Width           =   975
         End
         Begin VB.Label lblCobros 
            BackStyle       =   0  'Transparent
            Caption         =   "Cotizacion:"
            Height          =   255
            Index           =   2
            Left            =   1170
            TabIndex        =   97
            Top             =   1365
            Width           =   795
         End
         Begin VB.Label lblCobros 
            BackStyle       =   0  'Transparent
            Caption         =   "Importe efectivo dólar:"
            Height          =   255
            Index           =   1
            Left            =   390
            TabIndex        =   96
            Top             =   1005
            Width           =   1575
         End
         Begin VB.Label lblCobros 
            BackStyle       =   0  'Transparent
            Caption         =   "Importe efectivo pesos:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   330
            TabIndex        =   92
            Top             =   630
            Width           =   1815
         End
      End
      Begin VB.Frame fraDepositos 
         Height          =   5385
         Left            =   -69940
         TabIndex        =   73
         Top             =   360
         Visible         =   0   'False
         Width           =   14025
         Begin XtremeSuiteControls.FlatEdit txtDepositoBanco 
            Height          =   315
            Index           =   0
            Left            =   1200
            TabIndex        =   79
            Top             =   270
            Width           =   1995
            _Version        =   851968
            _ExtentX        =   3519
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.PushButton pbCarga 
            Height          =   315
            Index           =   3
            Left            =   3330
            TabIndex        =   80
            Tag             =   "BancoDeposito"
            Top             =   270
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
            Left            =   3300
            TabIndex        =   83
            Tag             =   "BancoCuentaDeposito"
            Top             =   660
            Width           =   315
            _Version        =   851968
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtDepositoBanco 
            Height          =   315
            Index           =   1
            Left            =   3660
            TabIndex        =   81
            Top             =   270
            Width           =   10275
            _Version        =   851968
            _ExtentX        =   18124
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit txtDepositoBanco 
            Height          =   315
            Index           =   2
            Left            =   1200
            TabIndex        =   82
            Top             =   660
            Width           =   1995
            _Version        =   851968
            _ExtentX        =   3519
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit txtDepositoBanco 
            Height          =   315
            Index           =   3
            Left            =   3660
            TabIndex        =   84
            Top             =   660
            Width           =   10275
            _Version        =   851968
            _ExtentX        =   18124
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit txtDepositoImporte 
            Height          =   315
            Left            =   1200
            TabIndex        =   74
            Top             =   1140
            Width           =   2025
            _Version        =   851968
            _ExtentX        =   3572
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   255
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit txtDepositoComentario 
            Height          =   315
            Left            =   1500
            TabIndex        =   78
            Top             =   2880
            Width           =   12375
            _Version        =   851968
            _ExtentX        =   21828
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin Aplisoft_CajasDeTexto.TxF txtFechaDeposito 
            Height          =   315
            Left            =   1500
            TabIndex        =   76
            Top             =   2070
            Width           =   2325
            _ExtentX        =   4101
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
            BackStyle       =   0
         End
         Begin XtremeSuiteControls.FlatEdit txtnroInternoDeposito 
            Height          =   315
            Left            =   6300
            TabIndex        =   114
            Top             =   2070
            Width           =   7605
            _Version        =   851968
            _ExtentX        =   13414
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Alignment       =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtNroChequeDeposito 
            Height          =   315
            Left            =   1500
            TabIndex        =   77
            Top             =   2490
            Width           =   12375
            _Version        =   851968
            _ExtentX        =   21828
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin Aplisoft_CajasDeTexto.TxF vfechaDeposito 
            Height          =   315
            Left            =   1500
            TabIndex        =   75
            Top             =   1650
            Width           =   2325
            _ExtentX        =   4101
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
            BackStyle       =   0
         End
         Begin XtremeSuiteControls.FlatEdit vmarcainternaDeposito 
            Height          =   315
            Left            =   6270
            TabIndex        =   154
            Top             =   1560
            Width           =   1455
            _Version        =   851968
            _ExtentX        =   2566
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.PushButton PushButton7 
            Height          =   315
            Left            =   7800
            TabIndex        =   155
            Top             =   1560
            Width           =   315
            _Version        =   851968
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmCobros.frx":1969
         End
         Begin XtremeSuiteControls.PushButton PushButton8 
            Height          =   345
            Left            =   3330
            TabIndex        =   165
            Top             =   1140
            Width           =   2505
            _Version        =   851968
            _ExtentX        =   4419
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "Copiar Total Doc. Seleccionado"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Marca interna:"
            Height          =   195
            Left            =   4440
            TabIndex        =   156
            Top             =   1620
            Width           =   1695
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha operación:"
            Height          =   195
            Left            =   90
            TabIndex        =   125
            Top             =   1740
            Width           =   1245
         End
         Begin VB.Label lblDeposito 
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. Cheque:"
            Height          =   195
            Index           =   4
            Left            =   360
            TabIndex        =   122
            Top             =   2520
            Width           =   975
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Nro Interno:"
            Height          =   195
            Left            =   5280
            TabIndex        =   116
            Top             =   2130
            Width           =   915
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Fec. Acreditación:"
            Height          =   195
            Left            =   60
            TabIndex        =   115
            Top             =   2130
            Width           =   1335
         End
         Begin VB.Label lblDeposito 
            BackStyle       =   0  'Transparent
            Caption         =   "Banco/Caja:"
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   88
            Top             =   315
            Width           =   945
         End
         Begin VB.Label lblDeposito 
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta: "
            Height          =   195
            Index           =   1
            Left            =   540
            TabIndex        =   87
            Top             =   705
            Width           =   645
         End
         Begin VB.Label lblDeposito 
            BackStyle       =   0  'Transparent
            Caption         =   "Importe: "
            ForeColor       =   &H80000006&
            Height          =   195
            Index           =   2
            Left            =   510
            TabIndex        =   86
            Top             =   1185
            Width           =   645
         End
         Begin VB.Label lblDeposito 
            BackStyle       =   0  'Transparent
            Caption         =   "Comentario: "
            Height          =   195
            Index           =   3
            Left            =   450
            TabIndex        =   85
            Top             =   2925
            Width           =   915
         End
      End
      Begin VB.Frame fraTarjeta 
         Height          =   5325
         Left            =   -69940
         TabIndex        =   62
         Top             =   420
         Visible         =   0   'False
         Width           =   14025
         Begin VB.TextBox txtImporteCuponTarjeta 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1080
            TabIndex        =   67
            Top             =   1770
            Width           =   1545
         End
         Begin VB.ComboBox cboTarjeta 
            Height          =   315
            Left            =   1080
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   1230
            Width           =   3735
         End
         Begin VB.TextBox txtNroCuponTarjeta 
            Height          =   285
            Left            =   1080
            TabIndex        =   65
            Top             =   450
            Width           =   1455
         End
         Begin VB.ComboBox cboBancoTarjeta 
            Height          =   315
            ItemData        =   "frmCobros.frx":1F03
            Left            =   1080
            List            =   "frmCobros.frx":1F05
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   810
            Width           =   6855
         End
         Begin VB.TextBox txtCantCuotas 
            Height          =   285
            Left            =   3840
            TabIndex        =   63
            Top             =   1830
            Width           =   975
         End
         Begin VB.Label lblImporte 
            BackStyle       =   0  'Transparent
            Caption         =   "Importe:"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   150
            TabIndex        =   72
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label lblBanco 
            BackStyle       =   0  'Transparent
            Caption         =   "Banco:"
            Height          =   255
            Left            =   420
            TabIndex        =   71
            Top             =   840
            Width           =   615
         End
         Begin VB.Label lblTarjeta 
            BackStyle       =   0  'Transparent
            Caption         =   "Tarjeta:"
            Height          =   255
            Left            =   390
            TabIndex        =   70
            Top             =   1230
            Width           =   585
         End
         Begin VB.Label lblNroCupón 
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. cupón:"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   450
            Width           =   975
         End
         Begin VB.Label lblCantCuotas 
            BackStyle       =   0  'Transparent
            Caption         =   "Cant. cuotas:"
            Height          =   255
            Left            =   2760
            TabIndex        =   68
            Top             =   1830
            Width           =   975
         End
      End
      Begin VB.Frame fraCheques 
         Height          =   5445
         Left            =   -69940
         TabIndex        =   51
         Top             =   360
         Visible         =   0   'False
         Width           =   14085
         Begin MSDataGridLib.DataGrid dgCheques 
            Height          =   1845
            Left            =   120
            TabIndex        =   132
            Top             =   3210
            Width           =   13635
            _ExtentX        =   24051
            _ExtentY        =   3254
            _Version        =   393216
            AllowUpdate     =   -1  'True
            BackColor       =   16777215
            HeadLines       =   1
            RowHeight       =   15
            RowDividerStyle =   1
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
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
                  LCID            =   1034
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
                  LCID            =   1034
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
         Begin XtremeSuiteControls.FlatEdit vchequeComentario 
            Height          =   315
            Left            =   8160
            TabIndex        =   20
            Top             =   2790
            Width           =   5565
            _Version        =   851968
            _ExtentX        =   9816
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vmarcainterna 
            Height          =   315
            Left            =   5790
            TabIndex        =   15
            Top             =   1410
            Width           =   1455
            _Version        =   851968
            _ExtentX        =   2566
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vSucursal 
            Height          =   315
            Left            =   1170
            TabIndex        =   10
            Top             =   900
            Width           =   7425
            _Version        =   851968
            _ExtentX        =   13097
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   30
            TabIndex        =   123
            Top             =   150
            Width           =   13965
            Begin XtremeSuiteControls.PushButton PushButton1 
               Height          =   315
               Left            =   180
               TabIndex        =   124
               Top             =   0
               Width           =   13545
               _Version        =   851968
               _ExtentX        =   23892
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "Seleccionar cheque en Cartera <F3>"
               UseVisualStyle  =   -1  'True
               Picture         =   "frmCobros.frx":1F07
            End
         End
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Height          =   465
            Left            =   90
            TabIndex        =   109
            Top             =   2700
            Width           =   5955
            Begin XtremeSuiteControls.PushButton cmdAgregarCheque 
               Height          =   345
               Left            =   30
               TabIndex        =   21
               Top             =   90
               Width           =   2325
               _Version        =   851968
               _ExtentX        =   4101
               _ExtentY        =   609
               _StockProps     =   79
               Caption         =   "Agregar a lista de cheques"
               UseVisualStyle  =   -1  'True
               Picture         =   "frmCobros.frx":2919
            End
            Begin XtremeSuiteControls.PushButton cmdEliminarCheque 
               Height          =   345
               Left            =   2400
               TabIndex        =   110
               Top             =   90
               Width           =   3435
               _Version        =   851968
               _ExtentX        =   6059
               _ExtentY        =   609
               _StockProps     =   79
               Caption         =   "Eliminar el cheque seleccionado de la lista"
               UseVisualStyle  =   -1  'True
               Picture         =   "frmCobros.frx":2EB3
            End
         End
         Begin VB.TextBox txtImporteTotalCheque 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   12030
            TabIndex        =   52
            Top             =   5010
            Width           =   1695
         End
         Begin Aplisoft_CajasDeTexto.TxF dtpDepositoCheque 
            Height          =   315
            Left            =   1170
            TabIndex        =   13
            Top             =   1980
            Width           =   2025
            _ExtentX        =   3572
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
            BackStyle       =   0
         End
         Begin XtremeSuiteControls.FlatEdit txtFirmanteCheque 
            Height          =   315
            Left            =   9810
            TabIndex        =   26
            Top             =   1590
            Width           =   3915
            _Version        =   851968
            _ExtentX        =   6906
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.ComboBox cboEstadoCheque 
            Height          =   315
            Left            =   9810
            TabIndex        =   24
            Top             =   900
            Width           =   3915
            _Version        =   851968
            _ExtentX        =   6906
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "No Acreditado"
         End
         Begin XtremeSuiteControls.FlatEdit txtNroCheque 
            Height          =   315
            Left            =   1170
            TabIndex        =   11
            Top             =   1230
            Width           =   2025
            _Version        =   851968
            _ExtentX        =   3572
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit txtImporteCheque 
            Height          =   285
            Left            =   1170
            TabIndex        =   14
            Top             =   2310
            Width           =   2025
            _Version        =   851968
            _ExtentX        =   3572
            _ExtentY        =   503
            _StockProps     =   77
            ForeColor       =   255
            BackColor       =   16777215
            BackColor       =   16777215
         End
         Begin XtremeSuiteControls.FlatEdit txtNroInternoCheque 
            Height          =   315
            Left            =   9810
            TabIndex        =   25
            Top             =   1230
            Width           =   3915
            _Version        =   851968
            _ExtentX        =   6906
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Alignment       =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtBancoCheque 
            Height          =   315
            Index           =   0
            Left            =   1170
            TabIndex        =   8
            Top             =   540
            Width           =   645
            _Version        =   851968
            _ExtentX        =   1138
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.PushButton pbCarga 
            Height          =   315
            Index           =   0
            Left            =   1830
            TabIndex        =   27
            Tag             =   "Banco"
            Top             =   540
            Width           =   315
            _Version        =   851968
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtBancoCheque 
            Height          =   315
            Index           =   1
            Left            =   2190
            TabIndex        =   9
            Top             =   540
            Width           =   6405
            _Version        =   851968
            _ExtentX        =   11298
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit txtBancoCheque 
            Height          =   315
            Index           =   2
            Left            =   9810
            TabIndex        =   22
            Top             =   540
            Width           =   645
            _Version        =   851968
            _ExtentX        =   1138
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.PushButton pbCarga 
            Height          =   315
            Index           =   1
            Left            =   10440
            TabIndex        =   31
            Tag             =   "BancoCuenta"
            Top             =   540
            Width           =   315
            _Version        =   851968
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtBancoCheque 
            Height          =   315
            Index           =   3
            Left            =   10785
            TabIndex        =   23
            Top             =   540
            Width           =   2955
            _Version        =   851968
            _ExtentX        =   5212
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin Aplisoft_CajasDeTexto.TxF vfechaCheque 
            Height          =   330
            Left            =   1170
            TabIndex        =   12
            Top             =   1590
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   582
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
         Begin XtremeSuiteControls.PushButton pbCarga 
            Height          =   315
            Index           =   5
            Left            =   6930
            TabIndex        =   32
            Tag             =   "Banco"
            Top             =   1980
            Width           =   315
            _Version        =   851968
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtBancoCheque 
            Height          =   315
            Index           =   5
            Left            =   7320
            TabIndex        =   17
            Top             =   1980
            Width           =   6405
            _Version        =   851968
            _ExtentX        =   11298
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   16711680
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.PushButton bCajaDestino 
            Height          =   315
            Left            =   6900
            TabIndex        =   33
            Tag             =   "Banco"
            Top             =   2340
            Width           =   345
            _Version        =   851968
            _ExtentX        =   609
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit vDCajaDestino 
            Height          =   315
            Left            =   7320
            TabIndex        =   19
            Top             =   2340
            Width           =   6405
            _Version        =   851968
            _ExtentX        =   11298
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   255
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit txtBancoCheque 
            Height          =   315
            Index           =   4
            Left            =   5820
            TabIndex        =   16
            Top             =   1980
            Width           =   1065
            _Version        =   851968
            _ExtentX        =   1879
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   16711680
            BackColor       =   -2147483643
            Alignment       =   1
         End
         Begin XtremeSuiteControls.FlatEdit vCCajaDestino 
            Height          =   315
            Left            =   5820
            TabIndex        =   18
            Top             =   2310
            Width           =   1065
            _Version        =   851968
            _ExtentX        =   1879
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   255
            BackColor       =   -2147483643
            Alignment       =   1
         End
         Begin XtremeSuiteControls.PushButton PushButton5 
            Height          =   315
            Left            =   7320
            TabIndex        =   149
            Top             =   1410
            Width           =   615
            _Version        =   851968
            _ExtentX        =   1085
            _ExtentY        =   556
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmCobros.frx":344D
         End
         Begin VB.Label Label17 
            Caption         =   "Comentarios del cheque:"
            Height          =   195
            Left            =   6300
            TabIndex        =   151
            Top             =   2850
            Width           =   1845
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Marca interna:"
            Height          =   195
            Left            =   3960
            TabIndex        =   148
            Top             =   1470
            Width           =   1695
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sucursal:"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   147
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "Caja destino del cheque:"
            Height          =   195
            Left            =   3960
            TabIndex        =   134
            Top             =   2400
            Width           =   1845
         End
         Begin VB.Label lblIngresoDe 
            Caption         =   "Caja origen del cheque: "
            Height          =   195
            Left            =   4020
            TabIndex        =   133
            Top             =   2040
            Width           =   1695
         End
         Begin VB.Label lblCobros 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha :"
            Height          =   255
            Index           =   5
            Left            =   570
            TabIndex        =   121
            Top             =   1620
            Width           =   735
         End
         Begin VB.Label lblTotalCheque 
            BackStyle       =   0  'Transparent
            Caption         =   "Total cheque:"
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   10980
            TabIndex        =   61
            Top             =   5130
            Width           =   1095
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000013&
            Caption         =   "> Importe:"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   120
            TabIndex        =   60
            Top             =   2325
            Width           =   915
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Firmante:"
            Height          =   195
            Left            =   9120
            TabIndex        =   59
            Top             =   1620
            Width           =   735
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Acrediración:"
            Height          =   195
            Left            =   90
            TabIndex        =   58
            Top             =   1980
            Width           =   1005
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. cheque:"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   57
            Top             =   1230
            Width           =   1005
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Banco:"
            Height          =   195
            Index           =   0
            Left            =   570
            TabIndex        =   56
            Top             =   600
            Width           =   585
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Estado:"
            Height          =   195
            Left            =   9240
            TabIndex        =   55
            Top             =   930
            Width           =   645
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Nro Interno:"
            Height          =   195
            Left            =   8940
            TabIndex        =   54
            Top             =   1260
            Width           =   945
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta:"
            Height          =   195
            Index           =   1
            Left            =   9240
            TabIndex        =   53
            Top             =   600
            Width           =   585
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   165
            Left            =   60
            TabIndex        =   111
            Top             =   5040
            Width           =   12075
            _Version        =   851968
            _ExtentX        =   21299
            _ExtentY        =   291
            _StockProps     =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientColorLight=   -2147483637
            GradientColorDark=   -2147483637
         End
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   285
         Left            =   -67900
         TabIndex        =   140
         Top             =   1230
         Visible         =   0   'False
         Width           =   1575
         _Version        =   851968
         _ExtentX        =   2778
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Borrar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   285
         Left            =   -69520
         TabIndex        =   142
         Top             =   1230
         Visible         =   0   'False
         Width           =   1575
         _Version        =   851968
         _ExtentX        =   2778
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Agregar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit vtimporteRetenciones 
         Height          =   315
         Left            =   -67360
         TabIndex        =   146
         Top             =   5430
         Visible         =   0   'False
         Width           =   1635
         _Version        =   851968
         _ExtentX        =   2884
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   255
         BackColor       =   -2147483643
      End
      Begin VB.Label Label14 
         Caption         =   "Importe total de retenciones:"
         Height          =   255
         Left            =   -69520
         TabIndex        =   145
         Top             =   5490
         Visible         =   0   'False
         Width           =   2115
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   8370
      TabIndex        =   36
      Top             =   4980
      Visible         =   0   'False
      Width           =   8415
      Begin VB.TextBox txtTotal 
         Height          =   285
         Left            =   840
         TabIndex        =   49
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtPendiente 
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3960
         TabIndex        =   48
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtPagado 
         ForeColor       =   &H0000C000&
         Height          =   285
         Left            =   6600
         TabIndex        =   47
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtTipoComp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6600
         TabIndex        =   46
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtNroComprobante 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3960
         TabIndex        =   43
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblTipo 
         Caption         =   "Tipo:"
         Height          =   255
         Left            =   5880
         TabIndex        =   45
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblNroComprobante 
         Caption         =   "Nro. comprobante:"
         Height          =   255
         Left            =   2520
         TabIndex        =   44
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblParcialA 
         Caption         =   "Importe:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblPendiente 
         Caption         =   "Pendiente:"
         Height          =   255
         Left            =   2520
         TabIndex        =   38
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblPagado 
         Caption         =   "Pagado:"
         Height          =   255
         Left            =   5880
         TabIndex        =   37
         Top             =   240
         Width           =   735
      End
   End
   Begin XtremeSuiteControls.TabControl TabCobros 
      Height          =   4215
      Left            =   -8670
      TabIndex        =   29
      Top             =   -270
      Width           =   8415
      _Version        =   851968
      _ExtentX        =   14843
      _ExtentY        =   7435
      _StockProps     =   68
      Color           =   4
      ItemCount       =   3
      SelectedItem    =   1
      Item(0).Caption =   "Efectivo"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "Label1(0)"
      Item(0).Control(1)=   "txtImporteEfectivo"
      Item(1).Caption =   "Cheques"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "txtImporteTarjeta(1)"
      Item(2).Caption =   "Tarjeta"
      Item(2).ControlCount=   3
      Item(2).Control(0)=   "Pus(1)"
      Item(2).Control(1)=   "Picture1"
      Item(2).Control(2)=   "lblImporteCobrado(2)"
      Begin VB.TextBox txtImporteTarjeta 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         ForeColor       =   &H000000FF&
         Height          =   480
         HelpContextID   =   1
         HideSelection   =   0   'False
         Index           =   1
         Left            =   100
         TabIndex        =   42
         Top             =   100
         Width           =   1200
      End
      Begin VB.TextBox txtImporteEfectivo 
         Height          =   285
         Left            =   -68560
         TabIndex        =   41
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin XtremeSuiteControls.PushButton Pus 
         Height          =   315
         Index           =   1
         Left            =   0
         TabIndex        =   30
         Tag             =   "CodigoPostal"
         Top             =   1080
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin VB.PictureBox Picture1 
         Height          =   15
         Left            =   -70000
         ScaleHeight     =   15
         ScaleWidth      =   15
         TabIndex        =   34
         Top             =   0
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Label lblImporteCobrado 
         Caption         =   "Importe cobrado:"
         Height          =   255
         Index           =   2
         Left            =   -69760
         TabIndex        =   40
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Importe cobrado:"
         Height          =   255
         Index           =   0
         Left            =   -69880
         TabIndex        =   35
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin MSAdodcLib.Adodc bcheques 
      Height          =   360
      Left            =   870
      Top             =   4590
      Visible         =   0   'False
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   635
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
      Caption         =   "bcheques"
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
   Begin MSAdodcLib.Adodc bpccliente 
      Height          =   330
      Left            =   2430
      Top             =   4350
      Visible         =   0   'False
      Width           =   1965
      _ExtentX        =   3466
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
      Caption         =   "bpccliente"
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
   Begin MSAdodcLib.Adodc bbanco 
      Height          =   330
      Left            =   4710
      Top             =   4230
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "bbanco"
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
   Begin MSAdodcLib.Adodc bsucursal 
      Height          =   330
      Left            =   4590
      Top             =   4470
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "bsucursal"
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
   Begin MSAdodcLib.Adodc bccliente 
      Height          =   330
      Left            =   2550
      Top             =   4110
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
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
      Caption         =   "bccliente"
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
   Begin XtremeSuiteControls.GroupBox GroDatosCliente 
      Height          =   585
      Left            =   60
      TabIndex        =   100
      Top             =   390
      Width           =   14205
      _Version        =   851968
      _ExtentX        =   25056
      _ExtentY        =   1032
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtCliente 
         Height          =   315
         Index           =   0
         Left            =   960
         TabIndex        =   0
         Top             =   210
         Width           =   1185
         _Version        =   851968
         _ExtentX        =   2090
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtCliente 
         Height          =   345
         Index           =   1
         Left            =   3630
         TabIndex        =   28
         Top             =   240
         Width           =   5715
         _Version        =   851968
         _ExtentX        =   10081
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   2
         Left            =   750
         TabIndex        =   2
         Tag             =   "Clientes"
         Top             =   240
         Visible         =   0   'False
         Width           =   135
         _Version        =   851968
         _ExtentX        =   238
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PusBuscarCliente 
         Height          =   315
         Left            =   2220
         TabIndex        =   90
         Top             =   240
         Width           =   1275
         _Version        =   851968
         _ExtentX        =   2249
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Busca <F1>"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PusBuscarDocumento 
         Height          =   435
         Left            =   9390
         TabIndex        =   102
         Top             =   150
         Width           =   2325
         _Version        =   851968
         _ExtentX        =   4101
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Sel. documentos a pagar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton6 
         Height          =   435
         Left            =   11760
         TabIndex        =   157
         Top             =   150
         Width           =   2415
         _Version        =   851968
         _ExtentX        =   4260
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Ver movimientos de  CtaCte"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label lblNombre 
         BackStyle       =   0  'Transparent
         Caption         =   ">Nombre:"
         ForeColor       =   &H80000006&
         Height          =   165
         Left            =   90
         TabIndex        =   101
         Top             =   300
         Width           =   960
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtObservaciones 
      Height          =   330
      Left            =   1800
      TabIndex        =   6
      Top             =   7650
      Width           =   12435
      _Version        =   851968
      _ExtentX        =   21934
      _ExtentY        =   582
      _StockProps     =   77
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit vdocSeleccionado 
      Height          =   285
      Left            =   1800
      TabIndex        =   158
      Top             =   7980
      Width           =   12405
      _Version        =   851968
      _ExtentX        =   21881
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.PushButton PusSelConceptos 
      Height          =   345
      Left            =   60
      TabIndex        =   164
      Top             =   1470
      Width           =   2865
      _Version        =   851968
      _ExtentX        =   5054
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Sel. conceptos de aplicación <F4>"
      ForeColor       =   0
      BackColor       =   -2147483644
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label19 
      Height          =   225
      Left            =   60
      TabIndex        =   159
      Top             =   8010
      Width           =   1635
      _Version        =   851968
      _ExtentX        =   2884
      _ExtentY        =   397
      _StockProps     =   79
      Caption         =   "> Doc.Seleccionados:"
      Alignment       =   1
   End
   Begin XtremeSuiteControls.Label lblObservaciones 
      Height          =   285
      Left            =   60
      TabIndex        =   120
      Top             =   7680
      Width           =   1635
      _Version        =   851968
      _ExtentX        =   2884
      _ExtentY        =   503
      _StockProps     =   79
      Caption         =   "> Observaciones :"
      Alignment       =   1
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Caja:"
      Height          =   195
      Index           =   2
      Left            =   -30
      TabIndex        =   119
      Top             =   5865
      Width           =   435
   End
End
Attribute VB_Name = "frmCobros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------
Public vsql, vdocseleccionados As String
Public vtotalseleccionados As Double
Dim vsqlpago() As String
Dim vsqlpagoAuto() As String   ' para el pago automático
Dim vsqlpago2() As String
Dim vsqlpago2_temp() As String
'-------------------------------------
Dim imprimerecibo As Boolean
Dim vImportePagoPesos As Double
Dim vImpresionCorrecta As Boolean
Public vIdCtaCteC As Long
Public codCliente As String
Public pendiente, pagado, total As Double
Public remito As Long
Public esComprobanteAutomatico As Boolean
Dim rsCheque As New ADODB.Recordset
Dim rsRecibo As New ADODB.Recordset
Dim concepto As String
Public esFacturacion As Boolean
Public NroComprobante As Long
Public tipoComprobante As String
Public fechaDocumento As Date
Dim SaldoAnterior, totaldebito, totalCredito, credito, debito As Double
Dim vImporteTotalAPagar As Double
Dim CP As CobrosPagos  ' typo para setear si es un cobro o un pago
Public cpInstancia As String  ' identifica si es una instancia de Cobro o una de Pago
Public vidcheque, vnrobalance As Long
Dim vnrointerno As Long
Dim vtablaFactura As String
Dim vfecha As Date
Dim vvc As String
Dim vDraft As Boolean
Dim vnrorecibo  As Long

Dim vcuit As String

Private Enum MedioPago
    efectivoPesos = 1
    efectivoDolar = 2
    tarjeta = 3
    cheque = 4
    Deposito = 5
    NotaC = 8
    ContadoCredito = 11
    AjusteCredito = 12
End Enum
Dim vMontoAPagarParaImprimir As Double

Public Sub setvsqlPago(v() As String)
    vsqlpago = v
End Sub

Public Sub setvsqlPagoAuto(v() As String)
    vsqlpagoAuto = v
End Sub

Public Sub setvsqlPago2(v() As String)
    vsqlpago2 = v
End Sub

Public Sub setvsqlPago2_temp(v() As String)
    vsqlpago2_temp = v
End Sub

Public Sub initCobro()
    Me.vfechaCredito.SetFocus
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2 - 1000
End Sub
Private Sub init()

vfechaCredito = Date

Me.vfechaCheque.Value = Date
Me.vfechaDeposito.Value = Date

'vnroBalance = TraerDato("balances", " Activo='S' order by NroBalance Desc", "NroBalance", pathDBMySQL)
'Me.Caption = Me.Caption + "      [Nro. de Balance: " + Str(vnroBalance) + "]"

Me.TabCobros.SelectedItem = 0
TabControl1.SelectedItem = 0

Me.gretencion.TextMatrix(0, 0) = "Id"
Me.gretencion.ColWidth(0) = 1000

Me.gretencion.TextMatrix(0, 1) = "Descripción"
Me.gretencion.ColWidth(1) = 5000

Me.gretencion.TextMatrix(0, 2) = "Importe"
Me.gretencion.ColWidth(2) = 2000
Me.gretencion.ColAlignment(2) = 4

Me.gretencion.Rows = 1

Select Case cpInstancia
Case "cobro"
    Set CP.FormBuscaDocumento = frmBuscarFactura
    Set CP.FormPersona = frmClientes
    CP.TablaCtaCte = "cuentascorrientes"
    CP.TablaPersona = "clientes"
    Me.Caption = "Ing. de movimiento de cobro a un Cliente:"
    Me.txtBancoCheque(4).Enabled = True
    Me.txtBancoCheque(5).Enabled = True
    Me.pbCarga(5).Enabled = True
    vtablaFactura = "factura"
    Me.TabControl1.Item(0).Caption = "Cobro Efectivo"
    Me.TabControl1.Item(1).Caption = "Ingreso Cheques tercero"
    Me.TabControl1.Item(2).Caption = "Tarjeta"
    Me.TabControl1.Item(3).Caption = "Retenciones"
    Me.TabControl1.Item(4).Caption = "Deposito Bancario"
    
Case "pagos"
    Set CP.FormBuscaDocumento = frmBuscarCompra
    Set CP.FormPersona = frmProveedores
    CP.TablaCtaCte = "pcuentascorrientes"
    CP.TablaPersona = "proveedores"
    Me.Caption = "Ing. de movimiento de pago a un Proveedor:"
    Me.txtBancoCheque(4).Enabled = True
    Me.txtBancoCheque(5).Enabled = True
    Me.pbCarga(5).Enabled = False
    vtablaFactura = "pfactura"
  '  Me.TabCobros(1).Caption = "Egreso Cheque Caja"
    Me.TabControl1.Item(0).Caption = "Pago Efectivo"
    Me.TabControl1.Item(1).Caption = "Egreso Cheques tercero"
    Me.TabControl1.Item(2).Caption = "Tarjeta"
    Me.TabControl1.Item(3).Caption = "Retenciones"
    Me.TabControl1.Item(4).Caption = "Egreso Cheques propios"
    
End Select


Me.LimpiarCampos

txtFechaDeposito.Value = Date

End Sub

Private Sub PagarCtaCteDirecto(importe As Double, idMedioPago As Integer, vcomentario As String)
Dim sqlInsert As String

sqlInsert = "Insert Into CuentasCorrientes ( cuentascorrientes.Fecha, cuentascorrientes.Codigo, cuentascorrientes.Nombre,cuentascorrientes.debito,cuentascorrientes.Credito, cuentascorrientes.comentario)" & _
            "VALUES ('" & strfechaMySQL(dtpFecha.Value) & "', '" & Trim(codCliente) & "', '" & Me.txtCliente(1) & "',0, " & Str(importe) & ",'" & vcomentario & "')"
            'Cn.Execute sqlInsert
Call EjecutarScript(sqlInsert, pathDBMySQL)
End Sub

Private Sub bCajaDestino_Click()
Call fbuscarGrilla("Bancos", "Descripcion", "idBancos", Me.vDCajaDestino.Name, Me) ' ema:
End Sub


Function Validar() As Boolean
Validar = True

If UCase(LeerXml("Textil")) = "ADBA" Then
    If Not (InStr(Me.txtObservaciones.Text, "Documento") > 0 Or InStr(Me.txtObservaciones.Text, "Fact") > 0) Then
        MsgBox "Debe selecccionar un tipo de movimiento", vbInformation
        Validar = False
        Exit Function
    End If
End If


If Me.txtBancoCheque(6).Text = "" And Me.txtDepositoBanco(0).Text = "" Then
    Exit Function
End If

If Not cajaAbierta(Me.vfechaCredito) Then
    Validar = False
    MsgBox "La caja está cerrada", vbCritical
End If
End Function

Private Sub cmdCobrar_Click()
On Error Resume Next
'Dim vvc As String
Dim vtexto As String

vnrorecibo = getNroRecibo




If Not Validar Then Exit Sub


Me.txtObservaciones.Text = Me.txtObservaciones + Me.vdocSeleccionado


vtexto = ""

'vvc = Me.txtCliente(0).Text


Dim vTotalCredito As Double

'Me.vnroOrdenPago.Text = UltimoNroOrdenPago(Val(vnroOrdenPago.Text))

'If Val(Me.vnroOrdenPago.Text) = 0 Then Exit Sub

    
Me.txtNroInterno = UltimoNroInterno2 + 1
    
vnrointerno = Me.txtNroInterno
    
dtpFecha = vfechaCredito
    
   ' VaciarChequesTemp ' borra ls base temporal
    
    
 ' ------------ verifica nro interno ----------------------
 If existeRegistro(Val(Me.txtNroInterno)) Then Exit Sub
 
 If variosPagos Then Exit Sub
 
 '----------------------------------------------------------
    

    
    vTotalCredito = 0
    
   ' If Me.txtBancoCheque(4) = "" And Val(txtImporteTotalCheque) > 0 And Me.txtBancoCheque(4).Enabled Then
   '     If MsgBox("Desea seleccionar una caja ?", vbYesNo, "Movimiento de Caja") = vbYes Then
   '         pbCarga(5).SetFocus
   '         Exit Sub
   '     End If
   '
   '
   ' End If
    
    
    VaciarTablaRecibo

    If Not Val(TxtTotalAPagar.Text) = Val(vImporteTotalAPagar) Then
        MsgBox "Debe ingresar un importe correcto", vbCritical, "Mensaje ..."
        Exit Sub
    End If
    If Val(TxtTotalAPagar.Text) = 0 Then
        MsgBox "El monto a pagar debe ser mayor a 0", vbInformation, "Mensaje ..."
        Exit Sub
    Else
    
    
    
        '---------------------------------------------------------------------
        ' ------------- hace el pago por el total ----------------------------
        ' Call PagarCtaCteAutomaticamente(Val(TxtTotalAPagar.Text), 0)
        '---------------------------------------------------------------------
        
        
        ' fejecutaPagoEfectivo
        
        If Val(txtImporteEfectivoPesos) > 0 Then
    
            Dim importeAPagar, importePagadoPesos As Double
            importePagadoPesos = 0
            vImportePagoPesos = 0
            
            importePagadoPesos = Val(txtImporteEfectivoPesos.Text)
            'Agrego una linea en el recibo por el monto pagado en efectivo
            Call AgregarPagoRecibo(1, "Efectivo en Pesos:", importePagadoPesos)
            
           ' Call GuardarBancosMovimientos2(Str(Me.txtBancoCheque(6)), CInt(0), CDbl(Val(0)), CDbl(Val(Me.txtImporteEfectivo)), "") ', CInt(Val(0)), CLng(Val(Me.txtNroInterno)))  ' , CDate("12/12/12"), "0", "CE", "CH")
            
            
          ' guardo movimiento en la caja
          
          If Me.cpInstancia = "pagos" Then
             ' cuando e un pago que hago el movimiento va en el haber (malo)
             Call GuardarBancosMovimientos(vnrorecibo, Me.txtBancoCheque(6), CInt(Val(bancoToCuenta(Me.txtBancoCheque(6)))), CDbl(Val(0)), CDbl(Val(Me.txtImporteEfectivoPesos.Text)), Me.txtObservaciones.Text, CInt(Val(0)), CLng(Val(Me.txtNroInterno)), CDate(Me.dtpFecha.Value), "0", "", "", , Me.txtCliente(0))
          Else
             ' cuando es un cobro que hago va en el debe (bueno)
             Call GuardarBancosMovimientos(vnrorecibo, Me.txtBancoCheque(6), CInt(Val(bancoToCuenta(Me.txtBancoCheque(6)))), CDbl(Val(Me.txtImporteEfectivoPesos.Text)), CDbl(Val(0)), Me.txtObservaciones.Text, CInt(Val(0)), CLng(Val(Me.txtNroInterno)), CDate(Me.dtpFecha.Value), "0", "", "", , Me.txtCliente(0))
          End If
            
         ' Call GuardarBancosMovimientos(Me.txtBancoCheque(6), 0, 0, rsCheque.Fields("Monto"), EsNulo(rsCheque.Fields("Observaciones")), 0, Val(Me.txtNroInterno), rsCheque.Fields("FechaDeposito"), EsNulo(rsCheque.Fields("Ncheque")), "IC", "CH")

            vImportePagoPesos = importePagadoPesos

        End If
        
       ' ---------------------------------------------------------------------
       ' ---------------------- cheques --------------------------------------
       ' ---------------------------------------------------------------------

        ' fejecutaPagoCheques

        If Val(txtImporteTotalCheque.Text) > 0 Then
        
           ' If Not hayCaja Then Exit Sub
        
             'vvfecha = Me.dtpFecha
        
            Dim importePagadoCheque As Double
            importePagadoCheque = 0
            

            'Agrego una linea en el recibo por el monto pagado con cada cheque
            If Not rsCheque.RecordCount = 0 Then
                rsCheque.MoveFirst
                Do While Not rsCheque.EOF = True
                    AgregarPagoRecibo 4, "Cheque:" & Str(rsCheque.Fields("NCheque").Value) & " -Banco: " & rsCheque.Fields("Banco") & " -Suc: " & Trim(rsCheque.Fields("sucursal")) & " -Marca: " & Trim(Str(rsCheque.Fields("marcainterna"))) & " -Acred:" & Trim(Str(rsCheque.Fields("FechaAcreditacion"))), Val(rsCheque.Fields("Monto"))
                    importePagadoCheque = Val(importePagadoCheque) + Val(Format(rsCheque.Fields("Monto").Value, "#######0.00"))
                    rsCheque.MoveNext
                Loop
            End If
    
    
            ' guarda los movimientos de bancos en el módulo de caja para que lo tome como en cartera 'Ale: hacer
            'Call GuardarBancosMovimientos(Trim(txtDepositoBanco(0).Text), Val(txtDepositoBanco(2).Text), totalPagadoDeposito, 0, txtDepositoComentario.Text, 0, Val(Me.txtnroInternoDeposito.Text), Me.txtFechaDeposito)

            GuardarCheques 'Guardo los cheques en el módulo de cheques y los paso a la caja

            If esComprobanteAutomatico Then


            Else
                'Guardo en cta cte
                Dim i As Integer
                For i = 1 To KlexDetalle.Rows - 1
                    If Val(txtImporteTotalCheque.Text) <= 0 Then
                        MsgBox KlexDetalle.TextMatrix(i, 2) & " " & KlexDetalle.TextMatrix(i, 3) & " no pudo ser pagados", vbInformation, "WGestion"
                    Else
                        'Si el total a pagar es menor que lo pendiente pago ese total, sino todo lo pendiente de ese documento
                        If (KlexDetalle.TextMatrix(i, 5) > Me.txtImporteTotalCheque) Then
                            txtImporteEfectivoPesos.Text = 0
                        Else
                            txtImporteTotalCheque.Text = Val(Me.txtImporteTotalCheque.Text) - Me.KlexDetalle.TextMatrix(i, 5)
                        End If
                        
                    End If
                Next i
                
            End If
    
        End If

       ' ---------------------------------------------------------------------
       ' ---------------------- tarjetas --------------------------------------
       ' ---------------------------------------------------------------------
        
        ' fejecutaPagoTarjeta
        
        If Val(txtImporteCuponTarjeta.Text) > 0 Then
    
            Dim totalpagadoTarjeta As Double
            totalpagadoTarjeta = 0
            
            If Me.esComprobanteAutomatico Then
                totalpagadoTarjeta = totalpagadoTarjeta + Val(Me.txtImporteCuponTarjeta.Text)
            Else
                'Guardo en cta cte
                For i = 1 To Me.KlexDetalle.Rows - 1
                    If Val(Me.txtImporteCuponTarjeta.Text) <= 0 Then
                        MsgBox KlexDetalle.TextMatrix(i, 2) & " " & KlexDetalle.TextMatrix(i, 3) & " no pudo ser pagados", vbInformation, "WGestion"
                        totalpagadoTarjeta = totalpagadoTarjeta + Val(Me.txtImporteCuponTarjeta.Text)
                    Else
                        'Si el total a pagar es menor que lo pendiente pago ese total, sino todo lo pendiente de ese documento
                        If (KlexDetalle.TextMatrix(i, 5) > Val(Me.txtImporteCuponTarjeta.Text)) Then
                            txtImporteCuponTarjeta.Text = 0
                            totalpagadoTarjeta = totalpagadoTarjeta + Val(Me.txtImporteCuponTarjeta)
                        Else
                            txtImporteCuponTarjeta.Text = Val(Me.txtImporteCuponTarjeta.Text) - Me.KlexDetalle.TextMatrix(i, 5)
                            totalpagadoTarjeta = totalpagadoTarjeta + KlexDetalle.TextMatrix(i, 5)
                        End If
                        
                    End If
                Next i
                
            End If
            
            'Agrego una linea en el recibo por el monto pagado con tarjeta
            AgregarPagoRecibo 3, "Tarjeta " & Me.cboTarjeta.Text & " de Banco " & Me.cboBancoTarjeta.Text, totalpagadoTarjeta
    
            Dim idCuponTarjetaNuevo As Integer
    
            'Guardo los datos de la operacion en cupon tarjeta
            idCuponTarjetaNuevo = GuardarCuponTarjeta
    
            'Guardo un movimiento en la cuenta de bancos por el total de lo pagado con tarjeta
            Call GuardarBancosMovimientos(vnrorecibo, cboBancoTarjeta.Tag, 1, 0, totalpagadoTarjeta, Me.txtObservaciones.Text, idCuponTarjetaNuevo, Val(txtNroInterno.Text))
    
        End If


       ' ---------------------------------------------------------------------
       ' ---------------------- deposito -------------------------------------
       ' ---------------------------------------------------------------------

        ' fejecutaPagoDeposito

        If Val(txtDepositoImporte.Text) > 0 Then
    
            Dim totalPagadoDeposito As Double
            totalPagadoDeposito = 0
            'txtDepositoImporte.Text = 0
            
            If esComprobanteAutomatico Then
                totalPagadoDeposito = totalPagadoDeposito + Val(txtDepositoImporte.Text)
            Else
                'Guardo en cta cte
                For i = 1 To Me.KlexDetalle.Rows - 1
                    If Val(txtDepositoImporte.Text) <= 0 Then
                        MsgBox KlexDetalle.TextMatrix(i, 2) & " " & KlexDetalle.TextMatrix(i, 3) & " no pudo ser pagados", vbInformation, "WGestion"
                        totalPagadoDeposito = totalPagadoDeposito + Val(txtDepositoImporte.Text)
                    Else
                        'Si el total a pagar es menor que lo pendiente pago ese total, sino todo lo pendiente de ese documento
                        If (KlexDetalle.TextMatrix(i, 5) > Val(txtDepositoImporte.Text)) Then
                            totalPagadoDeposito = totalPagadoDeposito + Val(txtDepositoImporte.Text)
                        Else
                            txtDepositoImporte.Text = Val(txtDepositoImporte.Text) - Val(KlexDetalle.TextMatrix(i, 5))
                            totalPagadoDeposito = Val(totalPagadoDeposito) + Val(KlexDetalle.TextMatrix(i, 5))
                        End If
                        
                    End If
                Next i
                        
            End If
            
            'Agrego una linea en el recibo por el monto pagado con Deposito Bancario
            vtexto = "NC: " + Trim(Me.txtNroChequeDeposito.Text) + " Banco " + Trim(txtDepositoBanco(1).Text) + " Cta: " + Trim(txtDepositoBanco(1).Text) + " F.A:  " + Trim(Me.txtFechaDeposito)
            
            Call AgregarPagoRecibo(5, Trim(vtexto), totalPagadoDeposito)

            'Guardo un movimiento en la cuenta de bancos por el total de lo Depositado

            If totalPagadoDeposito = 0 Then Err = 10
        

            If Not UCase(LeerXml("Puesto")) = "CAJA" Then
                            If Me.cpInstancia = "cobro" Then
                                ' el cobro va en el deposito
                                Call GuardarBancosMovimientos(vnrorecibo, Trim(txtDepositoBanco(0).Text), Val(txtDepositoBanco(2).Text), totalPagadoDeposito, 0, txtDepositoComentario.Text, 0, vnrointerno, CDate(Me.txtFechaDeposito)) ', Me.txtNroChequeDeposito)
                
                            Else
                               ' el pago va en el haber
                                Call GuardarBancosMovimientos(vnrorecibo, Trim(txtDepositoBanco(0).Text), Val(txtDepositoBanco(2).Text), 0, totalPagadoDeposito, txtDepositoComentario.Text, 0, vnrointerno, Me.txtFechaDeposito, Me.txtNroChequeDeposito)
                    
                            End If
            End If
            DepositoACheques ' graba los datos del deposito al cheque
    
    
        End If
       
        

        retEnComentario
        
        ' -------------------------------------------------------------
        ' Ctacte
          GuardoCreditoEnCtaCte ' guarda el pago en la ctacte
        ' -------------------------------------------------------------
    


        '--------------------------------------------------------------------
   '    ' Retenciones
        GuardarRetenciones (vIdCtaCteC)
        ' -------------------------------------------------------------------
    
        
        If Err.Number > -1 Then
            MsgBox "La Cobranza se realizo exitosamente", vbInformation, "WGestion"
        Else
            MsgBox Err.Description
        End If
        
        
     
       'If vConfigGral.vImprimirReciboCliente = True Then ImprimirRecibo
        
        
        'ImprimirRecibo
        
        
        'VaciarTablaRecibo
        
        HabilitarControles (False)

    End If

    
    '---- pagar en la tabla factura y pfactura
    MarcarDocumnetosPagos
 
    MarcarDocumnetosPagosAuto
    
    '----------------------------------
    
    Call EjecutarScript("truncate temp2", pathDBMySQL)
    
    EjecutarPagos2
    
    Call llenar_pago1_temp
    
    Unload frmBuscarFactura
    '--------------------------------------
    
     If vConfigGral.vIncluyeContabilidad = True Then CargarContabilidad
     
     Call vsaldo_Click
     
    
    ImprimirRecibo
    
    
    LimpiarCampos

If Err < 0 Then
    MsgBox "Hubo errores al guardar, consulte el log del sistema", vbInformation, "WGestion"
Else
    initCobro
    'MsgBox "El pago se realizo exitosamente", vbInformation, "WGestion"
End If
 
 If vConfigGral.vIncluyeContabilidad = True Then
    frmAsientosAlta.SetFocus
    frmAsientosAlta.TabControl1.SelectedItem = 1
    
    If imprimerecibo Then drRecibo.SetFocus
    
 End If

End Sub


Function llenar_retenciones_temp(ByVal vtotal As Double) As Double
Dim vsql As String
Dim i As Integer

Dim vv, vcampos, v2 As String
'Dim vtotal As Double

vcampos = "c02,c05"


vsql = "insert into temp2 (" + vcampos + ") values ('','')"
Call EjecutarScript(vsql, pathDBMySQL)



vsql = "insert into temp2 (" + vcampos + ") values ('RETENCIONES: ','')"
Call EjecutarScript(vsql, pathDBMySQL)


vv = ""


For i = 1 To Me.gretencion.Rows - 1

        vv = vv + "'" + fc(Me.gretencion.TextMatrix(i, 0), 10) + " "
        
        vv = vv + " " + fc(Me.gretencion.TextMatrix(i, 1), 30) + "',"
        
        vv = vv + "'" + Format(Me.gretencion.TextMatrix(i, 2), "###,###,##0.00") + "'"
        
        vsql = "insert into temp2 (" + vcampos + ") values (" + vv + ")"
    
        vtotal = vtotal + Val(Me.gretencion.TextMatrix(i, 2))

        Call EjecutarScript(vsql, pathDBMySQL)
        
        vv = ""
        
Next



'        v2 = "'','-------')"

'        vsql = "insert into temp2 (" + vcampos + ") values (" + v2
    
    
'        Call EjecutarScript(vsql, pathDBMySQL)
'
'
'        v2 = "'','" + Format(vtotal, "###,###,##0.00") + "')"
'
'        vsql = "insert into temp2 (" + vcampos + ") values (" + v2
'
'
'        Call EjecutarScript(vsql, pathDBMySQL)

llenar_retenciones_temp = vtotal

End Function


Function valMD() As Boolean
Dim i As Double
Dim vmen As String

i = Val(Me.TxtTotalAPagar) - Val(Me.txtMontoTotalPendienteSeleccionado)

If i = 0 Or Val(txtMontoTotalPendienteSeleccionado = 0) Then
    valMD = True
Else

    valMD = False
    
    vmen = "Ud. ha ingresado un importe de pago diferente al total de los documentos seleccionados. " + Chr(13) + _
    "Quiere marcar a los documentos como pagos de todas manera ?"
    
    
    If MsgBox(vmen, vbYesNo) = vbYes Then
        valMD = True
    Else
        valMD = False
    End If

End If

End Function


Private Sub EjecutarPagos2()
Dim i As Integer
Dim v As String

'If Trim(vsqlpago2(1)) = "" Then Exit Sub


For i = 1 To 100
    v = vsqlpago2(i)
    If Not v = "" Then Call EjecutarScript(v, pathDBMySQL)
Next


    v = "insert into temp2 (c02,c05) values ('','')"
    If Not v = "" Then Call EjecutarScript(v, pathDBMySQL)


    v = "insert into temp2 (c02,c05) values ('','')"
    If Not v = "" Then Call EjecutarScript(v, pathDBMySQL)


    v = "insert into temp2 (c02,c05) values ('DOCUMENTOS CANCELADOS:          ','')"
    If Not v = "" Then Call EjecutarScript(v, pathDBMySQL)


    ' v = "insert into temp2 (c02,c05) values ('','')"
    'If Not v = "" Then Call EjecutarScript(v, pathDBMySQL)


For i = 1 To 100
    v = vsqlpago2_temp(i)
    If Not v = "" Then Call EjecutarScript(v, pathDBMySQL)
Next

    v = "insert into temp2 (c02,c05) values ('','')"
    If Not v = "" Then Call EjecutarScript(v, pathDBMySQL)
    
    v = "insert into temp2 (c02,c05) values ('----------------------------------------------','')"
    If Not v = "" Then Call EjecutarScript(v, pathDBMySQL)


    v = "insert into temp2 (c02,c05) values ('','')"
    If Not v = "" Then Call EjecutarScript(v, pathDBMySQL)
    
End Sub



Private Sub MarcarDocumnetosPagosAuto()
Dim i As Integer
Dim v As String

If Trim(vsqlpagoAuto(1)) = "" Then Exit Sub

If Not valMD Then Exit Sub ' valido la posibilidad

For i = 1 To 100
    v = vsqlpagoAuto(i)
    If Not v = "" Then Call EjecutarScript(v, pathDBMySQL)
Next

End Sub


Private Sub MarcarDocumnetosPagos()
On Error Resume Next
Dim i As Integer
Dim v As String

'If Trim(vsqlpago(1)) = "" Then Exit Sub

If Not valMD Then Exit Sub ' valido la posibilidad

For i = 1 To 100
    v = ""
    v = vsqlpago(i)
    If Not v = "" Then Call EjecutarScript(v, pathDBMySQL)
Next

If Err Then Exit Sub
End Sub

Private Sub verCtaCte(vc As String)

frmCtaCteC.Tag = IIf(Me.cpInstancia = "cobro", "Clientes", "Proveedores")
frmCtaCteC.init
frmCtaCteC.txtCliente = vc
Call frmCtaCteC.txtCliente_KeyUp(13, 1)
frmCtaCteC.txtCliente.Tag = vc
Call frmCtaCteC.cmdFiltroMovimientos_Click

End Sub


Private Sub GuardarRetenciones(victacte As Long)
On Error Resume Next
Dim vsql, vvalores As String
Dim i As Integer
Dim vTotalRetenciones As Double


vTotalRetenciones = 0

With Me.gretencion

    If .Rows = 1 Then Exit Sub
    
    For i = 1 To .Rows - 1
        If Val(.TextMatrix(i, 0)) > 0 Then
                        vvalores = Str$(vnrointerno) + "," + Str$(victacte) + "," + .TextMatrix(i, 0) + "," + .TextMatrix(i, 2)
                        vsql = "insert into retencionesmovimientos (nrointerno,idctacte,idretenciones,importe) values (" + vvalores + ")"
                        Call EjecutarScript(vsql, pathDBMySQL)
                        vTotalRetenciones = vTotalRetenciones + Val(.TextMatrix(i, 2))
                        AgregarPagoRecibo 5, .TextMatrix(i, 1), Val(.TextMatrix(i, 2))
                     '   Me.txtObservaciones = Me.txtObservaciones + " " + Trim(.TextMatrix(i, 1)) + " (" + .TextMatrix(i, 2) + ")"
        End If
    Next

    '-------- imprime el total de las retenciones ----------------------------
    AgregarPagoRecibo 5, "............Total Retenciones: ", vTotalRetenciones
    '-------------------------------------------------------------------------

End With

If Err Then Exit Sub
End Sub

Private Sub retEnComentario()

Dim i As Integer

With Me.gretencion

    If .Rows = 1 Then Exit Sub
    
    For i = 1 To .Rows - 1
        If Val(.TextMatrix(i, 0)) > 0 Then
               
                        Me.txtObservaciones = Me.txtObservaciones + " " + Trim(.TextMatrix(i, 1)) + " (" + Format(.TextMatrix(i, 2), "###,###,###0.00") + ")"
        End If
    Next

End With


End Sub

Function variosPagos() As Boolean
variosPagos = False
Dim i As Integer
i = 0
If Val(Me.txtImporteCheque) > 0 Then i = i + 1
If Val(Me.txtImporteEfectivo) > 0 Then i = i + 1
If Val(Me.txtDepositoImporte) > 0 Then i = i + 1

If i > 1 Then
    MsgBox "Cuidado. Usted ha ingresado varias forma de pagos-cobro al mismo tiempo", vbInformation, "Error..."
End If

End Function
Private Sub DepositoACheques()
On Error Resume Next
Dim vcampos1, vValor, vcp As String

If Me.cpInstancia = "cobro" Then
    vcp = "C"
Else
    vcp = "P"
End If


vcampos1 = "propietario,idcustodia,marcainterna,idEstadoCheque,Fecha,Codigo,Nombre,idBancos,idBancosCuentas,Ncheque,cp,FechaDeposito,Monto,NroInterno,Observaciones,TipoMovimiento,banco,bancoscuentas"
vValor = "'Propio','098'," + Str(Val(Me.vmarcainternaDeposito.Text)) + ",2,'" + strfechaMySQL(Date) + "','" + (Me.txtCliente(0).Text) + "','" + Me.txtCliente(1).Text + "','" + (Me.txtDepositoBanco(0)) + "','" + (Str(Val(Me.txtDepositoBanco(2)))) + "','" + (Me.txtNroChequeDeposito) + "','" + vcp + "','" + strfechaMySQL(Me.txtFechaDeposito.Value) + "'," + Me.txtDepositoImporte + "," + Str(vnrointerno) + ",'" + Me.txtDepositoComentario + "','" + "CH" + "','" + Me.txtDepositoBanco(1).Text + "','" + Me.txtDepositoBanco(3).Text + "'"

 grabarCheque vcampos1, vValor ' graba el cheque en el modulo de cheque con el seguimiento correspondiente
 
 setMarcaInterna (Str(Val(Me.vmarcainternaDeposito.Text)))
 
If Err Then
    Call GrabarLog("DepositoACheque", Err.Description, Me.Name)
    Exit Sub
End If
End Sub
Function hayCaja() As Boolean

If Trim$(txtBancoCheque(4).Text) = "" Then
    MsgBox "Debe seleccionar una Caja para los cheques de terceros ingresados", vbCritical, "Cheques de tercero..."
    hayCaja = False
Else
    hayCaja = True
End If

End Function


Private Function GuardoCreditoEnCtaCte() As Double
On Error Resume Next

    Dim vTotalCredito As Double
    Dim i As Integer
    
    vTotalCredito = 0
    
    vTotalCredito = GenerarDato("SELECT SUM(Monto) as TotalRecibo FROM Recibo_Temp", "TotalRecibo")
    
    Dim rsCredito As New ADODB.Recordset, sqlCredito As String
    
    
    
    For i = 1 To 10
         sqlCredito = "SELECT * FROM " + CP.TablaCtaCte + " WHERE 1=2"
    Next
    
    
    With rsCredito
        Call .Open(sqlCredito, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        .AddNew
        
        .Fields("Fecha").Value = strfechaMySQL(vfechaCredito)
        .Fields("Codigo").Value = Trim(txtCliente(0).Text)
        .Fields("Nombre").Value = Trim(txtCliente(1).Text)
        
        
       '  If Me.cpInstancia = "cobro" Then
            .Fields("Credito").Value = Val(Me.TxtTotalAPagar.Text) 'Val(txtDepositoImporte.Text) + Val(Me.txtImporteEfectivo.Text)   ' panic!! esto está mal
            .Fields("Debito").Value = 0
       ' Else
       '     .Fields("Debito").Value = Val(Me.TxtTotalAPagar.Text)    'Val(txtDepositoImporte.Text) + Val(Me.txtImporteEfectivo.Text)   ' panic!! esto está mal
       '     .Fields("Credito").Value = 0
       ' End If
        
        
        
        .Fields("Comentario").Value = Left(Trim(txtObservaciones.Text), 250)
        .Fields("Remito").Value = Null
        .Fields("NroInterno").Value = Val(txtNroInterno.Text)
        .Fields("idMedioPago").Value = 99
        .Fields("NroAsiento").Value = Null
        .Fields("TipoMovimiento").Value = "RC"
        
        .Update
        
        vIdCtaCteC = .Fields(0).Value
    
    End With

    GuardoCreditoEnCtaCte = vTotalCredito

    sqlCredito = ""
    rsCredito.Close
    If rsCredito.State = 1 Then
        rsCredito.Close
        Set rsCredito = Nothing
    End If
    
If Err < 0 Then
    MsgBox "Cuidado. Ocurrió un error al intentar grabar el movimiento de crédito", vbCritical, "Error..."
    GrabarLog "GuardoCreditoEnCtaCte", Err.Number & " " & Err.Description + "Panic!!!", Me.Caption
End If

End Function

Private Sub cmdRecibo2_Click()

'Call llenar_pago1

Call llenar_pago1_temp

End Sub

Private Sub Command1_Click()
llenar_pago1
End Sub





Private Sub llenar_pago_temp()
        

With drRecibo11
    
    .Sections("titulos").Controls("enro").Caption = Format(Me.vnroOrdenPago, "000000")

    .Sections("titulos").Controls("efecha").Caption = Str(Me.vfechaCredito)
    
    .Sections("titulos").Controls("enombre").Caption = Trim$(Me.txtCliente(1))
    
    .Sections("titulos").Controls("edomicilio").Caption = ""

    .Sections("titulos").Controls("ecuit").Caption = vcuit
    
   ' .Sections("pie").Controls("enletras").Caption = EnLetras(Val(Me.TxtTotalAPagar.Text))
    
    .Sections("titulos").Controls("eobservacion").Caption = txtObservaciones.Text
    
    .Sections("titulos").Controls("ett").Caption = ""
    
End With

        
    
End Sub





Private Sub llenar_pago()
    Dim vsql As String
    
    Dim vvalues As String
    
    vvalues = "99, '10-10-2016', 'nombre1', 'comentario1'"
    
    vsql = " INSERT INTO pago ( nroorden, fecha, nombre, comentario ) " + _
    " VALUES (" + vvalues + ")"
    
    Call EjecutarScript(vsql, pathDBMySQL)

End Sub

Private Sub Command2_Click()
'llenar_pago

drRecibo11.Show

End Sub

Private Sub dtpFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  '  txtNroInterno.SetFocus
End If
End Sub

Private Sub dtpFecha_LostFocus()
vfechaCredito = dtpFecha
End Sub



Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = vbKeyF1 Then
   Call PusBuscarCliente_Click
End If

If KeyCode = vbKeyF2 Then
    Call cmdCobrar_Click
End If

If KeyCode = 13 Then
    SendKeys "{tab}"
    'KeyCode = 13
End If

If KeyCode = vbKeyF3 Then
   Call PushButton1_Click
End If


If Err Then GrabarLog "Form_KeyUp", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub Frame2_DblClick()
ImprimirRecibo
End Sub

Private Sub pbCarga_Click(Index As Integer)
On Error Resume Next
'Call fbuscarGrilla("Bancos", "Descripcion", "idBancos", Me.vbc1.Name, Me)      ' ema:
'Exit Sub
    vVuelveBusqueda = Me.Name
    
    If Index = 5 Then
        vVieneBusqueda = "Caja"
    Else
        vVieneBusqueda = pbCarga(Index).Tag
    End If
    
    frmBusqueda.Show

If Err Then GrabarLog "pbCarga_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub Resizer_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub

Private Sub PusCopiarTotal_Click()
Me.txtImporteEfectivoPesos.Text = txtMontoTotalPendienteSeleccionado
End Sub

Private Sub PushButton1_Click()

Call chequesEnCartera(Me.cpInstancia, Me.txtCliente(0).Text, Me.txtCliente(1).Text)

Exit Sub

If cpInstancia = "cobro" Then Exit Sub ' si es un cobro no puede seleccionar cheques

frmCheques.ComListado.Text = "En Cartera"
frmCheques.CombOrdenamiento.Text = "marcainterna"

Call frmCheques.PBFiltrar_Click
frmCheques.Show
frmCheques.WindowState = vmaximizar
frmCheques.vViene = Me.cpInstancia

gbldsCheques.Codigo = Me.txtCliente(0).Text
gbldsCheques.Nombre = Me.txtCliente(1).Text

frmCheques.vbusca.SetFocus

'frmcheques.seleccionar(cpinstancia,"cartera")

End Sub

Private Sub PushButton2_Click()
Call fbuscarGrilla("retenciones", "descrip", "idretencion", Me.vdretencion.Name, Me) ' ema:
End Sub

Private Sub PushButton3_Click()
On Error Resume Next
Dim vline As String

If Me.vCretencion = "" Or Me.vdretencion = "" Or Me.vImporteRet = "" Then
    MsgBox "Debe seleccionar alguna retención válida", vbCritical
    Exit Sub
End If

vline = Me.vCretencion + vbTab + Me.vdretencion + vbTab + Str$(Me.vImporteRet)
Me.gretencion.AddItem vline

limpiarRetenciones

calTotalRetenciones

If Err Then Exit Sub
End Sub

Private Sub calTotalRetenciones()
Dim i As Integer
Dim vtotal As Double

With Me.gretencion
    For i = 1 To .Rows - 1
                
        vtotal = vtotal + Val(.TextMatrix(i, 2))
    
    Next
End With

Me.vtimporteRetenciones.Text = vtotal

End Sub
Private Sub limpiarRetenciones()

Me.vCretencion.Text = ""
Me.vdretencion.Tag = ""
Me.vdretencion.Text = ""
Me.vImporteRet.Text = ""

End Sub

Private Sub PushButton4_Click()
On Error Resume Next
Me.gretencion.RemoveItem (Me.gretencion.RowSel)
calTotalRetenciones
If Err Then Exit Sub
End Sub

Private Sub PushButton5_Click()
Me.vmarcaInterna = getMarcaIntarna
End Sub

Private Sub PushButton6_Click()
vvc = Me.txtCliente(0)
verCtaCte (vvc)
End Sub

Private Sub PushButton7_Click()
On Error Resume Next
    Me.vmarcainternaDeposito = getMarcaIntarna
If Err Then Exit Sub
End Sub

Private Sub PushButton8_Click()
Me.txtDepositoImporte.Text = txtMontoTotalPendienteSeleccionado
End Sub

Private Sub PusImprimirComprobante_Click()
    vDraft = True
    llenarDrRecibo
    drRecibo.Show
End Sub

Private Sub ocultarEtiquetas()
drRecibo.Sections("TituloEmpresa").Controls("Etiqueta7").Visible = False
drRecibo.Sections("Sección5").Controls("Etiqueta15").Visible = False
End Sub

Private Sub PusSelConceptos_Click()
Call fbuscarGrilla("conceptos2", "descripcion", "idconceptos", Me.vconcepto.Name, Me)   ' ema:
End Sub

Private Sub radioTodoDoc_Click()

If Not InStr(Me.txtObservaciones, "Documento") Then
    Me.txtObservaciones.Text = Me.txtObservaciones.Text + " Documento"
End If

End Sub

Private Sub Radsolofact_Click()

If Not InStr(Me.txtObservaciones, "Fact") Then
    Me.txtObservaciones.Text = Me.txtObservaciones.Text + " Fact"
End If
End Sub

Private Sub TabControl1_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error Resume Next
If Item.Index = 1 Then Me.txtBancoCheque(0).SetFocus
If Item.Index = 0 Then txtImporteEfectivoPesos.SetFocus

If Err Then Exit Sub
End Sub

Public Sub txtBancoCheque_Change(Index As Integer)
Dim vsql As String


If Not txtBancoCheque(6) = "" And LeerXml("Perfil") = "CargaCompras" Then
    MsgBox "No tiene permisos para imputar movimientos en caja"
    txtBancoCheque(6) = ""
End If



'If Index = 1 Then
'Me.txtBancoCheque(0).Text = Me.txtBancoCheque(1).Tag'
'End If
If Index = 0 Or Index = 2 Or Index = 4 Then

    vsql = "select * from bancos where idBancos='" + Trim(txtBancoCheque(Index)) + "'"
    txtBancoCheque(Index).Tag = traerDatos2(Trim(vsql), Trim("idBancos"), pathDBMySQL)

End If
End Sub

Private Sub txtBancoCheque_GotFocus(Index As Integer)
On Error Resume Next

   ' Resizer.VScrollPosition = 840
    'Call CargarCombo("Bancos", "Descripcion", txtBancoCheque, False) ', Str(idBancos))

If Err Then GrabarLog "cboBanco_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub

Public Sub txtBancoCheque_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
    
        Select Case Index
        
            Case 0
               ' txtBancoCheque(Index + 1).Text = TraerDato("Bancos", "idBancos = '" & Trim(txtBancoCheque(Index).Text) & "'", "Descripcion")
                
                 txtBancoCheque(Index + 1).Text = TraerDato("Bancos", "right(concat ('000', trim(idbancos)),3) = '" & Trim(txtBancoCheque(Index).Text) & "' and length(trim(idbancos)) < 4 and  escaja = 'N'", "Descripcion")
                txtBancoCheque(Index).Text = TraerDato("Bancos", "right(concat ('000', trim(idbancos)),3) = '" & Trim(txtBancoCheque(Index).Text) & "' and length(trim(idbancos)) < 4 and  escaja = 'N'", "idbancos")
                
                
                'txtBancoCheque(Index + 2).SetFocus
            
            Case 2
                txtBancoCheque(Index + 1).Text = TraerDato("BancosCuentas", "idBancosCuentas = " & Trim(txtBancoCheque(Index).Text) & "", "Cuenta")
                'txtNroCheque.SetFocus
        Case 4
                txtBancoCheque(Index + 1).Text = TraerDato("Bancos", "idBancos = '" & Trim(txtBancoCheque(Index).Text) & "'", "Descripcion")
                txtNroCheque.SetFocus
        
          'txtBancoCheque(Index + 1).Text = TraerDato("Bancos", "right(concat ('00000', trim(idbancos)),5) = '" & Trim(txtBancoCheque(Index).Text) & "' and escaja = 'N'", "Descripcion")
          'txtBancoCheque(Index).Text = TraerDato("Bancos", "right(concat ('00000', trim(idbancos)),5) = '" & Trim(txtBancoCheque(Index).Text) & "' and escaja = 'N'", "idbancos")
        
        End Select
    
    
       ' If txtBancoCheque(Index).Text = "" Then txtNroCheque.SetFocus
    
    End If

If Err Then GrabarLog "cboBanco_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cboBancoTarjeta_Click()
On Error Resume Next

    Call CargarComboTarjetaPorBanco("Tarjeta", "Nombre", cboTarjeta, False, "Nombre", cboBancoTarjeta.Text)
    cboTarjeta.SetFocus

If Err Then GrabarLog "cboBancoTarjeta_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cboBancoTarjeta_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
      '  Me.cboTarjeta.SetFocus
    End If
        
End Sub
Private Sub cboEstadoCheque_GotFocus()
On Error Resume Next

    Call CargarComboNew("EstadoCheque", "Descripcion", cboEstadoCheque, False) ', Str(idEstadoCheque))

If Err Then GrabarLog "cboEstadoCheque_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub cboEstadoCheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      '  vfechaCheque.SetFocus
    End If
End Sub

Private Sub cboTarjeta_GotFocus()
On Error Resume Next

    Call CargarComboTarjetaPorBanco("Tarjeta", "Nombre", cboTarjeta, False, "idBancos", Me.cboBancoTarjeta.Tag) ', Str(idTarjeta))

If Err Then GrabarLog "cboTarjeta_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub cboTarjeta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   ' Me.txtImporteCuponTarjeta.SetFocus
End If
End Sub
Private Sub dgCheques_AfterDelete()
    'rsCheque.Delete
End Sub
Private Sub dgCheques_BeforeDelete(Cancel As Integer)
    Me.txtImporteTotalCheque.Text = Val(Me.txtImporteTotalCheque.Text) - rsCheque.Fields("monto").Value
End Sub
Private Sub dgCheques_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        EliminarFilaCheque
    End If
End Sub
Private Sub dtpDepositoCheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    '    Me.txtFirmanteCheque.SetFocus
    End If
End Sub
Private Sub Form_Load()
On Error Resume Next


    Call init     ' inicializa si es un cobro o un pago

    txtTotal.Text = Val(Format(total, "########0.00"))
    
    'Resizer.HScrollPosition = 0
    'Resizer.VScrollPosition = 0
    
    Call CargarCombo("Bancos", "Descripcion", cboBancoTarjeta, False) ', Str(idBancos))
    
    txtCotizacionDolar.Text = ObtenerCotizacionMoneda("002", True)
    
    LimpiarCampos
    
    HabilitarControles (False)
    
    FormatoGrillaDetalle (1)

    'Me.Height = 9135
    'Me.Width = 14325

    Call CentrarFormulario(Me)
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub VaciarChequesTemp()
On Error Resume Next

    Dim sqlCheque As String
    
    With rsCheque
         sqlCheque = "DELETE FROM cheques_temp"
         If .State = 0 Then
            .CursorLocation = adUseClient
            .Open sqlCheque, ConnDDBB, adOpenDynamic, adLockPessimistic
        Else
            Set rsCheque = ConnDDBB.Execute(sqlCheque)
        End If
        
        Set dgCheques.DataSource = rsCheque
    End With

If Err Then GrabarLog "VaciarChequesTemp", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub actualizarRSCheque()
Dim sqlCheque As String

With rsCheque
         sqlCheque = "select *  FROM cheques_temp"
         If .State = 0 Then
            .CursorLocation = adUseClient
            .Open sqlCheque, ConnDDBB, adOpenDynamic, adLockPessimistic
        Else
            Set rsCheque = ConnDDBB.Execute(sqlCheque)
        End If
        
        'Set dgCheques.DataSource = rsCheque
End With
End Sub


Private Sub GuardarCheques()
On Error Resume Next

    Dim rsChequeGuardar As New ADODB.Recordset, sqlCheque As String
    Dim i As Integer
    Dim vidBancoCaja As Integer
    Dim vid, vIdCheques As Long
    Dim vsql, vsql2 As String
    

  '  rsCheque.Requery
  
  Call actualizarRSCheque
  

    rsCheque.MoveFirst
    If Not rsCheque.EOF Then
    
        sqlCheque = "SELECT * FROM cheques"
    
        With rsChequeGuardar
            .CursorLocation = adUseClient
            Call .Open(sqlCheque, ConnDDBB, adOpenStatic, adLockPessimistic)
           
            
            Do Until rsCheque.EOF = True  ' recorro la base temporal
            
                
                'Call GuardarBancosMovimientos(Trim(txtDepositoBanco(0).Text), Val(txtDepositoBanco(2).Text), totalPagadoDeposito, 0, txtDepositoComentario.Text, 0, Val(Me.txtnroInternoDeposito.Text), Me.txtFechaDeposito)

                
               '---------- Guardo el movimiento en los módulos de Banco -----------------------------
                If Me.cpInstancia = "cobro" Then 'Alfredo: Ale: esto  quedo mal resuelto
                     ' si es un cobro, entonces paso a los pagos de los cheues a la caja como debitos (algo BUENO para la caja)
                     
                     
                 '   vnrointerno = UltimoNroInterno2 + 1
                     
                                            
                                             ' ------------ guarda el movimiento en el modulo de cheques --------
                                            .AddNew
                                            For i = 1 To rsCheque.Fields.Count - 3
                                                Debug.Print EsNulo(rsCheque.Fields(i).Value)
                                                .Fields(i).Value = EsNulo(rsCheque.Fields(i).Value)
                                            Next i
                                            .Fields("cp").Value = "C"
                                            .Fields("TipoMovimiento").Value = "RC"
                                            .Fields("idbancocaja").Value = vidBancoCaja
                                            .Fields("propietario").Value = "Cartera"
                                            '.Fields("idCustodia").Value = EsNulo(rsCheque.Fields("idCustodia").Value)
                                             .Fields("idCustodia").Value = EsNulo(rsCheque.Fields("idCajaDestino").Value)
                                            .Fields("NroInterno").Value = vnrointerno
                                            .Fields("marcainterna").Value = EsNulo(rsCheque.Fields("marcainterna").Value)
                                            .Fields("sucursal").Value = EsNulo(rsCheque.Fields("sucursal").Value)
                                            .Fields("Firmante").Value = EsNulo(rsCheque.Fields("Firmante").Value)
                                            .Update
                                            ' ------------------------------------------------------------------
                
                setMarcaInterna (EsNulo(rsCheque.Fields("marcainterna").Value))
                
                'vidBancoCaja = GenerarDato("SELECT MAX(bancosmovimientos.idBancosMovimientos) As IDBANCOCAJA From  bancosmovimientos", "IDBANCOCAJA", pathDBMySQL)
                
                vIdCheques = GenerarDato("SELECT MAX(idcheques) As id From  cheques", "id", pathDBMySQL)
                
                ' entra a una la caja destino
                'Call GuardarBancosMovimientos(EsNulo(rsCheque.Fields("idCajaDestino").Value), 0, rsCheque.Fields("Monto"), 0, EsNulo(Me.txtObservaciones.Text), 0, vnrointerno, rsCheque.Fields("FechaDeposito"), EsNulo(rsCheque.Fields("Ncheque")), "IC", "CH", vIdCheques, Me.txtCliente(0).Text)
                
                
                If Not EsNulo(rsCheque.Fields("idCajaOrigen").Value) = "" Then
                    Call GuardarBancosMovimientos(vnrorecibo, EsNulo(rsCheque.Fields("idCajaOrigen").Value), 0, 0, rsCheque.Fields("Monto"), EsNulo(Me.txtObservaciones.Text), 0, vnrointerno, vfechaCredito.Value, EsNulo(rsCheque.Fields("Ncheque")), "IC", "CH", vIdCheques, Me.txtCliente(0).Text)
                End If
                
                If Not EsNulo(rsCheque.Fields("idCajaDestino").Value) = "" Then
                    Call GuardarBancosMovimientos(vnrorecibo, EsNulo(rsCheque.Fields("idCajaDestino").Value), 0, rsCheque.Fields("Monto"), 0, EsNulo(Me.txtObservaciones.Text), 0, vnrointerno, vfechaCredito.Value, EsNulo(rsCheque.Fields("Ncheque")), "IC", "CH", vIdCheques, Me.txtCliente(0).Text)
                End If
                
                
                
                
                
               'Call GuardarBancosMovimientos(Me.txtBancoCheque(4), 0, rsCheque.Fields("Monto"), 0, EsNulo(Me.txtObservaciones.Text), 0, vnrointerno, rsCheque.Fields("FechaDeposito"), EsNulo(rsCheque.Fields("Ncheque")), "IC", "CH", vIdCheques, Me.txtCliente(0).Text)
           
                
                Else
                    ' si es un pago, entonces paso a los pagos de los cheues a la caja como credito (algo MALO para la caja)
                     
  '                   vnrointerno = UltimoNroInterno2 + 1
                     
                   '  Call GuardarBancosMovimientos(Me.txtBancoCheque(4), 0, 0, rsCheque.Fields("Monto"), EsNulo(Me.txtObservaciones.Text), 0, vnrointerno, rsCheque.Fields("FechaDeposito"), EsNulo(rsCheque.Fields("Ncheque")), "IC", "CH")
                
                
                                          '  vidBancoCaja = GenerarDato("SELECT MAX(bancosmovimientos.idBancosMovimientos) As IDBANCOCAJA From  bancosmovimientos", "IDBANCOCAJA", pathDBMySQL)
                                            
                                            vsql = "select * from cheques where idCheques=" + Str(rsCheque.Fields("idCheques"))
                                            
                                            vid = 0
                                            vid = Val(traerDatos2(vsql, "idCheques", pathDBMySQL))
                                            
                                            If vid > 0 Then
                                                ' el cheque estaba en carterla y le tengo que cambiar el estado
                                                vsql2 = "update  cheques set idcustodia='" + EsNulo(rsCheque.Fields("idCustodia").Value) + "', endoso='(" + Trim(Me.txtCliente(0)) + ") " + Trim(Me.txtCliente(1).Text) + "', comentarios='" + EsNulo(rsCheque.Fields("comentarios").Value) + "', firmante = '" + Trim(rsCheque.Fields("firmante").Value) + "' where idCheques=" + Str(vid)
                                                Call EjecutarScript(vsql2)
                                                
                                                vIdCheques = vid
                                                
                                            Else
                                             
                                                'si el cheque con el que acabo de pagar no estaba en la cartera, lo ingreso como entregado
                                                
                                             ' ------------ guarda el movimiento en el modulo de cheques --------
                                                .AddNew
                                                For i = 1 To rsCheque.Fields.Count - 3
                                                    .Fields(i).Value = EsNulo(rsCheque.Fields(i).Value)
                                                Next i
                                                .Fields("cp").Value = "P"
                                                .Fields("TipoMovimiento").Value = "RC"
                                                .Fields("idbancocaja").Value = vidBancoCaja
                                                .Fields("propietario").Value = Me.codCliente  ' le pone el nuevo propietario del cheque
                                                .Fields("idCustodia").Value = EsNulo(rsCheque.Fields("idCustodia").Value)
                                                .Fields("marcainterna").Value = EsNulo(rsCheque.Fields("marcainterna").Value)
                                                .Fields("sucursal").Value = EsNulo(rsCheque.Fields("sucursal").Value)
                                                .Fields("firmante").Value = EsNulo(rsCheque.Fields("firmante").Value)
                                                
                                                .Update
                                                ' ------------------------------------------------------------------
                                                setMarcaInterna (EsNulo(rsCheque.Fields("marcainterna").Value))
                                           
                                                vIdCheques = GenerarDato("SELECT MAX(idcheques) As id From  cheques", "id", pathDBMySQL)
                                           
                                           End If
                
                
                If Not EsNulo(rsCheque.Fields("idCajaOrigen").Value) = "" Then
                    Call GuardarBancosMovimientos(vnrorecibo, EsNulo(rsCheque.Fields("idCajaOrigen").Value), 0, 0, rsCheque.Fields("Monto"), EsNulo(Me.txtObservaciones.Text), 0, vnrointerno, vfechaCredito.Value, EsNulo(rsCheque.Fields("Ncheque")), "IC", "CH", vIdCheques, Me.txtCliente(0).Text)
                End If
                
                If Not EsNulo(rsCheque.Fields("idCajaDestino").Value) = "" Then
                    Call GuardarBancosMovimientos(vnrorecibo, EsNulo(rsCheque.Fields("idCajaDestino").Value), 0, rsCheque.Fields("Monto"), 0, EsNulo(Me.txtObservaciones.Text), 0, vnrointerno, vfechaCredito.Value, EsNulo(rsCheque.Fields("Ncheque")), "IC", "CH", vIdCheques, Me.txtCliente(0).Text)
                End If
               
               
                End If
                '-------------------------------------------------------------------------------------
                           
                
               
                
                rsCheque.MoveNext
            Loop

        End With
        
        sqlCheque = ""

        If rsChequeGuardar.State = 1 Then
            rsChequeGuardar.Close
            Set rsChequeGuardar = Nothing
        End If
        
    End If

If Err Then GrabarLog "GuardarCheques", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    VaciarChequesTemp

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub cmdAgregarCheque_Click()
    On Error Resume Next
    
    Dim vIDCaja As Integer

    If (Val(remito) = 0) And (Trim(codCliente) = "") Then
        MsgBox "Debe seleccionar un Cliente o un remito antes de iniciar la operacion", vbOKOnly, "Mensaje ..."
        Exit Sub
    End If
    
    
    If ValidarIngresoCheque = True Then
            Dim sqlCheque As String
          
            With rsCheque
                sqlCheque = "SELECT * FROM cheques_temp where ncheque = '" + Left(Trim(txtNroCheque.Text), 20) + "'"
                
                If .State = 0 Then
                    .CursorLocation = adUseClient
                    Call .Open(sqlCheque, ConnDDBB, adOpenStatic, adLockPessimistic)
                    If Not .State = 1 Then
                        MsgBox "No Pudo abrirse la DDBB", vbExclamation, "Mensaje ..."
                        Exit Sub
                    End If
                    
                End If
                
                
              '  If .RecordCount > 0 Then
              '      MsgBox "Este cheque fue cargado", vbCritical
              '      Exit Sub
              '  End If
                
                .AddNew
                .Fields("idCheques").Value = Me.vidcheque ' guardo el id de la tabla cheque
                .Fields("idEstadoCheque").Value = Val(cboEstadoCheque.Tag)
                .Fields("Codigo").Value = EsNulo(txtCliente(0).Text)
                .Fields("Nombre").Value = EsNulo(txtCliente(1).Text)
                '.Fields("Fecha").Value = strfechaMySQL(dtpFecha.Value)
                .Fields("Fecha").Value = strfechaMySQL(dtpFecha.Value)
                .Fields("FechaDeposito").Value = strfechaMySQL(dtpDepositoCheque.Value)
                .Fields("FechaAcreditacion").Value = strfechaMySQL(dtpDepositoCheque.Value)
                
                .Fields("Monto").Value = Val(txtImporteCheque.Text)
                
                .Fields("Banco").Value = EsNulo(txtBancoCheque(1).Text)
                .Fields("idbancos").Value = EsNulo(txtBancoCheque(0).Tag)
                
                .Fields("bancoscuentas").Value = EsNulo(txtBancoCheque(3).Text)
                .Fields("idbancoscuentas").Value = EsNulo(txtBancoCheque(2).Tag)
                
                
                ' --- caja origen y destino ------------
                .Fields("idCajaOrigen").Value = EsNulo(txtBancoCheque(4).Tag)
                .Fields("idCajaDestino").Value = EsNulo(Me.vCCajaDestino)
                '---------------------------------------
                
               '------- custodia del cheque ----------------
                If Me.cpInstancia = "cobro" Then
                
                            '.Fields("idCustodia").Value = EsNulo(txtBancoCheque(4).Tag)
                            .Fields("idCustodia").Value = Me.vCCajaDestino
                
               Else
                             .Fields("idCajaDestino").Value = traerDatos2("select * from bancos where idbancos = '" + vCodigoChequesEntregados + "'", "idbancos", pathDBMySQL)
                             .Fields("idCustodia").Value = .Fields("idCajaDestino").Value 'es la caja  que indica que está entregado
               End If
               '------------------------------------------
               
               
               
                .Fields("NCheque").Value = Left(Trim(txtNroCheque.Text), 20)
            
                .Fields("Remito").Value = remito
                .Fields("Firmante").Value = Trim(txtFirmanteCheque.Text)
                .Fields("NroInterno").Value = Val(txtNroInternoCheque.Text)
                .Fields("sucursal").Value = Me.vsucursal
                .Fields("marcainterna").Value = Me.vmarcaInterna
                .Fields("comentarios").Value = Trim(Me.vchequeComentario)
                
               '.Fields("endoso").Value = Me.vf
                
                
                .Update
                
                txtImporteTotalCheque.Text = GenerarDato("SELECT SUM(Monto) as TotalCheques FROM Cheques_Temp", "TotalCheques")
            
                FormatoGrillaCheques
                
                Me.vmarcaInterna = Val(Me.vmarcaInterna) + 1
                
                LimpiarCheques
                
                txtBancoCheque(0).SetFocus
    
            End With
    
    End If
        
    
    
If Err Then GrabarLog "cmdAgregarCheque_Click", Left(Err.Number & " " & Err.Description, 99), Me.Name
End Sub
Private Sub FormatoGrillaCheques()
On Error Resume Next

    'Set Me.ms.Recordset = rsCheque
    With dgCheques
        Set .DataSource = rsCheque
        
        .Columns(0).Width = 0
        .Columns(1).Width = 0
        .Columns(2).Width = 1000
        .Columns(3).Width = 0
        .Columns(4).Width = 0
        .Columns(5).Width = 1000
        .Columns(6).Width = 1000
        
        .Columns(7).Width = 750
        
        .Columns(8).Width = 750
        .Columns(9).Width = 750
        .Columns(10).Width = 750
        .Columns(11).Width = 750
        .Columns(12).Width = 0
        .Columns(13).Width = 0
        .Columns(14).Width = 0
    
        .Columns(15).Width = 0
        .Columns(16).Width = 0
        .Columns(17).Width = 0
        .Columns(18).Width = 0

    End With
    
If Err Then GrabarLog "FormatoGrillaCheques", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub LimpiarCheques()
On Error Resume Next

    txtBancoCheque(0).Text = ""
    txtBancoCheque(1).Text = ""
    txtBancoCheque(2).Text = ""
    txtBancoCheque(3).Text = ""
    txtNroCheque.Text = ""
    txtFirmanteCheque.Text = ""
    dtpDepositoCheque.Value = Date
    cboEstadoCheque.Text = "No Acreditado"
    cboEstadoCheque.Tag = 2
    txtImporteCheque.Text = ""
    txtNroCheque.Text = ""
    
    Me.vidcheque = 0
    
    Me.txtBancoCheque(4).Text = ""
    Me.txtBancoCheque(5).Text = ""
    
    Me.vsucursal.Text = ""
    
   
   ' Me.vDCajaDestino.Text = ""
   ' Me.vDCajaDestino.Tag = ""
   ' Me.vCCajaDestino.Text = ""
   
   Call limpiarControles(75, 80)
    
If Err Then GrabarLog "LimpiarCheques", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cmdAgregarCheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAgregarCheque_Click
    End If
End Sub

Private Sub PusBuscarCliente_Click()
On Error Resume Next

    With CP.FormPersona ' tomo la persona según si es un cliente o un proveedor
        
        .Show
        If Me.cpInstancia = "cobro" Then 'Alfredo: Ale: esto  quedo mal resuelto
            .vieneCobro = True
        Else
            .vienePago = True
        End If
        .txtBuscar.SetFocus
    End With
    
    HabilitarControles (True)

    

If Err Then GrabarLog "PusBuscarCliente_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub PusBuscarDocumento_Click()
On Error Resume Next

    With frmBuscarFactura
            
            .vImporteSeleccionado.Tag = Me.TxtTotalAPagar.Tag
            .vImporteSeleccionado.Caption = Me.TxtTotalAPagar
           .Show
        If Trim(codCliente) <> "" Then
            .txtCliente.Text = Trim(txtCliente(1).Text)
            .txtCliente.Tag = Trim(txtCliente(0).Text)
            .cpFactura = vtablaFactura  '"Factura"
        End If
        .CmdEjecutarCobro.Enabled = True
        
        
        .finit
        .vieneCobro = True
        '.Show
        .chkFechaTodas.Value = True
        .cmdFiltrar_Click
        .Show
         .viene = "cobro"
    End With
    
    HabilitarControles (True)
    

If Err Then GrabarLog "PusBuscarDocumento_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub PusCerrar_Click(Index As Integer)
On Error Resume Next


Unload Me ' panic

End Sub
'    If Not frmRemito.vActualizaNombre = True Then
'
'
'        If (esFacturacion = True) And (vImpresoras.vNombreImpresora = "Hasar") Then
'            If Not vImpresionCorrecta = True Then
'                If MsgBox("Desea Imprimir el Documentos ?", vbInformation + vbYesNo, "Mensaje ...") = vbYes Then
'                    If Not (remito = 0) Then
'                        Call ImprimirHasar(remito, Val(vImportePagoPesos))
'                    End If
'                End If
'            End If
'
'        Else
'
'            imprimirFacturaNoFiscal
'
'
'        End If
'
'
'
'
'
'        Unload Me
'
'    Else
'
'        MsgBox "No puede realizar una Cta. Cte a un cliente CONSUMIDOR FINAL."
'
'    End If
'
'
'
'
'
'
'If Err Then GrabarLog "PusCerrar_Click", Err.Number & " " & Err.Description, Me.Caption
'End Sub

Private Sub imprimirFacturaNoFiscal()
' hacer el módulo de impresión no fiscal
ifacta.Show
End Sub

Private Sub cmdEliminarCheque_Click()
On Error Resume Next

    EliminarFilaCheque

If Err Then GrabarLog "cmdEliminarCheque_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub EliminarFilaCheque()
On Error Resume Next

    With rsCheque
        If Not .EOF = True Then
            txtImporteTotalCheque.Text = Val(txtImporteTotalCheque.Text) - Val(rsCheque.Fields("monto").Value)
            rsCheque.Delete
        End If
    End With

If Err Then GrabarLog "EliminarFilaCheque", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub PusGrabar_Click(Index As Integer)
On Error Resume Next
Exit Sub

    vImpresionCorrecta = False

    VaciarTablaRecibo

    If Not Val(TxtTotalAPagar.Text) = Val(vImporteTotalAPagar) Then
        MsgBox "EEEEERRRRRRRRRRROOOOOOOOORRRRRR", vbCritical, "Mensaje ..."
        Exit Sub
    End If
    If Val(TxtTotalAPagar.Text) = 0 Then
        MsgBox "El monto a pagar debe ser mayor a 0", vbInformation, "Mensaje ..."
        Exit Sub
    Else
        If Val(txtImporteEfectivoPesos) > 0 Then
    
            Dim importeAPagar, importePagadoPesos As Double
            importePagadoPesos = 0
            vImportePagoPesos = 0
            
            If esComprobanteAutomatico = True Then
                Call PagarCtaCteAutomaticamente(Val(txtImporteEfectivoPesos.Text), 1)
                importePagadoPesos = Val(txtImporteEfectivoPesos.Text)
                
            Else
                'Guardo en cta cte
                Dim i As Integer
                For i = 1 To Me.KlexDetalle.Rows - 1
                    If Val(txtImporteEfectivoPesos.Text) <= 0 Then
                        MsgBox KlexDetalle.TextMatrix(i, 2) & " " & KlexDetalle.TextMatrix(i, 3) & " no pudo ser pagados", vbInformation, "WGestion"
                    Else
                        'Si el total a pagar es menor que lo pendiente pago ese total, sino todo lo pendiente de ese documento
                        If (Val(KlexDetalle.TextMatrix(i, 5)) > Val(txtImporteEfectivoPesos)) Then
                            Call PagarCtaCte2(KlexDetalle.TextMatrix(i, 7), Val(Me.txtImporteEfectivoPesos), dtpFecha.Value, 1)
                            importePagadoPesos = importePagadoPesos + Val(txtImporteEfectivoPesos.Text)
                            'vImportePagoPesos = Val(importePagadoPesos)
                            txtImporteEfectivoPesos.Text = 0
                        Else
                            Call PagarCtaCte2(KlexDetalle.TextMatrix(i, 7), Me.KlexDetalle.TextMatrix(i, 5), dtpFecha.Value, 1)
                            txtImporteEfectivoPesos.Text = Val(Me.txtImporteEfectivoPesos.Text) - Me.KlexDetalle.TextMatrix(i, 5)
                            importePagadoPesos = importePagadoPesos + KlexDetalle.TextMatrix(i, 5)
                        End If
                        
                        'Agrego una linea en el recibo por el monto pagado en efectivo
                        Call AgregarPagoRecibo(1, "Efectivo en Pesos:", importePagadoPesos)
                
                        'Guardo el importe en efectivo en caja
                        Call WCaja(importePagadoPesos)
                        
                        vImportePagoPesos = importePagadoPesos
            
                    End If
                Next i
                
            End If
            
            If esFacturacion = True And vConfigGral.vIncluyeTicket = True Then
                ' PrinterFiscal.SendInvoicePayment "Pago en efectivo: ", importePagadoPesos, "T"
            Else

            End If
        End If
    
    ' ------------------------------------------
    ' cheques
    ' ------------------------------------------

        If Val(txtImporteTotalCheque.Text) > 0 Then
            Dim importePagadoCheque As Double
            importePagadoCheque = 0
            'Agrego una linea en el recibo por el monto pagado con cada cheque
            If Not rsCheque.RecordCount = 0 Then
                rsCheque.MoveFirst
                Do While Not rsCheque.EOF = True
                    AgregarPagoRecibo 4, "Cheque Nro.: " & Str(rsCheque.Fields("NCheque").Value) & " de Banco " & rsCheque.Fields("Banco") & "-" & rsCheque.Fields("bancoscuentas") & "- F.Dep: " & rsCheque.Fields("Deposito"), Val(rsCheque.Fields("Monto"))
                    importePagadoCheque = Val(importePagadoCheque) + Val(Format(rsCheque.Fields("Monto").Value, "#######0.00"))
                    rsCheque.MoveNext
                Loop
            End If
    
            'Guardo los cheques
            GuardarCheques
    
    
            If esComprobanteAutomatico Then
                PagarCtaCteAutomaticamente (Val(txtImporteTotalCheque.Text)), 4
            'Else
                'Guardo en cta cte
            '    Call PagarCtaCte2(remito, Val(txtImporteCheque), Me.dtpFecha.Value, 4)
            'End If
            Else
                'Guardo en cta cte
                
                For i = 1 To KlexDetalle.Rows - 1
                    If Val(txtImporteTotalCheque.Text) <= 0 Then
                        MsgBox KlexDetalle.TextMatrix(i, 2) & " " & KlexDetalle.TextMatrix(i, 3) & " no pudo ser pagados", vbInformation, "WGestion"
                    Else
                        'Si el total a pagar es menor que lo pendiente pago ese total, sino todo lo pendiente de ese documento
                        If (KlexDetalle.TextMatrix(i, 5) > Me.txtImporteTotalCheque) Then
                            Call PagarCtaCte2(KlexDetalle.TextMatrix(i, 7), Val(Me.txtImporteTotalCheque), dtpFecha.Value, 4)
                            Me.txtImporteEfectivoPesos.Text = 0
                        Else
                            Call PagarCtaCte2(KlexDetalle.TextMatrix(i, 7), Me.KlexDetalle.TextMatrix(i, 5), dtpFecha.Value, 4)
                            txtImporteTotalCheque.Text = Val(Me.txtImporteTotalCheque.Text) - Me.KlexDetalle.TextMatrix(i, 5)
                        End If
                        
                    End If
                Next i
                
            End If
    
            If esFacturacion = True And vConfigGral.vIncluyeTicket = True Then
                ' PrinterFiscal.SendInvoicePayment "Pago con cheque/s: ", importePagadoCheque, "T"
            End If
        End If


        If Val(txtImporteCuponTarjeta.Text) > 0 Then
    
            Dim totalpagadoTarjeta As Double
            totalpagadoTarjeta = 0
            
            If Me.esComprobanteAutomatico Then
                Me.PagarCtaCteAutomaticamente (Val(Me.txtImporteCuponTarjeta.Text)), 3
                totalpagadoTarjeta = totalpagadoTarjeta + Val(Me.txtImporteCuponTarjeta.Text)
            Else
                'Guardo en cta cte
                For i = 1 To Me.KlexDetalle.Rows - 1
                    If Val(Me.txtImporteCuponTarjeta.Text) <= 0 Then
                        MsgBox KlexDetalle.TextMatrix(i, 2) & " " & KlexDetalle.TextMatrix(i, 3) & " no pudo ser pagados", vbInformation, "WGestion"
                        totalpagadoTarjeta = totalpagadoTarjeta + Val(Me.txtImporteCuponTarjeta.Text)
                    Else
                        'Si el total a pagar es menor que lo pendiente pago ese total, sino todo lo pendiente de ese documento
                        If (Me.KlexDetalle.TextMatrix(i, 5) > Val(Me.txtImporteCuponTarjeta.Text)) Then
                            Call PagarCtaCte2(KlexDetalle.TextMatrix(i, 7), Val(Me.txtImporteCuponTarjeta), dtpFecha.Value, 3)
                            Me.txtImporteCuponTarjeta.Text = 0
                            totalpagadoTarjeta = totalpagadoTarjeta + Val(Me.txtImporteCuponTarjeta)
                        Else
                            Call PagarCtaCte2(KlexDetalle.TextMatrix(i, 7), Me.KlexDetalle.TextMatrix(i, 5), dtpFecha.Value, 3)
                            txtImporteCuponTarjeta.Text = Val(Me.txtImporteCuponTarjeta.Text) - Me.KlexDetalle.TextMatrix(i, 5)
                            totalpagadoTarjeta = totalpagadoTarjeta + KlexDetalle.TextMatrix(i, 5)
                        End If
                        
                    End If
                Next i
                
            End If
            
            'Agrego una linea en el recibo por el monto pagado con tarjeta
            AgregarPagoRecibo 3, "Tarjeta " & Me.cboTarjeta.Text & " de Banco " & Me.cboBancoTarjeta.Text, totalpagadoTarjeta
    
            Dim idCuponTarjetaNuevo As Integer
    
            'Guardo los datos de la operacion en cupon tarjeta
            idCuponTarjetaNuevo = GuardarCuponTarjeta
    
            'Guardo un movimiento en la cuenta de bancos por el total de lo pagado con tarjeta
            Call GuardarBancosMovimientos(vnrorecibo, cboBancoTarjeta.Tag, 1, 0, totalpagadoTarjeta, Me.txtObservaciones.Text, idCuponTarjetaNuevo, Val(txtNroInterno.Text))
    
            If esFacturacion = True And vConfigGral.vIncluyeTicket = True Then
                ' PrinterFiscal.SendInvoicePayment "Pago con tarjeta " & Me.cboTarjeta.Text & ":", totalpagadoTarjeta, "T"
            End If
    
        End If

        If Val(txtDepositoImporte.Text) > 0 Then
    
            Dim totalPagadoDeposito As Double
            totalPagadoDeposito = 0
            'txtDepositoImporte.Text = 0
            
            If esComprobanteAutomatico Then
                Call PagarCtaCteAutomaticamente(Val(txtDepositoImporte.Text), 5)
                totalPagadoDeposito = totalPagadoDeposito + Val(txtDepositoImporte.Text)
            Else
                'Guardo en cta cte
                For i = 1 To Me.KlexDetalle.Rows - 1
                    If Val(txtDepositoImporte.Text) <= 0 Then
                        MsgBox KlexDetalle.TextMatrix(i, 2) & " " & KlexDetalle.TextMatrix(i, 3) & " no pudo ser pagados", vbInformation, "WGestion"
                        totalPagadoDeposito = totalPagadoDeposito + Val(txtDepositoImporte.Text)
                    Else
                        'Si el total a pagar es menor que lo pendiente pago ese total, sino todo lo pendiente de ese documento
                        If (KlexDetalle.TextMatrix(i, 5) > Val(txtDepositoImporte.Text)) Then
                            Call PagarCtaCte2(KlexDetalle.TextMatrix(i, 7), Val(Me.txtDepositoImporte.Text), dtpFecha.Value, 5)
                            
                            totalPagadoDeposito = totalPagadoDeposito + Val(txtDepositoImporte.Text)
                        Else
                            Call PagarCtaCte2(KlexDetalle.TextMatrix(i, 7), Me.KlexDetalle.TextMatrix(i, 5), dtpFecha.Value, 5)
                            txtDepositoImporte.Text = Val(txtDepositoImporte.Text) - Val(KlexDetalle.TextMatrix(i, 5))
                            totalPagadoDeposito = Val(totalPagadoDeposito) + Val(KlexDetalle.TextMatrix(i, 5))
                        End If
                        
                    End If
                Next i
                        
            End If
            
            'Agrego una linea en el recibo por el monto pagado con Deposito Bancario
            Call AgregarPagoRecibo(5, Trim("Deposito en la Cuenta " & Trim(txtDepositoBanco(3).Text) & " de Banco " & Trim(txtDepositoBanco(1).Text)), totalPagadoDeposito)

            'Guardo un movimiento en la cuenta de bancos por el total de lo Depositado
            Call GuardarBancosMovimientos(vnrorecibo, Trim(txtDepositoBanco(0).Text), Val(txtDepositoBanco(2).Text), totalPagadoDeposito, 0, txtDepositoComentario.Text, 0, Val(txtNroInterno.Text))
    

    
        End If
        
        If esFacturacion = True And vConfigGral.vIncluyeTicket = True Then
            ' PrinterFiscal.CloseInvoice "T", "A", "TOTAL:"
        End If


        If Err.Number = 0 Then
            MsgBox "El pago se realizo exitosamente", vbInformation, "WGestion"
        Else
            MsgBox Err.Description
        End If
        
        If (esFacturacion = True) And (InStr(vConfigGral.vImpresoraSeleccionada, "Hasar") > 0) Then
            If Not vImpresionCorrecta = True Then
                Call ImprimirHasar(remito, vImportePagoPesos)
            End If
        End If
                
        If vConfigGral.vIncluyeContabilidad = True Then CargarContabilidad
        
        If vConfigGral.vImprimirReciboCliente = True Then ImprimirRecibo
        
        VaciarTablaRecibo
        
        KlexDetalle.Rows = 1
        HabilitarControles (False)

    End If

    frmRemito.vActualizaNombre = False
    
If Err Then
    MsgBox "Hubo errores al guardar, consulte el log del sistema", vbInformation, "WGestion"
Else
    Unload Me

    'MsgBox "El pago se realizo exitosamente", vbInformation, "WGestion"
End If
End Sub
Private Sub CargarContabilidad()
On Error Resume Next

    With frmAsientosAlta
        .Show
        .vVieneTabla = "pfactura"
        
        .chkControlar.Value = xtpChecked
        .txtCuentaVieneDe.Text = Me.Caption
        .txtCuentaVieneDe.Tag = Trim(txtCliente(0).Text)
        If txtObservaciones.Text = "" Then
            .txtLeyenda.Text = "Cobro: " & txtTipoComp.Text & " " & txtNroComprobante.Text
        Else
            .txtLeyenda.Text = Trim(txtObservaciones.Text)
        End If
        .dtpFecha.Value = vfechaCredito.Value
        
        'Panic
        .txtImporteVieneDe.Text = Val(vImporteTotalAPagar)
        
        
        .lblNroInterno.Caption = Val(txtNroInterno.Text)
        
        .cboTipoMovimiento.Tag = "RC"
        .cboTipoMovimiento.Text = "Recibo de Cobro"
        
        .vVieneTabla = "CuentasCorrientes"
        .vVieneIdNombre = "id"
        .vVieneIdValor = vIdCtaCteC
        
        .vcliprovee.Tag = Me.txtCliente(0)
        .vcliprovee.Text = Me.txtCliente(1)
        
        
        If cpInstancia = "cobro" Then
            .vCodigoCliente = Me.txtCliente(0).Text
            .vCodigoProveedor = ""
        End If
    
        If cpInstancia = "pagos" Then
            .vCodigoProveedor = Me.txtCliente(0).Text
            .vCodigoCliente = ""
        End If
        
        .vconcepto.Tag = Me.vconcepto.Tag
        .vconcepto.Text = Me.vconcepto.Text
        
        .Show
              
    End With

If Err Then GrabarLog "CargarContabilidad", Err.Number & " " & Err.Description, Me.Caption
End Sub
Public Sub WCaja(importePagado As Double)
    On Error Resume Next
    
    Dim rsCaja As New ADODB.Recordset
    Dim sqlCaja As String
    
    sqlCaja = "SELECT * FROM caja"
    
    With rsCaja
        Call .Open(sqlCaja, ConnDDBB, adOpenDynamic, adLockPessimistic)
    
        .AddNew
        .Fields("remito").Value = Val(remito)
        
        .Fields("fecha").Value = strfechaMySQL(dtpFecha.Value)
        .Fields("Importe").Value = importePagado
        
        .Fields("CodigoCliente").Value = Trim(codCliente)
        
        .Fields("Usuario").Value = vConfigGral.vUser
        .Fields("CodigoConcepto").Value = 221
        .Fields("comentario").Value = ""
            
        .Fields("NroCheque") = Null
        .Fields("FechaDeposito") = Null
        .Fields("FechaConfeccion") = Null
        .Fields("idCajas") = Null
        
        .Update
    
    End With
    
    sqlCaja = ""
    
    If rsCaja.State = 1 Then
        rsCaja.Close
        Set rsCaja = Nothing
    End If
    
If Err Then GrabarLog "WCaja", Left(Err.Number & " " & Err.Description, 99), Me.Name
End Sub

Private Sub txtBancoCheque_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
    If KeyCode = vbKeyF3 Then
        If Index = 2 Then
            pbCarga_Click (1)
        End If
    End If
    
If Err Then GrabarLog "txtBancoCheque_KeyUp", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtBancoCheque_LostFocus(Index As Integer)

'If Index = 0 Or Index = 2 Or Index = 4 Then
'
'    vsql = "select * bancos where idcodigo='" + Trim(txtBancoCheque(Index)) + "'"
'    txtBancoCheque(Index).Tag = traerDatos2(vsql, "idBancos", pathDBMySQL)'
'
'End If

End Sub

Private Sub txtCantCuotas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   ' Me.PusGrabar(0).SetFocus
End If
End Sub

Private Sub txtCliente_Change(Index As Integer)
Dim vsql As String


If Not txtCliente(0).Text = "" Then
    Call vsaldo_Click
    
    vsql = "select * from clientes where codigo = '" + Trim(Me.txtCliente(0).Text) + "'"
    
    vcuit = traerDatos2(vsql, "cuit", pathDBMySQL)


End If

End Sub

Private Sub txtCliente_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 Call PusBuscarCliente_Click
End Sub

Private Sub txtCliente_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        
        
        Select Case Index
        
            Case 0
'                txtCliente(1).Text = TraerDato("Clientes", "Codigo = '" & Trim(txtCliente(0).Text) & "'", "RazonSocial")
'
'                Me.codCliente = txtCliente(0)
'
'                If Not Trim(txtCliente(1).Text) = "" Then
'                    HabilitarControles (True)
'                Else
'                    txtCliente(0).Text = ""
'                    txtCliente(1).Text = ""
'                    txtCliente(0).SetFocus
'                End If
            
            Case 1
        
        End Select
        
    End If

If Err Then GrabarLog "txtCliente_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtCliente_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
    If KeyCode = vbKeyF3 Then
        If Index = 0 Then
            pbCarga_Click (2)
        End If
    
    End If
    
If Err Then GrabarLog "txtCliente_KeyUp", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtCotizacionDolar_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
   ' Me.txtBancoCheque(0).SetFocus
End If
If Err Then Exit Sub
End Sub
Private Sub txtDepositoBanco_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
    
        Select Case Index
        
            Case 0
                txtDepositoBanco(Index + 1).Text = TraerDato("Bancos", "idBancos = '" & Trim(txtDepositoBanco(Index).Text) & "'", "Descripcion")
            '    txtDepositoBanco(Index + 2).SetFocus
            
            Case 2
                txtDepositoBanco(Index + 1).Text = TraerDato("BancosCuentas", "idBancosCuentas = " & Trim(txtDepositoBanco(Index).Text) & "", "Cuenta")
            '    txtDepositoImporte.SetFocus
                
        End Select
    
    
      '  If txtDepositoBanco(Index).Text = "" Then txtDepositoImporte.SetFocus
    
    End If

If Err Then GrabarLog "txtDepositoBanco_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtDepositoBanco_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
    If KeyCode = vbKeyF3 Then
        If Index = 2 Then
            pbCarga_Click (4)
        End If
    End If
    
If Err Then GrabarLog "txtBancoCheque_KeyUp", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtDepositoComentario_Change()
txtObservaciones = Me.txtDepositoComentario.Text
End Sub

Private Sub txtDepositoComentario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'Me.cmdCobrar.SetFocus
End If
End Sub

Private Sub txtDepositoImporte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  '  Me.vfechaDeposito.SetFocus
End If
End Sub

Private Sub txtFechaDeposito_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'Me.txtnroInternoDeposito.SetFocus
End If
End Sub

Private Sub txtFechaDeposito_LostFocus()
'If txtFechaDeposito.Value = Me.vfechaDeposito.Value Then
'    MsgBox "Le advertimos que las fechas de depósito y de movimiento son iguales", vbInformation
'End If
End Sub

Private Sub txtFirmanteCheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'Me.vfechaDeposito.SetFocus
    End If
End Sub

Private Sub txtImporteCuponTarjeta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'Me.txtCantCuotas.SetFocus
End If
End Sub

Private Sub txtImporteCheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      '  Me.pbCarga(5).SetFocus
    End If
End Sub

Private Sub calTotales()
    'calTotales
    TxtTotalAPagar.Text = Val(txtImporteTotalCheque.Text) + Val(txtImporteEfectivoPesos.Text) + Val(txtImporteCuponTarjeta.Text) + Val(txtDepositoImporte.Text) + Val(Me.vtimporteRetenciones)
End Sub


Private Sub txtImporteEfectivoPesos_Change()
    calTotales
    'TxtTotalAPagar.Text = Val(txtImporteTotalCheque.Text) + Val(txtImporteEfectivoPesos.Text) + Val(txtImporteCuponTarjeta.Text) + Val(txtDepositoImporte.Text)
End Sub
Private Sub txtImporteCuponTarjeta_Change()
    calTotales
    'TxtTotalAPagar.Text = Val(txtImporteTotalCheque.Text) + Val(txtImporteEfectivoPesos) + Val(txtImporteCuponTarjeta) + Val(txtDepositoImporte.Text)
End Sub
Private Sub txtdepositoImporte_Change()
    calTotales
    'TxtTotalAPagar.Text = Val(txtImporteTotalCheque.Text) + Val(txtImporteEfectivoPesos) + Val(txtImporteCuponTarjeta) + Val(txtDepositoImporte.Text)
End Sub
Private Sub txtImporteEfectivoDolar_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Me.txtImporteEfectivoDolar.Text = "" Then
            
           ' Me.txtBancoCheque(0).SetFocus
        Else
        '    Me.txtCotizacionDolar.SetFocus
        End If
    End If
End Sub

Private Sub txtImporteEfectivoPesos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    '    Me.txtImporteEfectivoDolar.SetFocus
    End If
End Sub

Private Sub txtImporteTotalCheque_Change()
    calTotales
    'TxtTotalAPagar.Text = Val(txtImporteTotalCheque.Text) + Val(txtImporteEfectivoPesos.Text) + Val(txtImporteCuponTarjeta.Text) + Val(txtDepositoImporte.Text)
End Sub

Private Sub txtMontoTotalPendienteSeleccionado_Change()
    'txtMontoTotalPendienteSeleccionado.Text = Val(txtMontoTotalPendienteSeleccionado.Text)
End Sub

Private Sub txtNroCheque_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  '  Me.cboEstadoCheque.SetFocus
End If
End Sub

Private Sub txtNroComprobante_Change()

    'Me.txtPendiente = CalcularSaldo(remito)
    'Me.txtTotal.Text = CalcularTotal(remito)
    'Me.txtPagado.Text = CalcularPagado(remito)

End Sub

Public Sub AgregarDocumentoAPagar(total As Double, pendiente As Double, pagado As Double, nroComp As Long, tipoComp As String, fechaComp As Date, remito As Integer)
    On Error Resume Next
    Dim i, j As Integer
    
    With KlexDetalle
        If .Rows <= 2 And .TextMatrix(.Rows - 1, 4) = "" Then
            FormatoGrillaDetalle (1)
        Else
            .Rows = .Rows + 1
        End If
        j = .Rows - 1
        
        .TextMatrix(j, 1) = fechaComp
        .TextMatrix(j, 2) = tipoComp
        .TextMatrix(j, 3) = nroComp
        .TextMatrix(j, 4) = total
        .TextMatrix(j, 5) = pendiente
        .TextMatrix(j, 6) = pagado
        .TextMatrix(j, 7) = remito

    End With
    
    If Err Then GrabarLog "AgregarDocumentoAPagar", Left(Err.Number & " " & Err.Description, 99), Me.Name
End Sub
Public Sub BuscarDatosOperacionesCliente(codCli As String, remito As Long)
    On Error Resume Next
    
    If LeerConfig(23) = True Then
    
        Dim rsCtaCteC As New ADODB.Recordset, sqlCtaCteC As String, i As Integer
    
        SaldoAnterior = 0
        totaldebito = 0
        totalCredito = 0
        credito = 0
        debito = 0
    
        If remito <> 0 Then
            sqlCtaCteC = "SELECT * FROM " + CP.TablaCtaCte + " WHERE (codigo = '" & codCli & "') and (remito = " & remito & ")"
        Else
            sqlCtaCteC = "SELECT * FROM " + CP.TablaCtaCte + " WHERE (codigo = '" & codCli & "')"
        End If
    
        With rsCtaCteC
            .CursorLocation = adUseClient
                   
            Call .Open(sqlCtaCteC, ConnDDBB, adOpenStatic, adLockPessimistic)
            
            Do While Not .EOF = True
                If IsNull(.Fields("debito").Value) Or .Fields("debito").Value = "" Then
                    debito = 0
                Else
                    debito = .Fields("debito").Value
                    If Not remito = 0 Then
                        credito = GenerarDato("SELECT Sum(Importe) FROM Cobros WHERE (Remito = " & Val(remito) & " )", "Sum(Importe)")
                    Else
                        If IsNull(.Fields("credito").Value) Or .Fields("credito").Value = "" Then
                            credito = 0
                        Else
                            credito = .Fields("credito").Value
                        End If
                    End If
                End If
                            
    
                            
                SaldoAnterior = SaldoAnterior + debito - credito
                totalCredito = totalCredito + credito
                totaldebito = totaldebito + debito
                
                .MoveNext
            Loop
            
        End With

    
        If totaldebito = "" Then
            totaldebito = 0
        End If
    
        If totalCredito = "" Then
            totalCredito = 0
        End If
        
        If SaldoAnterior = "" Then
            SaldoAnterior = 0
        End If
        
        total = totaldebito
        pendiente = SaldoAnterior
        pagado = totalCredito
        
        Dim yaCargado As Boolean
        yaCargado = False
        For i = 1 To KlexDetalle.Rows - 1
            If Me.remito = Val(KlexDetalle.TextMatrix(i, 7)) And KlexDetalle.TextMatrix(i, 7) <> "" Then
                MsgBox KlexDetalle.TextMatrix(i, 2) & " " & KlexDetalle.TextMatrix(i, 3) & " ya ha sido cargado", vbInformation, "WGestion"
                yaCargado = True
            End If
        Next i
        
        If Not esComprobanteAutomatico Then
            If Not yaCargado Then
                AgregarDocumentoAPagar Me.total, Me.pendiente, Me.pagado, Me.NroComprobante, Me.tipoComprobante, Me.fechaDocumento, Me.remito
                txtTotal.Text = totaldebito
                txtPendiente.Text = SaldoAnterior
                txtPagado.Text = totalCredito
                
                'Muestro el total pendiente
                txtMontoTotalPendienteSeleccionado.Text = Format(Val(txtMontoTotalPendienteSeleccionado) + Val(pendiente), "######0.00")
            End If
        Else
           Dim rsFac As New ADODB.Recordset, sqlFac As String
           txtMontoTotalPendienteSeleccionado = 0
           sqlFac = "SELECT * FROM Factura WHERE (codigo = '" & codCli & "')"
        
           With rsFac
               .CursorLocation = adUseClient
                  
               Call .Open(sqlFac, ConnDDBB, adOpenStatic, adLockPessimistic)
               Do While Not rsFac.EOF
                   CalcularSaldosPorRemito rsFac("remito").Value, codCli
                   If Me.pendiente > 0 Then
                       Me.AgregarDocumentoAPagar Me.total, Me.pendiente, Me.pagado, rsFac("NComprobante"), rsFac("Tipo"), rsFac("Fecha"), rsFac("remito")
                   
                       txtTotal.Text = Val(txtTotal.Text) + totaldebito
                       txtPendiente.Text = Val(txtPendiente.Text) + SaldoAnterior
                       txtPagado.Text = Val(txtPagado.Text) + totalCredito
                       
                       'Muestro el total pendiente
                       txtMontoTotalPendienteSeleccionado = Val(Me.txtMontoTotalPendienteSeleccionado) + Me.pendiente
                       
                   End If
                   rsFac.MoveNext
               Loop
           End With
        End If
    
        If Val(Me.TxtTotalAPagar) < Val(txtMontoTotalPendienteSeleccionado) Then
            txtMontoTotalPendienteSeleccionado.ForeColor = &HFF&
        Else
            txtMontoTotalPendienteSeleccionado.ForeColor = &H80000008
        End If
    Else
        txtMontoTotalPendienteSeleccionado.Text = GenerarDato("SELECT Sum(Debito), Sum(Credito), Sum(Debito)-Sum(Credito) FROM CuentasCorrientes WHERE Codigo = '" & Trim(codCli) & "';", "Sum(Debito)-Sum(Credito)")
    End If
    
    SetearDatosCliente (codCli)
    
    'If Not esFacturacion = True Then
    '    txtNroComprobante.Text = "Automático"
    '    txtTipoComp.Text = "No aplica"
    'End If
    
If Err Then GrabarLog "BuscarDatosOperacionesCliente", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub CalcularSaldosPorRemito(remito As Long, codCli As String)
    
    SaldoAnterior = 0
    totaldebito = 0
    totalCredito = 0
    credito = 0
    debito = 0
    
    Dim rsCtaCteC As New ADODB.Recordset, sqlCtaCteC As String
    sqlCtaCteC = "SELECT * FROM CuentasCorrientes WHERE (codigo = '" & codCli & "') and (remito = " & remito & ")"
    
    With rsCtaCteC
        .CursorLocation = adUseClient
               
        Call .Open(sqlCtaCteC, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        Do While Not .EOF
            If IsNull(.Fields("debito").Value) Or .Fields("debito").Value = "" Then
                debito = 0
            Else
                debito = .Fields("debito").Value
            End If
                        
            If IsNull(.Fields("credito").Value) Or .Fields("credito").Value = "" Then
                credito = 0
            Else
                credito = .Fields("credito").Value
            End If
                        
            SaldoAnterior = SaldoAnterior + debito - credito
            totalCredito = totalCredito + credito
            totaldebito = totaldebito + debito
            
            .MoveNext
        Loop
        
    End With
    
    If totaldebito = "" Then
        totaldebito = 0
    End If
    
    If totalCredito = "" Then
        totalCredito = 0
    End If
    
    If SaldoAnterior = "" Then
        SaldoAnterior = 0
    End If
    
    total = totaldebito
    pendiente = SaldoAnterior
    pagado = totalCredito
End Sub

Private Sub FormatoGrillaDetalle(vCantidadRenglones As Integer)
On Error Resume Next

    Dim i As Integer

    With KlexDetalle
        .FixedRows = 1
        .FixedCols = 1
    
        .Cols = 8
        .Rows = vCantidadRenglones + 1
        
        If vCantidadRenglones = 1 Then
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
                .ColWidth(i) = 0
            Next
        End If
        
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 400
        
        .TextMatrix(0, 1) = "Fecha"
        .ColWidth(1) = 1100
               
        .TextMatrix(0, 2) = "Tipo comprobante"
        .ColWidth(2) = 2340
        
        .TextMatrix(0, 3) = "Nro. comprobante"
        .ColWidth(3) = 1500
        
        .TextMatrix(0, 4) = "Total"
        .ColWidth(4) = 1000
        .ColDisplayFormat(4) = "#0.000"
        
        .TextMatrix(0, 5) = "Pendiente"
        .ColWidth(5) = 1000
        .ColDisplayFormat(5) = "#0.000"
        
        .TextMatrix(0, 6) = "Pagado"
        .ColWidth(6) = 1000
        .ColDisplayFormat(6) = "#0.000"
        
        .TextMatrix(0, 7) = "Remito"
        .ColWidth(7) = 1000

    End With
    
If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtNroCuponTarjeta_GotFocus()
    'Me.Resizer.VScrollPosition = 4905
End Sub
Private Sub txtNroCuponTarjeta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.txtNroCuponTarjeta.Text = "" Then
       ' PusGrabar(0).SetFocus
    Else
     '   Me.cboBancoTarjeta.SetFocus
    End If
End If
End Sub

Private Sub txtNroCheque_LostFocus()
If Me.cpInstancia = "cobro" Then
    ValidadNroChe (txtNroCheque) ' valida si eexiste el nro de cheque en bancosmovimientos
End If
End Sub

Private Sub txtNroChequeDeposito_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Me.txtDepositoComentario.SetFocus
End If
End Sub

Private Sub txtNroChequeDeposito_LostFocus()
ValidadNroChe (txtNroChequeDeposito)
End Sub

Private Sub txtNroInternoCheque_Change()
'Me.txtNroInterno = Me.txtNroInternoCheque
End Sub

Private Sub txtNroInternoCheque_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  '  dtpDepositoCheque.SetFocus
End If
End Sub

Private Sub txtnroInternoDeposito_Change()
'Me.txtNroInterno = Me.txtnroInternoDeposito
End Sub

Private Sub txtnroInternoDeposito_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  '  Me.txtNroChequeDeposito.SetFocus
End If
End Sub

Private Sub txtObservaciones_GotFocus()
On Error Resume Next
    
    With txtObservaciones
        .SelStart = 0
        .SelLength = Len(txtObservaciones.Text)
    End With

If Err Then GrabarLog "txtObservaciones_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtPagado_Change()
    'If txtPagado <> "" Then
    '    Me.txtPagado.Text = pagado
    'End If
End Sub

Private Sub txtPendiente_Change()
    'If Me.txtPendiente <> "" Then
    '    Me.txtPendiente.Text = pendiente
    'End If
End Sub

Private Sub txtTotal_Change()
    'If txtTotal <> "" Then
    '    Me.txtTotal.Text = total
    'End If
    'me.txtPendiente.Text =
End Sub
Private Function SetearDatosCliente(codCli As String) As Long
    On Error Resume Next
    
    Dim rsCli As New ADODB.Recordset, sqlCtaCteC As String, sqlCli As String
    
    sqlCli = "SELECT * FROM  " + CP.TablaPersona + " WHERE (codigo = '" & codCli & "')"
     
    With rsCli
        .CursorLocation = adUseClient
               
        Call .Open(sqlCli, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not .EOF Then
        
            .MoveFirst
            
            'If frmRemito.vActualizaNombre = True Then
            '    txtCliente(1).Text = EsNulo(frmRemito.vNombreNuevo)
            'Else
            '    txtCliente(1).Text = EsNulo(.Fields("Nombre").Value)
            'End If
            txtCliente(0).Text = EsNulo(.Fields("Codigo").Value)
            
            'If Not IsNull(.Fields("CreditoMax").value) And Not .Fields("CreditoMax").value = 0 Then
            '    txtCredMax.Text = .Fields("CreditoMax").value
            'Else
            '    txtCredMax.Text = "No definido"
            'End If
        End If
    End With
    
    sqlCli = ""

    If rsCli.State = 1 Then
        rsCli.Close
        Set rsCli = Nothing
    End If
    
If Err Then GrabarLog "SetearDatosCliente", Err.Number & " " & Err.Description, Me.Caption
End Function

Public Sub PagarCtaCteAutomaticamente(importeAPagar As Double, idMedioPago As Integer)
    On Error Resume Next
    
    Dim rsCtaCteC As New ADODB.Recordset, sqlCtaCteC As String
    Dim cmd As New ADODB.Command
    
    With cmd
        Set .ActiveConnection = ConnDDBB
        .CommandText = "traer_DocumentosConImporteDeuda"
        .CommandType = adCmdStoredProc
        .Parameters.Append cmd.CreateParameter("codcli", adVarChar, adParamInput, 50, codCliente)
        .Prepared = True
    End With
        
    Set rsCtaCteC = cmd.Execute
         
    If Not rsCtaCteC.EOF = True Then rsCtaCteC.MoveFirst
    
    With rsCtaCteC
        
    'If Me.chkProrratear = 0 Then
    
    'comentarioFinal
    
    Me.txtObservaciones = Me.txtObservaciones + " [N.Int:" + Me.txtNroInterno + "]"
    
    Call PagarCtaCteDirecto(importeAPagar, idMedioPago, Me.txtObservaciones)
    
   ' Else
        
        
        Dim importePagado As Double
        
        Do While Not (.EOF) And (importeAPagar >= importePagado)
            'Llamo a pagarCtaCte con ese nro de remito
            If importeAPagar - importePagado > .Fields("ImporteDeudaDocumento").Value Then
                Call PagarCtaCte(.Fields("Remito"), .Fields("ImporteDeudaDocumento").Value, idMedioPago)
            Else
                Call PagarCtaCte(.Fields("Remito"), importeAPagar - importePagado, idMedioPago)
            End If
            
            
            If Left(Trim(concepto), Len(.Fields("comentario").Value)) <> Left(Trim(.Fields("comentario").Value), Len(concepto)) Or Len(concepto) = 0 Then
                concepto = .Fields("Comentario") & concepto & "; "
            End If
            
            
            '
            'Unload frmRemito
            importePagado = Val(Format(importePagado, "######0.00")) + Val(Format(.Fields("ImporteDeudaDocumento").Value, "######0.00"))
            .MoveNext
        Loop
    'End If
    End With
    
    
If Err Then GrabarLog "PagarCtaCteAutomaticamente", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub LimpiarCampos()
On Error Resume Next


    RadioButton1.Value = True
    
    With Me
        .Top = 0
        .Left = 0
        '.Height = 7875
        '.Width = 8580
        .KeyPreview = True
    End With
    
    

    Me.vconcepto.Text = ""
    Me.vconcepto.Tag = ""
            
    cboEstadoCheque.Text = ""
    txtFirmanteCheque.Text = ""
    txtImporteCheque.Text = ""
    txtImporteTotalCheque.Text = ""
    txtImporteEfectivoPesos.Text = ""
    txtImporteCuponTarjeta.Text = ""
    txtCliente(0).Text = ""
    txtCliente(1).Text = ""
    txtNroComprobante.Text = ""
    txtNroCheque.Text = ""
    txtPagado.Text = ""
    txtPendiente.Text = ""
    txtTipoComp.Text = ""
    txtTotal.Text = ""
    txtNroCuponTarjeta.Text = ""
    txtMontoTotalPendienteSeleccionado.Text = ""
    
    txtDepositoBanco(0).Text = ""
    txtDepositoBanco(1).Text = ""
    txtDepositoBanco(2).Text = ""
    txtDepositoBanco(3).Text = ""
    
    txtDepositoImporte.Text = ""
    txtDepositoComentario.Text = ""
    txtObservaciones.Text = ""
    concepto = ""
    txtObservaciones.Text = ""
    Me.vdocSeleccionado.Text = ""
    
    vMontoAPagarParaImprimir = 0
    codCliente = ""
    remito = 0
    
    dtpDepositoCheque.Value = Date
    dtpFecha.Value = Date

     LimpiarCheques
     
    VaciarChequesTemp
    
  '  Me.txtNroInterno = UltimoNroInterno2 + 1
    
   ' Me.txtnroInternoDeposito = UltimoNroInterno2 + 1
    
    
    Me.gretencion.Clear
    

If Err Then GrabarLog "LimpiarCampos", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Function GuardarCuponTarjeta() As Integer
    On Error Resume Next
    
    Dim rs As New ADODB.Recordset
    
    If (remito = 0) And (codCliente = "") Then
        MsgBox "Debe seleccionar un cliente o un remito antes de iniciar la operacion", vbOKOnly, "WGestion"
    Else
        Dim sql As String
    
        With rs
            .CursorLocation = adUseClient
            sql = "SELECT * FROM CuponTarjeta"
            Call .Open(sql, ConnDDBB, adOpenDynamic, adLockPessimistic)
            
            .AddNew
            .Fields("idtarjeta").Value = Me.cboTarjeta.Tag
        
            .Fields("idBanco").Value = Me.cboBancoTarjeta.Tag
            .Fields("Importe").Value = Val(Me.txtImporteCuponTarjeta.Text)
        
            .Fields("CantCuotas").Value = Val(Me.txtCantCuotas.Text)
        
            .Fields("NroCupon").Value = Trim(Me.txtNroCuponTarjeta.Text)
            
            .Update
           
            GuardarCuponTarjeta = .Fields("idCuponTarjeta").Value
                
        End With
    
    End If
    
If Err Then GrabarLog "GuardarCuponTarjeta", Left(Err.Number & " " & Err.Description, 99), Me.Name
End Function
Public Sub HabilitarControles(b As Boolean)
On Error Resume Next

    Dim i As Integer

    'txtCliente(0).Enabled = b
    'txtCliente(1).Enabled = b
    cboEstadoCheque.Enabled = b
    txtFirmanteCheque.Enabled = b
    txtImporteCheque.Enabled = b
    txtMontoTotalPendienteSeleccionado.Enabled = b
    txtImporteTotalCheque.Enabled = b
    txtImporteEfectivoPesos.Enabled = b
    txtImporteCuponTarjeta.Enabled = b
    
    txtNroComprobante.Enabled = b
    txtNroCheque.Enabled = b
    txtPagado.Enabled = b
    txtPendiente.Enabled = b
    txtTipoComp.Enabled = b
    txtTotal.Enabled = b
    txtNroCuponTarjeta.Enabled = b
    
    For i = 0 To Val(txtBancoCheque.Count - 3)
        txtBancoCheque(i).Enabled = b
    Next
    
    For i = 0 To Val(Me.pbCarga.Count - 2)
        pbCarga(i).Enabled = b
    Next

    cboBancoTarjeta.Enabled = b
    txtImporteEfectivoDolar.Enabled = b
    dtpDepositoCheque.Enabled = b
    cmdAgregarCheque.Enabled = b
    cmdEliminarCheque.Enabled = b
    cboBancoTarjeta.Enabled = b
    cboTarjeta.Enabled = b
    txtCantCuotas.Enabled = b
    cboTarjeta.Enabled = b
    PusGrabar(0).Enabled = b
    dtpFecha.Enabled = b

If Err Then GrabarLog "HabilitarControles", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Function ValidarIngresoCheque() As Boolean
    
    ValidarIngresoCheque = True
    
    If Val(txtImporteCheque.Text) = 0 Or Val(txtImporteCheque.Text) < 0 Then
        MsgBox "El importe del cheque debe ser un valor mayor o igual a cero.", vbInformation, "WGestion"
        ValidarIngresoCheque = False
        Exit Function
    End If
    
    If Val(txtNroCheque.Text) = 0 Then
        MsgBox "El campo Nro. cheque es de ingreso obligatorio", vbInformation, "WGestion"
        ValidarIngresoCheque = False
        Exit Function
    End If
        
    
    If Me.vCCajaDestino.Text = "" Then
        MsgBox "Debe ingresar la caja destino del cheque", vbInformation, "WGestion"
        ValidarIngresoCheque = False
        Exit Function
    End If
    
    If Me.vmarcaInterna = "" Then
        MsgBox "Debe ingresar una marca interna para el cheque", vbInformation, "WGestion"
        Me.vmarcaInterna = getMarcaIntarna()
        ValidarIngresoCheque = False
        Exit Function
    End If
    

End Function
Private Sub AgregarPagoRecibo(idMedioPago As Integer, desc As String, monto As Double)
On Error Resume Next

Dim vsql, vcampos, vvalores As String

vcampos = "idMedioPago,Descripcion,Monto,lugar,Fecha,Concepto,Total"

vvalores = "" + Str(idMedioPago) + ",'" + desc + "'," + Str(monto) + ",'" + (vDatosEmpresa.Localidad) + "'," + "'" + strfechaMySQL(dtpFecha.Value) + "'" + "," + ("''") + "," + Str(TxtTotalAPagar.Text)


Call InsertarEnTabla("Recibo_Temp", vcampos, vvalores)



If Err Then GrabarLog "AgregarPagoRecibo", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub VaciarTablaRecibo()
On Error Resume Next

    Call BorrarBase("Recibo_Temp", pathDBMySQL)
    
If Err Then GrabarLog "", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub ImprimirRecibo()
On Error Resume Next

    
    Unload Mantenimiento
    Load Mantenimiento
    
    
    imprimerecibo = False
    If MsgBox("Imprime el recibo del cobro/pago ? ", vbYesNo, "Recibos ...") = vbYes Then
        'llenarDrRecibo
        
       ' Call llenar_pago_temp  ' llenar dr titulo
        
        If LeerXml("UEmpresa") = "WGESTIONPOLI2" Then
            Call llenar_pago_temp  ' llenar dr titulo
            drRecibo11.Show
            
            imprimerecibo = True
        Else
            Call llenarDrRecibo
            drRecibo.Show
            imprimerecibo = True
        End If
        
    End If
    
    
    

If Err Then GrabarLog "ImprimirRecibo", Err.Number & " " & Err.Description, Me.Caption
End Sub



Private Sub llenar_pago1() ' todoing
Dim i, j As Integer

Dim v, vv As String


Dim a(20, 10) As String

a(1, 1) = "mediodepago"
a(1, 2) = "banco"
a(1, 3) = "sucursal"

a(2, 1) = "nrointerno"
a(2, 2) = "marcainterna"


a(3, 1) = "nrodocumento"
a(3, 2) = "ncheque"

a(4, 1) = "fecha"
a(4, 2) = "fechaAcreditacion"

a(5, 1) = "importe"
a(5, 2) = "monto"


'a(6, 1) = "comentario"
'a(7, 1) = "nroorden"

Dim ro  As New ADODB.Recordset   '
Dim rd  As New ADODB.Recordset  '
Dim vo, vd As String


Dim vcampos, vvalores As String

vo = "select * from cheques_temp order by fecha"
Call ro.Open(vo, pathDBMySQL, adOpenDynamic, adLockReadOnly)



vd = "select * from pago1"
Call rd.Open(vd, PathDBListados, adOpenDynamic, adLockBatchOptimistic)


' borra las temporales de listado
Call EjecutarScript("delete from pago1", PathDBListados)
Call EjecutarScript("delete from  pago2", PathDBListados)
Call EjecutarScript("delete from pago", PathDBListados)
'Call EjecutarScript("delete from temp2", pathDBMySQL)

            
Do Until ro.EOF
 'rd.AddNew
            For i = 1 To 5
                    '    rd.AddNew
                        
                            For j = 2 To 2
                                vv = ro.Fields(a(i, j))
                                    
                                
                                If Not vv = "" Then
                                    v = v + " " + vv
                                End If
                                
                                If Val(v) > 0 Then
                            
                                        vcampos = vcampos + "" + a(i, 1) + ","
                                        vvalores = vvalores + "'" + v + "',"
                                        
                                       ' rd.Fields(a(i, 1)) = Val(v)
                                Else
                                
                                        vcampos = vcampos + "" + a(i, 1) + ","
                                        vvalores = vvalores + "'" + v + "',"
                                        
                                        'rd.Fields(a(i, 1)) = v
                                End If
                                v = ""
                            Next
                            
             '               rd.Update
            Next
            
    vcampos = Left$(vcampos, Len(vcampos) - 1) + ",NroOrden"
    vvalores = Left$(vvalores, Len(vvalores) - 1) + ",'99'"
    
    vvalores = "'Cartera: " + Right$(vvalores, Len(vvalores) - 1)
    
    Dim vsql As String
    
    vsql = "insert into pago1 (" + vcampos + ") values (" + vvalores + ")"
    Call EjecutarScript(vsql, PathDBListados)
    
    vvalores = ""
    vcampos = ""
    
    'rd.Update
    
    ro.MoveNext
Loop

'--------------- cargar pago de caja

If Val(Me.txtImporteEfectivoPesos.Text) > 0 Then
        vcampos = "mediodepago,importe" + ",NroOrden"
        vvalores = "'Efectivos: '," + Str(Val(Me.txtImporteEfectivoPesos.Text)) + ",99"
        
        vsql = "insert into pago1 (" + vcampos + ") values (" + vvalores + ")"
        Call EjecutarScript(vsql, PathDBListados)

End If

'--------------- carga valores propios

If Val(Me.txtDepositoImporte) > 0 Then
    
    vcampos = "mediodepago,nrodocumento,fecha,nrointerno,importe" + ",NroOrden"
    vvalores = "'Valor: " + Me.txtDepositoBanco(1) + " / " + Me.txtDepositoBanco(3) + "','" + Me.txtNroChequeDeposito + "','" + Me.txtFechaDeposito.Text + "', '" + Me.vmarcainternaDeposito + "'," + Str(Val(Me.txtDepositoImporte)) + ",99"
    vsql = "insert into pago1 (" + vcampos + ") values (" + vvalores + ")"
    Call EjecutarScript(vsql, PathDBListados)

End If


ro.Close
rd.Close

End Sub

Private Sub llenar_pago1_temp() ' todoing
Dim i, j As Integer

Dim v, vv As String


'Call EjecutarScript("truncate temp2", pathDBMySQL)

   ' v = "insert into temp2 (c02,c05) values ('MEDIOS DE PAGOS                                NRO.DOC.              NRO.INT.           FECHA  ',' IMPORTE')"
   ' If Not v = "" Then Call EjecutarScript(v, pathDBMySQL)


     v = "insert into temp2 (c02,c05) values ('','')"
    If Not v = "" Then Call EjecutarScript(v, pathDBMySQL)



Dim a(20, 10) As String

a(1, 1) = "mediodepago"
a(1, 2) = "banco"
a(1, 3) = "sucursal"


a(2, 1) = "nrodocumento"
a(2, 2) = "ncheque"


a(3, 1) = "nrointerno"
a(3, 2) = "marcainterna"



a(4, 1) = "fecha"
a(4, 2) = "fechaAcreditacion"

a(5, 1) = "importe"
a(5, 2) = "monto"


'a(6, 1) = "comentario"
'a(7, 1) = "nroorden"

Dim ro  As New ADODB.Recordset   '
Dim rd  As New ADODB.Recordset  '
Dim vo, vd As String


Dim vcampos, vvalores As String

vo = "select * from cheques_temp order by fecha"
Call ro.Open(vo, pathDBMySQL, adOpenDynamic, adLockReadOnly)

vv = ""

Dim vtotal As Double

vtotal = 0

Do Until ro.EOF
 'rd.AddNew
            For i = 1 To 4
                    '    rd.AddNew
                        
                            For j = 2 To 2
                            
                                    If i = 1 Then
                            
                                        vv = vv + fc(ro.Fields(a(i, j)), 47)
                                    Else
                                        
                                        vv = vv + fc(ro.Fields(a(i, j)), 10)
                                    End If
                            
                            Next
             '               rd.Update
            Next
            
    vcampos = "c02,c05"
    vvalores = "'V. en Cartera:" + vv + "','" + Format(ro.Fields(a(5, 2)), "###,###,##0.00") + "'"
    
    vtotal = vtotal + Val(ro.Fields(a(5, 2)))
    
    Dim vsql As String
    
    vsql = "insert into temp2 (" + vcampos + ") values (" + vvalores + ")"
  ' Call EjecutarScript(vsql, pathDBMySQL)
    
    vvalores = ""
    vcampos = ""
    
    'rd.Update
    
    ro.MoveNext
Loop

'--------------- cargar pago de caja

If Val(Me.txtImporteEfectivoPesos.Text) > 0 Then
        'vcampos = "mediodepago,importe" + ",NroOrden"
                   
        vvalores = "'Importe en efectivo recibido:     ','" + Format(Val(Me.txtImporteEfectivoPesos.Text), "###,###,##0.00") + "'"
        vcampos = " c02,c05 "
        vsql = "insert into temp2 (" + vcampos + ") values (" + vvalores + ")"
        Call EjecutarScript(vsql, pathDBMySQL)
        
        vtotal = vtotal + Val(Me.txtImporteEfectivoPesos.Text)
        

End If

'--------------- carga valores propios

'If Val(Me.txtDepositoImporte) > 0 Then

If Me.dgCheques.VisibleRows > 0 Then

    
   ' vcampos = "mediodepago,nrodocumento,fecha,nrointerno,importe" + ",NroOrden"
    
    Dim i5 As Integer
    
    For i5 = 0 To Me.dgCheques.VisibleRows - 1
        
        
        vvalores = "'" + getLinea(i5) + "'," + getLinea2(i5)
        ' vvalores = "'Valores: " + Me.dgCheques.
        'vvalores = "'Valores:      " + fc(Me.txtDepositoBanco(1), 25) + " / " + fc(Me.txtDepositoBanco(3), 13) + "  " + fc(Me.txtNroChequeDeposito, 20) + "   " + fc(Me.txtFechaDeposito.Text, 10) + " ','" + Format(Val(Me.txtDepositoImporte), "###,###,##0.00") + "'"
        vsql = "insert into temp2 (" + vcampos + ") values (" + vvalores + ")"
        Call EjecutarScript(vsql, pathDBMySQL)
        
    
    vtotal = vtotal + Val(Me.txtDepositoImporte)
    
    Next

End If

Dim v2 As String



' aca tengo que poner las retenciones -----------


'vtotal = vtotal + llenar_retenciones_temp(vtotal)




v2 = "'',''"
vsql = "insert into temp2 (c02,c05) values (" + v2 + ")"
Call EjecutarScript(vsql, pathDBMySQL)


v2 = "'----------------------',''"
vsql = "insert into temp2 (c02,c05) values (" + v2 + ")"
Call EjecutarScript(vsql, pathDBMySQL)


v2 = "'Importe total recibo: ','" + Format(vtotal, "###,###,##0.00") + "'"
vsql = "insert into temp2 (c02,c05) values (" + v2 + ")"
Call EjecutarScript(vsql, pathDBMySQL)


v2 = "'',''"
vsql = "insert into temp2 (c02,c05) values (" + v2 + ")"
Call EjecutarScript(vsql, pathDBMySQL)

v2 = "'Total recibido en letras: " + EnLetras(vtotal) + "',''"
vsql = "insert into temp2 (c02,c05) values (" + v2 + ")"

Call EjecutarScript(vsql, pathDBMySQL)


ro.Close
'rd.Close

End Sub

Function getLinea(irow As Integer)

Dim i6 As Integer
Dim vlin As String

With Me.dgCheques
    
            .Row = irow
                     
            .Col = 7
            vlin = vlin + "Nr.Ch: " + .Text
            
            .Col = 21
            vlin = vlin + " B.: " + .Text
            
            .Col = 10
            vlin = vlin + " F.Acr.: " + .Text
            
           ' .Col = 30
            ' vlin = vlin + " - " + .Text
            

End With

getLinea = vlin

End Function


Function getLinea2(irow As Integer)


Dim vlin As String

With Me.dgCheques
    
        .Row = irow
        
        .Col = 11
            
        getLinea2 = .Text
                    
End With

End Function




Private Sub llenarDrRecibo2()

'llenar_pago1

'llenar_Pago2

'llenear_pago

End Sub




Private Sub llenarDrRecibo()
On Error Resume Next

Dim vvsaldo As Double
vvsaldo = Format(CalSaldoPersona(Me.txtCliente(0).Text, CP.TablaCtaCte), "$###,###,##0.00")
 
 
 Call ocultarEtiquetas
 
 With drRecibo
    
        .Sections(2).Controls("enrocomprobante").Caption = Str(vnrorecibo)
        If cpInstancia = "cobro" Then
                 .Sections(2).Controls("etipo").Caption = "RECIBO"
                 .Sections("TituloEmpresa").Controls("etiqueta1").Caption = "Recibo"
        Else
                .Sections("TituloEmpresa").Controls("etiqueta1").Caption = "ORDEN DE PAGO"
                .Sections(2).Controls("etipo").Caption = "Entrego a:"
        End If
        
        .Sections(2).Controls("enroasociado").Caption = Trim(vnroOrdenPago)
        .Sections(2).Controls("econcepto").Caption = Trim(Me.vconcepto)
        .Sections(2).Controls("etiqueta9").Caption = vfechaCredito.Value
        .Sections(2).Controls("lbllugar").Caption = vDatosEmpresa.Localidad & ", "
        .Sections(2).Controls("lblfecha").Caption = Date
        .Sections(2).Controls("lblCliente").Caption = Trim(txtCliente(0).Text) + " - " + Trim(txtCliente(1).Text)
        
        If Me.esComprobanteAutomatico Then
            .Sections(5).Controls("lblconcepto").Caption = Me.txtObservaciones
        Else
            .Sections(5).Controls("lblconcepto").Caption = Me.txtObservaciones 'txtTipoComp.Text & " " & txtNroComprobante.Text
        End If
        
        .Sections(5).Controls("lbltotal").Caption = Format(Me.TxtTotalAPagar.Text, "$###,###,##0.00")
        .Sections(5).Controls("eletras").Caption = EnLetras(Val(Me.TxtTotalAPagar.Text))
        
        .Sections(5).Controls("esaldo").Caption = Format(vvsaldo, "$ ###,###,##0.00")
        .Hide
    
        If Not vDraft Then
            .Sections(2).Controls("enrorecibo").Caption = Str(getNroRecibo)
        Else
             .Sections(2).Controls("enrorecibo").Caption = "No definitivo"
        End If
        
        vDraft = False
    End With



If Err < 0 Then
    MsgBox "Error al intentar hacer el recibo" + Str$(Err)
    Exit Sub
End If

End Sub


Private Sub PagarCtaCte(vnroremito As Long, importe As Double, idMedioPago As Integer) 'Este metodo estaba en Remito
    On Error Resume Next
    
    Dim rsCtaCteC As New ADODB.Recordset, sqlCtaCteC As String, vTipoComprobante As String, vnrocomprobante As String
    
    sqlCtaCteC = "SELECT * FROM CuentasCorrientes WHERE (remito = " & vnroremito & ")"
     
    With rsCtaCteC
        .CursorLocation = adUseClient
               
        Call .Open(sqlCtaCteC, ConnDDBB, adOpenStatic, adLockPessimistic)
        Dim SaldoAnterior, debito, credito As Double
        Do While Not .EOF
            If IsNull(.Fields("debito").Value) Then
                debito = 0
            Else
                debito = .Fields("debito").Value
            End If
            
            If IsNull(.Fields("credito").Value) Then
                credito = 0
            Else
                credito = .Fields("credito").Value
            End If
                        
            SaldoAnterior = Val(SaldoAnterior) + Val(debito) - Val(credito)
            
            .MoveNext
        Loop
        
        vTipoComprobante = TraerDato("Factura", "Remito = " & vnroremito & "", "Tipo")
        vnrocomprobante = TraerDato("Factura", "Remito = " & vnroremito & "", "NComprobante")
        
        If .RecordCount > 0 Then .MoveLast
        

        .AddNew
        .Fields("remito").Value = Trim(vnroremito)
        .Fields("comentario").Value = "Pago: Nro. " & vTipoComprobante & " " & Trim(vnrocomprobante)
        
        .Fields("Fecha").Value = strfechaMySQL(dtpFecha.Value)
        .Fields("Fechainput").Value = strfechaMySQL(dtpFecha.Value)
        
        'Buscar el cliente segun el remito
        .Fields("Codigo").Value = ObtenerCodigoClienteDesdeCtaCte(vnroremito)
        .Fields("Nombre").Value = ObtenerNombreClienteDesdeCtaCte(vnroremito)
       
        .Fields("anomes").Value = Right(.Fields("Fecha").Value, 4) & Mid(.Fields("Fecha").Value, 4, 2)
    
        .Fields("idMedioPago") = idMedioPago

        If (vTipoComprobante = "Documento") Or (vTipoComprobante = "Fact A") Or (vTipoComprobante = "Fact B") Then
            
            .Fields("debito") = 0
            .Fields("credito") = importe
            .Fields("saldo") = SaldoAnterior - .Fields("credito") 'bclientes.Recordset("saldo") + bfactura_temp.Recordset("Total")
                    
        End If
        
        .Update
        
        vIdCtaCteC = .Fields("id").Value
        
    End With

If Err Then GrabarLog "PagarCtaCte", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub PagarCtaCte2(vnroremito As Long, importe As Double, fecha As Date, idMedioPago)
    On Error Resume Next
    
    Dim rsCtaCteC As New ADODB.Recordset, sqlCtaCteC As String
    Dim SaldoAnterior, debito, credito As Double
    Dim comentario, TipoDocumento As String
    
    sqlCtaCteC = "SELECT * FROM CuentasCorrientes WHERE (remito = " & vnroremito & ")"
    
    With rsCtaCteC
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
               
        Call .Open(sqlCtaCteC, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If .State = 1 And Not .EOF = True Then
                
            .MoveFirst
        
            comentario = EsNulo(.Fields("comentario").Value)
        
            'TipoDocumento = .Fields("tipodocumento")
        
            Do While Not .EOF
                If IsNull(.Fields("debito").Value) Then
                    debito = 0
                Else
                    debito = Val(Format(.Fields("debito"), "#####0.00"))
                End If
            
                If IsNull(.Fields("credito").Value) Or (.Fields("credito").Value = 0) Then
                    credito = 0
                Else
                    credito = Val(Format(.Fields("credito"), "#####0.00"))
                End If
                        
                SaldoAnterior = Val(Format(SaldoAnterior, "#####0.00")) + Val(Format(debito, "#####0.00")) - Val(Format(credito, "#####0.00"))
            
                .MoveNext
            Loop
        
            If .RecordCount > 0 Then .MoveLast
        
            'If .EOF = True Then
            .AddNew
            .Fields("remito").Value = Trim(vnroremito)
            .Fields("comentario").Value = "Pago : " & comentario
            'End If
        
            .Fields("Fecha").Value = fecha
            .Fields("Fechainput").Value = fecha
        
            'Buscar el cliente segun el remito
            .Fields("Codigo").Value = ObtenerCodigoClienteDesdeCtaCte(vnroremito)
            .Fields("Nombre").Value = ObtenerNombreClienteDesdeCtaCte(vnroremito)
       
            .Fields("anomes").Value = Right(.Fields("Fecha").Value, 4) & Mid(.Fields("Fecha").Value, 4, 2)
    
            .Fields("debito") = 0 'Val(Format(.Fields("debito"), "#####0.00")) - Val(Format(importe, "#####0.00"))
            .Fields("credito") = Val(Format(importe, "#####0.00"))
            .Fields("saldo") = SaldoAnterior - Val(Format(.Fields("credito").Value, "#####0.00"))
                            
            .Fields("idMedioPago") = idMedioPago
            .Fields("TipoMovimiento").Value = "RC"
                            
            Select Case idMedioPago
    
                Case 1, 2, 5, 11, 12
                    .Fields("NroInterno").Value = Val(txtNroInterno.Text)
                
                
                Case 3
                Case 4
                    .Fields("NroInterno").Value = Val(txtNroInternoCheque.Text)
                
                Case 8
                
            End Select
            
            .Update
        
            vIdCtaCteC = Val(.Fields(0).Value)
        End If
    
    End With

    sqlCtaCteC = ""

    If rsCtaCteC.State = 1 Then
        rsCtaCteC.Close
        Set rsCtaCteC = Nothing
    End If

If Err Then GrabarLog "PagarCtaCte", Err.Number & " " & Err.Description, "Global"
End Sub
Private Sub ImprimirHasar(vremito As Long, vMontoEnEF As Double)
On Error Resume Next

    Dim FS As String
    
    vImpresionCorrecta = False

    FS = Chr$(28) '// Separador de campos del comando

    Dim rsImprimirHasar As New ADODB.Recordset, sqlImprimirHasar As String
    
    MsgBox "Prepare la Impresora ", vbInformation, "Mensaje ..."
    
    sqlImprimirHasar = "SELECT * FROM ImpresionFactura WHERE (Remito = " & Val(vremito) & ")"

    With rsImprimirHasar
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        
        Call .Open(sqlImprimirHasar, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If .State = 1 Then
            If Not .EOF = True Then
                'Va todo bien
            Else
                MsgBox "Remito Nro : " & remito
                Exit Sub
            End If
        Else
            MsgBox "No se pudo abrir la Factura de Venta : ", vbCritical, "Mensaje ..."
        End If
    End With
    
    With frmPrincipal.FiscalHasar
        'Call .EspecificarNombreDeFantasia(" ", " ")
        .Encabezado(1) = EsNulo(UCase(vDatosEmpresa.Nombre))
        .Encabezado(2) = EsNulo(UCase(vDatosEmpresa.Direccion))
        .Encabezado(3) = EsNulo(UCase(vDatosEmpresa.Localidad))
        .Encabezado(4) = EsNulo(UCase(vDatosEmpresa.CondicionIva)) & "            " & EsNulo(UCase(vDatosEmpresa.cuit))
        .Encabezado(5) = EsNulo(UCase(vDatosEmpresa.Telefono))
        
        Select Case EsNulo(rsImprimirHasar.Fields("TipoIva").Value)
            
            Case "Iva Responsable Inscripto"
                .PrecioBase = True
                Call .DatosCliente(EsNulo(rsImprimirHasar.Fields("Nombre").Value), Replace(rsImprimirHasar.Fields("Cuit").Value, "-", ""), TIPO_CUIT, RESPONSABLE_INSCRIPTO, Left(EsNulo(rsImprimirHasar.Fields("CodigoPostal").Value) & "-" & EsNulo(rsImprimirHasar.Fields("Localidad").Value) & "     -     " & EsNulo(rsImprimirHasar.Fields("Direccion").Value), 50))
                
            Case "Responsable Monotributo"
                .PrecioBase = False
                Call .DatosCliente(rsImprimirHasar.Fields("Nombre").Value, Replace(rsImprimirHasar.Fields("Cuit").Value, "-", ""), TIPO_CUIT, MONOTRIBUTO, Left(EsNulo(rsImprimirHasar.Fields("CodigoPostal").Value) & "-" & EsNulo(rsImprimirHasar.Fields("Localidad").Value) & "     -     " & EsNulo(rsImprimirHasar.Fields("Direccion").Value), 50))
            
            Case "Iva Exento"
                .PrecioBase = False
                Call .DatosCliente(rsImprimirHasar.Fields("Nombre").Value, Replace(rsImprimirHasar.Fields("Cuit").Value, "-", ""), TIPO_CUIT, RESPONSABLE_EXENTO, Left(EsNulo(rsImprimirHasar.Fields("CodigoPostal").Value) & "-" & EsNulo(rsImprimirHasar.Fields("Localidad").Value) & "     -     " & EsNulo(rsImprimirHasar.Fields("Direccion").Value), 50))
            
            Case "Consumidor Final"
                .PrecioBase = False
                Call .DatosCliente(EsNulo(rsImprimirHasar.Fields("Nombre").Value), EsNulo(rsImprimirHasar.Fields("NroDocumento").Value), TIPO_DNI, CONSUMIDOR_FINAL, Left(EsNulo(rsImprimirHasar.Fields("CodigoPostal").Value) & "-" & EsNulo(rsImprimirHasar.Fields("Localidad").Value) & "     -     " & EsNulo(rsImprimirHasar.Fields("Direccion").Value), 50))
            
        End Select
        
   
        Select Case EsNulo(rsImprimirHasar.Fields("Tipo").Value)
        
            Case "Fact A"
                Call .AbrirComprobanteFiscal(FACTURA_A)
            
            Case "Ticket-Factura"
                Call .AbrirComprobanteFiscal(TICKET_FACTURA_A)
            
            Case "Fact B"
                Call .AbrirComprobanteFiscal(FACTURA_B)
            
            Case "Documento"
                Call .AbrirComprobanteNoFiscal

            Case "Nota C"
                If EsNulo(rsImprimirHasar.Fields("TipoIva").Value) = "001" Then
                    .AbrirComprobanteNoFiscalHomologado (NOTA_CREDITO_A)
                Else
                    .AbrirComprobanteNoFiscalHomologado (NOTA_CREDITO_B)
                End If
                
            Case "Remito"
                '
                
        End Select
        
        Dim rsDetalleHasar As New ADODB.Recordset, sqlDetalleHasar As String
        
        sqlDetalleHasar = "SELECT * FROM FDetalle WHERE (Remito = " & Val(vremito) & ") ORDER BY idFDetalle ASC"
        
        If rsDetalleHasar.State = 1 Then rsDetalleHasar.Close
        rsDetalleHasar.CursorLocation = adUseClient
        
        Call rsDetalleHasar.Open(sqlDetalleHasar, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not rsDetalleHasar.State = 1 Then
            If rsDetalleHasar.EOF = True Then
            
            End If
        End If
    
        Do Until rsDetalleHasar.EOF = True
            
            Select Case EsNulo(rsImprimirHasar.Fields("Tipo").Value)
        
                Case "Fact A", "Ticket-Factura", "Fact B", "Nota C"
                    Call .ImprimirItem(EsNulo(rsDetalleHasar.Fields("Detalle").Value), Val(rsDetalleHasar.Fields("Cantidad").Value), Val(rsDetalleHasar.Fields("Precio").Value), Val(rsDetalleHasar.Fields("TIVa").Value), 0)

                Case "Documento"
                    Call .ImprimirTextoNoFiscal(rsDetalleHasar.Fields("Detalle").Value)
        
            End Select
            
            rsDetalleHasar.MoveNext
        Loop

        '.DescuentoUltimoItem "Oferta del Dia", 5, True
        '.DescuentoGeneral "Oferta Pago Efectivo", 25, True
        '.EspecificarPercepcionPorIVA "Percep IVA21", 100, 21
        '.EspecificarPercepcionGlobal "Percep. RG 0000", 125#

        'Imprimir Comentarios
        Call .ImprimirPago("Efectivo", vMontoEnEF)  'Val(GenerarDato("SELECT SUM(Monto) AS TotalEF FROM Recibo_Temp WHERE IdMedioPago = 1 GROUP BY idMedioPago;", "TotalEF")))
        
        Call ImprimirComentariosFacturaHasar
        
        Select Case EsNulo(rsImprimirHasar.Fields("Tipo").Value)
        
            Case "Fact A", "Ticket-Factura", "Fact B"
                Call .CerrarComprobanteFiscal

            Case "Documento"
                Call .CerrarComprobanteNoFiscal
                
            Case "Nota C"
                Call .CerrarComprobanteNoFiscalHomologado
        
        End Select
        
        '.Finalizar
    End With
    
    
If Err Then
    Call GrabarLog("ImprimirHasar", Err.Number & " " & Err.Description, Me.Caption)
    Call MsgBox("Error Impresora:" & Err.Description, vbCritical, "Errores")
Else
    vImpresionCorrecta = True
End If
End Sub
Private Sub ImprimirComentariosFacturaHasar()
On Error Resume Next

    Dim rsComentariosFactura As New ADODB.Recordset, sqlComentariosFactura As String, l As Integer
    
    sqlComentariosFactura = "SELECT * FROM ComentariosFactura LIMIT 0,4"
    
    'No Tocar esto
    l = 11
    With rsComentariosFactura
        Call .Open(sqlComentariosFactura, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then .MoveFirst
        
        For l = 11 To 14
            If .Fields("Imprimir").Value = "S" Then
                frmPrincipal.FiscalHasar.Encabezado(l) = EsNulo(Left(.Fields("Comentario").Value, 50))
            Else
                frmPrincipal.FiscalHasar.Encabezado(l) = EsNulo(" ")
            End If
            
            .MoveNext
        Next
    
    End With

    sqlComentariosFactura = ""

    If rsComentariosFactura.State = 1 Then
        rsComentariosFactura.Close
        Set rsComentariosFactura = Nothing
    End If
    
If Err Then GrabarLog "ImprimirComentariosFacturaHasar", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub TxtTotalAPagar_Change()
On Error Resume Next

    If Val(TxtTotalAPagar.Text) > Val(txtMontoTotalPendienteSeleccionado.Text) Then
       'TxtTotalAPagar.Back Color = vbRed
    Else
       'TxtTotalAPagar.BackColor = vbWhite
    End If
    
    If Not Val(TxtTotalAPagar.Text) = 0 Then vImporteTotalAPagar = Val(TxtTotalAPagar.Text)
    
    TxtTotalAPagar.Tag = TxtTotalAPagar.Text
    
   ' TxtTotalAPagar.Text = Format(TxtTotalAPagar, "###,###,##0.00")
        
If Err Then GrabarLog "TxtTotalAPagar_Change", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub vbc1_Change()
Me.txtBancoCheque(1).Tag = Me.vbc1.Tag
Me.txtBancoCheque(1).Text = Me.vbc1.Text
End Sub

Private Sub vCCajaDestino_Change()
vCCajaDestino.Tag = traerDatos2("select * from bancos where idbancos='" + vCCajaDestino.Text + "'", "id", pathDBMySQL)
End Sub

Private Sub vDCajaDestino_Change()
Me.vCCajaDestino.Text = Me.vDCajaDestino.Tag
End Sub

Private Sub vdretencion_Change()
Me.vCretencion.Text = Me.vdretencion.Tag
Me.vImporteRet.SetFocus
End Sub

Private Sub vfechaCredito_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
   'Me.txtImporteEfectivo.SetFocus
End If
If Err Then Exit Sub
End Sub

Private Sub vfechaCredito_LostFocus()
If vfechaCredito.Value < CDate("01/01/2000") Or vfechaCredito.Text = "" Then
    MsgBox "Fecha ingresada inválida. Se fijará la fecha actual", vbCritical
    vfechaCredito.Value = Date
End If

End Sub

Private Sub vfechaCheque_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  '  txtNroInternoCheque.SetFocus
End If
End Sub

Private Sub vfechaCheque_LostFocus()
'vfechaCredito = vfechaCheque
End Sub

Private Sub vfechaDeposito_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  '  txtFechaDeposito.SetFocus
End If
End Sub

Private Sub vfechaDeposito_LostFocus()
vfechaCredito = vfechaDeposito
If txtFechaDeposito.Value = Me.vfechaDeposito.Value Then
    MsgBox "Le advertimos que las fechas de depósito y de movimiento son iguales", vbInformation
End If
End Sub

Private Sub vImporteRet_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  '  Me.PushButton3.SetFocus
End If
End Sub

Private Sub vmarcainterna_GotFocus()
Call PushButton5_Click
End Sub

Private Sub vsaldo_Click()
vsaldo.Caption = Format(CalSaldoPersona(Me.txtCliente(0).Text, CP.TablaCtaCte), "$###,###,##0.00")
End Sub

Private Sub vtimporteRetenciones_Change()
calTotales
End Sub


Public Sub limpiarControles(d As Integer, H As Integer)
'Variable de tipo Control Para los controles del contenedor en este caso del Frame
Dim ElControl As control
     
    'recorre los controles
      
    For Each ElControl In Controls
        'si está dentro lo deshabilita
        If ElControl.TabIndex >= d And ElControl.TabIndex <= H Then
           Call vaciarControl(ElControl)
        End If
    Next
End Sub
