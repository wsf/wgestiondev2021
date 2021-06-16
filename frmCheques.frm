VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "Copia de KlexGrid.ocx"
Object = "{9746E3DA-06E1-4D26-9CE4-D9F6411A9C70}#1.0#0"; "SMGA_OcxTxt2008.ocx"
Begin VB.Form frmCheques 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Cheques"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   17385
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   17385
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   615
      Left            =   0
      TabIndex        =   134
      Top             =   -120
      Width           =   17340
      _Version        =   851968
      _ExtentX        =   30586
      _ExtentY        =   1085
      _StockProps     =   79
      BackColor       =   -2147483638
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   375
         Left            =   15600
         TabIndex        =   182
         Top             =   180
         Width           =   675
         _Version        =   851968
         _ExtentX        =   1191
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Excel"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox4 
         Height          =   120
         Left            =   60
         TabIndex        =   179
         Top             =   1770
         Width           =   17325
         _Version        =   851968
         _ExtentX        =   30559
         _ExtentY        =   212
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   375
         Left            =   9810
         TabIndex        =   154
         Top             =   180
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3096
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Ver Transacciones"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCheques.frx":0000
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   375
         Left            =   3360
         TabIndex        =   152
         Top             =   180
         Width           =   885
         _Version        =   851968
         _ExtentX        =   1561
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCheques.frx":059A
         BorderGap       =   1
      End
      Begin XtremeSuiteControls.PushButton cmdAcciones 
         Height          =   375
         Index           =   2
         Left            =   14340
         TabIndex        =   135
         Top             =   180
         Visible         =   0   'False
         Width           =   945
         _Version        =   851968
         _ExtentX        =   1667
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCheques.frx":0B34
      End
      Begin XtremeSuiteControls.PushButton cmdAcciones 
         Height          =   375
         Index           =   4
         Left            =   7110
         TabIndex        =   136
         Top             =   180
         Width           =   1515
         _Version        =   851968
         _ExtentX        =   2672
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cambiar Estado"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCheques.frx":0F0F
      End
      Begin XtremeSuiteControls.PushButton cmdAcciones 
         Height          =   375
         Index           =   5
         Left            =   8610
         TabIndex        =   137
         Top             =   180
         Width           =   1185
         _Version        =   851968
         _ExtentX        =   2090
         _ExtentY        =   661
         _StockProps     =   79
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
         BorderGap       =   10
      End
      Begin XtremeSuiteControls.PushButton cmdAcciones 
         Height          =   375
         Index           =   6
         Left            =   16530
         TabIndex        =   138
         Top             =   180
         Width           =   825
         _Version        =   851968
         _ExtentX        =   1455
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cerrar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCheques.frx":14A9
      End
      Begin XtremeSuiteControls.PushButton cmdAcciones 
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   139
         Top             =   180
         Width           =   1065
         _Version        =   851968
         _ExtentX        =   1879
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Modificar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCheques.frx":18A9
      End
      Begin XtremeSuiteControls.PushButton cmdAcciones 
         Height          =   375
         Index           =   1
         Left            =   2280
         TabIndex        =   140
         Top             =   180
         Width           =   825
         _Version        =   851968
         _ExtentX        =   1455
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Borrar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCheques.frx":1E43
      End
      Begin XtremeSuiteControls.PushButton cmdAcciones 
         Height          =   375
         Index           =   3
         Left            =   6060
         TabIndex        =   141
         Top             =   180
         Visible         =   0   'False
         Width           =   1035
         _Version        =   851968
         _ExtentX        =   1826
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Diferidos"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCheques.frx":23DD
      End
      Begin XtremeSuiteControls.PushButton PusVerHistorial 
         Height          =   375
         Index           =   7
         Left            =   4770
         TabIndex        =   142
         Top             =   180
         Width           =   1245
         _Version        =   851968
         _ExtentX        =   2196
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Ver Historial"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCheques.frx":27B8
      End
      Begin XtremeSuiteControls.PushButton cmdAcciones 
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   153
         Top             =   180
         Width           =   1065
         _Version        =   851968
         _ExtentX        =   1879
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Nuevo"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCheques.frx":2D52
      End
      Begin XtremeSuiteControls.PushButton cmdAcciones 
         Height          =   375
         Index           =   8
         Left            =   12030
         TabIndex        =   168
         Top             =   180
         Width           =   1845
         _Version        =   851968
         _ExtentX        =   3254
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Historial del Cheque"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCheques.frx":32EC
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox5 
      Height          =   195
      Left            =   -30
      TabIndex        =   180
      Top             =   420
      Width           =   17385
      _Version        =   851968
      _ExtentX        =   30665
      _ExtentY        =   344
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
   End
   Begin XtremeSuiteControls.GroupBox GbEstado 
      Height          =   4815
      Left            =   2730
      TabIndex        =   63
      Top             =   6810
      Visible         =   0   'False
      Width           =   7005
      _Version        =   851968
      _ExtentX        =   12356
      _ExtentY        =   8493
      _StockProps     =   79
      Caption         =   "[Cambiar Estado del Cheque]"
      ForeColor       =   0
      BackColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   135
         Left            =   90
         TabIndex        =   178
         Top             =   4110
         Width           =   6855
         _Version        =   851968
         _ExtentX        =   12091
         _ExtentY        =   238
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
      End
      Begin VB.TextBox txtNroCheque 
         Height          =   285
         Left            =   120
         TabIndex        =   167
         Text            =   "Nro cheque"
         Top             =   3150
         Width           =   2235
      End
      Begin VB.PictureBox txtNroCheque1 
         Height          =   315
         Left            =   150
         ScaleHeight     =   255
         ScaleWidth      =   2115
         TabIndex        =   166
         Top             =   3480
         Width           =   2175
      End
      Begin VB.ComboBox vtipoOperacion 
         Height          =   315
         ItemData        =   "frmCheques.frx":3886
         Left            =   2610
         List            =   "frmCheques.frx":388D
         TabIndex        =   147
         Top             =   1350
         Width           =   4095
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   150
         ScaleHeight     =   555
         ScaleWidth      =   6795
         TabIndex        =   124
         Top             =   270
         Width           =   6795
         Begin VB.Label lblNroCheque 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   2430
            TabIndex        =   126
            Top             =   90
            Width           =   4050
         End
         Begin VB.Label lblCambiarEstado 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "> Nro de Cheque seleccionado : "
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   -600
            TabIndex        =   125
            Top             =   150
            Width           =   2985
         End
      End
      Begin MSComCtl2.DTPicker dtpCambioFecha 
         Height          =   315
         Left            =   2610
         TabIndex        =   65
         Top             =   900
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   556
         _Version        =   393216
         Format          =   25755649
         CurrentDate     =   40506
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   1005
         Index           =   14
         Left            =   2610
         TabIndex        =   122
         Top             =   2640
         Width           =   4095
         _Version        =   851968
         _ExtentX        =   7223
         _ExtentY        =   1773
         _StockProps     =   77
         BackColor       =   -2147483643
         MaxLength       =   250
         ScrollBars      =   2
      End
      Begin XtremeSuiteControls.ComboBox vpropietario 
         Height          =   315
         Left            =   2610
         TabIndex        =   143
         Top             =   2190
         Width           =   4095
         _Version        =   851968
         _ExtentX        =   7223
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.ComboBox cboCambiarEstado 
         Height          =   315
         Left            =   2610
         TabIndex        =   145
         Top             =   1770
         Width           =   4095
         _Version        =   851968
         _ExtentX        =   7223
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton cmdCambiarEstado 
         Height          =   405
         Index           =   1
         Left            =   4170
         TabIndex        =   159
         Top             =   4290
         Width           =   1470
         _Version        =   851968
         _ExtentX        =   2593
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Cambiar Todos"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCheques.frx":38AF
      End
      Begin XtremeSuiteControls.PushButton cmdCambiarEstado 
         Height          =   405
         Index           =   2
         Left            =   5670
         TabIndex        =   160
         Top             =   4290
         Width           =   1200
         _Version        =   851968
         _ExtentX        =   2117
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Volver"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCheques.frx":3CC9
      End
      Begin XtremeSuiteControls.PushButton cmdCambiarEstado 
         Height          =   405
         Index           =   0
         Left            =   2640
         TabIndex        =   161
         Top             =   4290
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Ejecutar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCheques.frx":40E9
      End
      Begin VB.Label lblTipoDe 
         Caption         =   "Tipo de Operación:"
         Height          =   225
         Left            =   1170
         TabIndex        =   148
         Top             =   1350
         Width           =   1365
      End
      Begin VB.Label lblCambiarEstado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "> Cambiar estado del cheque a:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   146
         Top             =   1800
         Width           =   2355
      End
      Begin VB.Label lblCambiarEstado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "> Entregado a:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   144
         Top             =   2220
         Width           =   2355
      End
      Begin VB.Label lblFicha 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> Observaciones:"
         Height          =   195
         Index           =   10
         Left            =   1200
         TabIndex        =   123
         Top             =   2640
         Width           =   1275
      End
      Begin VB.Label lblCambiarEstado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "> Fecha del Cambio de Estado:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   64
         Top             =   990
         Width           =   2475
      End
   End
   Begin XtremeSuiteControls.TabControl TabBusqueda 
      Height          =   6435
      Left            =   60
      TabIndex        =   66
      Top             =   630
      Width           =   17325
      _Version        =   851968
      _ExtentX        =   30559
      _ExtentY        =   11351
      _StockProps     =   68
      AllowReorder    =   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      ItemCount       =   6
      SelectedItem    =   2
      Item(0).Caption =   "Filtar por:"
      Item(0).ControlCount=   47
      Item(0).Control(0)=   "txtBuscar(2)"
      Item(0).Control(1)=   "lblSituación(2)"
      Item(0).Control(2)=   "pbCarga(1)"
      Item(0).Control(3)=   "txtFicha(2)"
      Item(0).Control(4)=   "txtFicha(0)"
      Item(0).Control(5)=   "txtFicha(1)"
      Item(0).Control(6)=   "txtFicha(3)"
      Item(0).Control(7)=   "txtFicha(5)"
      Item(0).Control(8)=   "txtFicha(6)"
      Item(0).Control(9)=   "txtFicha(8)"
      Item(0).Control(10)=   "pbCarga(2)"
      Item(0).Control(11)=   "pbCarga(3)"
      Item(0).Control(12)=   "txtFicha(7)"
      Item(0).Control(13)=   "txtFicha(9)"
      Item(0).Control(14)=   "pbCarga(4)"
      Item(0).Control(15)=   "txtFicha(10)"
      Item(0).Control(16)=   "txtFicha(4)"
      Item(0).Control(17)=   "lblFicha(6)"
      Item(0).Control(18)=   "lblFicha(5)"
      Item(0).Control(19)=   "lblFicha(4)"
      Item(0).Control(20)=   "lblFicha(3)"
      Item(0).Control(21)=   "lblFicha(2)"
      Item(0).Control(22)=   "lblFicha(1)"
      Item(0).Control(23)=   "lblFicha(0)"
      Item(0).Control(24)=   "vFirmante"
      Item(0).Control(25)=   "lblFicha(7)"
      Item(0).Control(26)=   "FraFechaDe"
      Item(0).Control(27)=   "Frame2"
      Item(0).Control(28)=   "Frame5"
      Item(0).Control(29)=   "chkActivarFecha"
      Item(0).Control(30)=   "chkActivarFDeposito"
      Item(0).Control(31)=   "PusLimpiarTodos"
      Item(0).Control(32)=   "lblSituación(0)"
      Item(0).Control(33)=   "vnrocheque"
      Item(0).Control(34)=   "lblSituación(1)"
      Item(0).Control(35)=   "vmarcainterna"
      Item(0).Control(36)=   "vdcustodia"
      Item(0).Control(37)=   "lblSituación(3)"
      Item(0).Control(38)=   "custodia(0)"
      Item(0).Control(39)=   "vccustodia"
      Item(0).Control(40)=   "vmarcainternaHasta"
      Item(0).Control(41)=   "lblSituación(4)"
      Item(0).Control(42)=   "lblSituación(5)"
      Item(0).Control(43)=   "GroupBox1"
      Item(0).Control(44)=   "chkChkSinCustodia"
      Item(0).Control(45)=   "vendoso"
      Item(0).Control(46)=   "lblFicha(8)"
      Item(1).Caption =   "Tipo de Listado"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "Frame6"
      Item(1).Control(1)=   "Frame9"
      Item(2).Caption =   "Ver Datos"
      Item(2).ControlCount=   4
      Item(2).Control(0)=   "KlexCheques"
      Item(2).Control(1)=   "vbusca"
      Item(2).Control(2)=   "Label27"
      Item(2).Control(3)=   "chknroExacto"
      Item(3).Caption =   "Historial del cheque"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "KlexGrid1"
      Item(4).Caption =   "Consulta Directa"
      Item(4).ControlCount=   0
      Item(5).Caption =   "Estadísticas"
      Item(5).ControlCount=   1
      Item(5).Control(0)=   "Frame11"
      Begin XtremeSuiteControls.CheckBox chknroExacto 
         Height          =   285
         Left            =   14910
         TabIndex        =   185
         Top             =   450
         Width           =   1785
         _Version        =   851968
         _ExtentX        =   3149
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Buscar nro exacto"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkChkSinCustodia 
         Height          =   255
         Left            =   -60670
         TabIndex        =   181
         Top             =   4020
         Visible         =   0   'False
         Width           =   2595
         _Version        =   851968
         _ExtentX        =   4577
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cheques sin custodia"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   555
         Left            =   -63010
         TabIndex        =   174
         Top             =   5340
         Visible         =   0   'False
         Width           =   4485
         _Version        =   851968
         _ExtentX        =   7911
         _ExtentY        =   979
         _StockProps     =   79
         Caption         =   "Tipos de cheques"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton rdterceros 
            Height          =   225
            Left            =   210
            TabIndex        =   175
            Top             =   240
            Width           =   1065
            _Version        =   851968
            _ExtentX        =   1879
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "Terceros"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbpropios 
            Height          =   225
            Left            =   1830
            TabIndex        =   176
            Tag             =   "Propio"
            Top             =   240
            Width           =   1065
            _Version        =   851968
            _ExtentX        =   1879
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "Propios"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbtodos 
            Height          =   225
            Left            =   3240
            TabIndex        =   177
            Top             =   270
            Width           =   1065
            _Version        =   851968
            _ExtentX        =   1879
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "Todos"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.FlatEdit vmarcainterna 
         Height          =   285
         Left            =   -66580
         TabIndex        =   158
         Top             =   3660
         Visible         =   0   'False
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3096
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton PusLimpiarTodos 
         Height          =   255
         Left            =   -61090
         TabIndex        =   131
         Top             =   4320
         Visible         =   0   'False
         Width           =   2595
         _Version        =   851968
         _ExtentX        =   4577
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Limpiar todos los campos"
         Appearance      =   3
         Picture         =   "frmCheques.frx":44EB
      End
      Begin VB.Frame Frame11 
         Height          =   645
         Left            =   -69460
         TabIndex        =   117
         Top             =   570
         Visible         =   0   'False
         Width           =   10305
         Begin XtremeSuiteControls.ComboBox ComboBox1 
            Height          =   315
            Left            =   780
            TabIndex        =   118
            Top             =   240
            Width           =   3735
            _Version        =   851968
            _ExtentX        =   6588
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            Text            =   "Listado"
         End
         Begin XtremeSuiteControls.ComboBox ComboBox2 
            Height          =   315
            Left            =   5730
            TabIndex        =   119
            Top             =   240
            Width           =   4365
            _Version        =   851968
            _ExtentX        =   7699
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            Text            =   "Listado"
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo:"
            Height          =   225
            Left            =   180
            TabIndex        =   121
            Top             =   270
            Width           =   735
         End
         Begin VB.Label lblTipo 
            BackStyle       =   0  'Transparent
            Caption         =   "Presentación:"
            Height          =   225
            Left            =   4620
            TabIndex        =   120
            Top             =   300
            Width           =   1005
         End
      End
      Begin VB.Frame Frame9 
         Height          =   4335
         Left            =   -69910
         TabIndex        =   112
         Top             =   420
         Visible         =   0   'False
         Width           =   15495
         Begin XtremeSuiteControls.ComboBox CombOrdenamiento 
            Height          =   315
            Left            =   2460
            TabIndex        =   113
            Top             =   690
            Width           =   8775
            _Version        =   851968
            _ExtentX        =   15478
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            Text            =   "Listado"
         End
         Begin XtremeSuiteControls.ComboBox vorden 
            Height          =   315
            Left            =   2460
            TabIndex        =   127
            Top             =   1110
            Width           =   8745
            _Version        =   851968
            _ExtentX        =   15425
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            Text            =   "Fecha"
         End
         Begin XtremeSuiteControls.ComboBox vgrupo 
            Height          =   315
            Left            =   2460
            TabIndex        =   129
            Top             =   1500
            Width           =   8745
            _Version        =   851968
            _ExtentX        =   15425
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            Text            =   "Listado"
         End
         Begin XtremeSuiteControls.ComboBox ComListado 
            Height          =   315
            Left            =   2460
            TabIndex        =   149
            Top             =   270
            Width           =   8775
            _Version        =   851968
            _ExtentX        =   15478
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Style           =   2
            Text            =   "Listado"
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dónde se encuentra el cheque: "
            Height          =   225
            Left            =   -300
            TabIndex        =   150
            Top             =   330
            Width           =   2655
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Agrupado por el/ los campo/s:"
            Height          =   225
            Left            =   90
            TabIndex        =   130
            Top             =   1560
            Width           =   2175
         End
         Begin VB.Label lblOrdenadoPor 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ordenado por el campo:"
            Height          =   225
            Left            =   90
            TabIndex        =   128
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label lblOrdenamiento 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Detalles:"
            Height          =   225
            Left            =   1410
            TabIndex        =   114
            Top             =   750
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         Height          =   675
         Left            =   -69910
         TabIndex        =   108
         Top             =   4860
         Visible         =   0   'False
         Width           =   15525
         Begin XtremeSuiteControls.RadioButton RadDetallado 
            Height          =   315
            Left            =   5610
            TabIndex        =   109
            ToolTipText     =   "Muestra todo los cheques y todas las imputaciones en cada caja."
            Top             =   240
            Width           =   2685
            _Version        =   851968
            _ExtentX        =   4736
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Detallando el historial del cheque"
            Transparent     =   -1  'True
            Appearance      =   6
         End
         Begin XtremeSuiteControls.RadioButton RadResumido 
            Height          =   285
            Left            =   3480
            TabIndex        =   110
            ToolTipText     =   "Muestra todos los cheques que tienen asignada una caja."
            Top             =   240
            Width           =   1245
            _Version        =   851968
            _ExtentX        =   2196
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Resumido"
            Transparent     =   -1  'True
            Appearance      =   6
            Value           =   -1  'True
         End
         Begin VB.Label lblNivelesDe 
            BackStyle       =   0  'Transparent
            Caption         =   "Niveles de Detalles:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   90
            TabIndex        =   111
            Top             =   270
            Width           =   1785
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Importe del Cheque:"
         Height          =   525
         Left            =   -69940
         TabIndex        =   105
         Top             =   5580
         Visible         =   0   'False
         Width           =   6255
         Begin VB.TextBox vImporteHasta 
            Height          =   285
            Left            =   4920
            TabIndex        =   116
            Top             =   180
            Width           =   1215
         End
         Begin VB.TextBox vImporteDesde 
            Height          =   285
            Left            =   2670
            TabIndex        =   115
            Top             =   180
            Width           =   1215
         End
         Begin XtremeSuiteControls.Label lblBanco 
            Height          =   255
            Index           =   5
            Left            =   4200
            TabIndex        =   107
            Top             =   210
            Width           =   735
            _Version        =   851968
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "> Hasta:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblBanco 
            Height          =   255
            Index           =   4
            Left            =   1770
            TabIndex        =   106
            Top             =   180
            Width           =   795
            _Version        =   851968
            _ExtentX        =   1402
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "> Desde:"
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.CheckBox chkActivarFecha 
         Height          =   255
         Left            =   -63910
         TabIndex        =   104
         Top             =   4740
         Visible         =   0   'False
         Width           =   765
         _Version        =   851968
         _ExtentX        =   1349
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Activar"
         Transparent     =   -1  'True
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox chkActivarFDeposito 
         Height          =   255
         Left            =   -69670
         TabIndex        =   103
         Top             =   5250
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Activar"
         Transparent     =   -1  'True
         Appearance      =   6
      End
      Begin VB.Frame Frame2 
         Caption         =   "Fecha de acreditación: "
         Height          =   495
         Left            =   -69940
         TabIndex        =   98
         Top             =   5070
         Visible         =   0   'False
         Width           =   6255
         Begin Aplisoft_CajasDeTexto.TxF txtFecha 
            Height          =   285
            Index           =   2
            Left            =   2700
            TabIndex        =   99
            Top             =   150
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   503
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
         Begin Aplisoft_CajasDeTexto.TxF txtFecha 
            Height          =   285
            Index           =   3
            Left            =   4920
            TabIndex        =   100
            Top             =   150
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
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
            Index           =   1
            Left            =   1740
            TabIndex        =   102
            Top             =   180
            Width           =   795
            _Version        =   851968
            _ExtentX        =   1402
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "> Desde:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblBanco 
            Height          =   255
            Index           =   0
            Left            =   4230
            TabIndex        =   101
            Top             =   180
            Width           =   675
            _Version        =   851968
            _ExtentX        =   1191
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "> Hasta:"
            Transparent     =   -1  'True
         End
      End
      Begin VB.Frame FraFechaDe 
         Caption         =   "Fecha de entrada del cheque:"
         Height          =   495
         Left            =   -69940
         TabIndex        =   93
         Top             =   4560
         Visible         =   0   'False
         Width           =   11505
         Begin Aplisoft_CajasDeTexto.TxF txtFecha 
            Height          =   285
            Index           =   0
            Left            =   8190
            TabIndex        =   94
            Top             =   150
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
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
         Begin Aplisoft_CajasDeTexto.TxF txtFecha 
            Height          =   285
            Index           =   1
            Left            =   10230
            TabIndex        =   95
            Top             =   150
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
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
            Left            =   9480
            TabIndex        =   97
            Top             =   150
            Width           =   645
            _Version        =   851968
            _ExtentX        =   1138
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "> Hasta:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblBanco 
            Height          =   255
            Index           =   2
            Left            =   7320
            TabIndex        =   96
            Top             =   150
            Width           =   1125
            _Version        =   851968
            _ExtentX        =   1984
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "> Desde:"
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.ComboBox vFirmante 
         Height          =   315
         Left            =   -68290
         TabIndex        =   90
         Top             =   1500
         Visible         =   0   'False
         Width           =   3525
         _Version        =   851968
         _ExtentX        =   6218
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtBuscar 
         Height          =   285
         Index           =   2
         Left            =   -68290
         TabIndex        =   67
         Top             =   2970
         Visible         =   0   'False
         Width           =   9795
         _Version        =   851968
         _ExtentX        =   17277
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   1
         Left            =   -66970
         TabIndex        =   69
         Tag             =   "CodigoCliente"
         Top             =   450
         Visible         =   0   'False
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
         Index           =   2
         Left            =   -68290
         TabIndex        =   70
         Top             =   750
         Visible         =   0   'False
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   0
         Left            =   -68290
         TabIndex        =   71
         Top             =   420
         Visible         =   0   'False
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   1
         Left            =   -66580
         TabIndex        =   72
         Top             =   450
         Visible         =   0   'False
         Width           =   8145
         _Version        =   851968
         _ExtentX        =   14367
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   3
         Left            =   -66580
         TabIndex        =   73
         Top             =   780
         Visible         =   0   'False
         Width           =   8145
         _Version        =   851968
         _ExtentX        =   14367
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   6
         Left            =   -68290
         TabIndex        =   74
         Top             =   1860
         Visible         =   0   'False
         Width           =   9795
         _Version        =   851968
         _ExtentX        =   17277
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   8
         Left            =   -66490
         TabIndex        =   75
         Top             =   2220
         Visible         =   0   'False
         Width           =   7995
         _Version        =   851968
         _ExtentX        =   14102
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   285
         Index           =   2
         Left            =   -66970
         TabIndex        =   76
         Tag             =   "Proveedor"
         Top             =   810
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   3
         Left            =   -66970
         TabIndex        =   77
         Tag             =   "EstadoCheque"
         Top             =   2220
         Visible         =   0   'False
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
         Index           =   7
         Left            =   -68290
         TabIndex        =   78
         Top             =   2220
         Visible         =   0   'False
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   9
         Left            =   -68290
         TabIndex        =   79
         Top             =   2580
         Visible         =   0   'False
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   4
         Left            =   -66970
         TabIndex        =   80
         Tag             =   "Banco"
         Top             =   2580
         Visible         =   0   'False
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
         Index           =   10
         Left            =   -66490
         TabIndex        =   81
         Top             =   2580
         Visible         =   0   'False
         Width           =   7995
         _Version        =   851968
         _ExtentX        =   14102
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   4
         Left            =   -68290
         TabIndex        =   82
         Top             =   1140
         Visible         =   0   'False
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin Grid.KlexGrid KlexCheques 
         Height          =   5385
         Left            =   30
         TabIndex        =   132
         Top             =   780
         Width           =   17235
         _ExtentX        =   30401
         _ExtentY        =   9499
         EnterKeyBehaviour=   0
         BackColorAlternate=   0
         GridLines       =   0
         GridLinesFixed  =   2
         AllowUserResizing=   1
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
         MouseIcon       =   "frmCheques.frx":4A85
         MousePointer    =   1
         Rows            =   10
         SelectionMode   =   1
      End
      Begin Grid.KlexGrid KlexGrid1 
         Height          =   7605
         Left            =   -70000
         TabIndex        =   151
         Top             =   330
         Visible         =   0   'False
         Width           =   17265
         _ExtentX        =   30454
         _ExtentY        =   13414
         EnterKeyBehaviour=   0
         BackColorAlternate=   0
         GridLines       =   0
         GridLinesFixed  =   2
         AllowUserResizing=   1
         Appearance      =   0
         BackColorFixed  =   -2147483626
         Cols            =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColorFixed  =   8421504
         HighLight       =   0
         MouseIcon       =   "frmCheques.frx":4AA1
         MousePointer    =   1
         Rows            =   10
         SelectionMode   =   1
      End
      Begin XtremeSuiteControls.FlatEdit vnrocheque 
         Height          =   285
         Left            =   -68290
         TabIndex        =   156
         Top             =   3330
         Visible         =   0   'False
         Width           =   9795
         _Version        =   851968
         _ExtentX        =   17277
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   5
         Left            =   -64810
         TabIndex        =   91
         Top             =   1140
         Visible         =   0   'False
         Width           =   6345
         _Version        =   851968
         _ExtentX        =   11192
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vccustodia 
         Height          =   285
         Left            =   -68290
         TabIndex        =   162
         Top             =   3990
         Visible         =   0   'False
         Width           =   1275
         _Version        =   851968
         _ExtentX        =   2249
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vdcustodia 
         Height          =   285
         Left            =   -66610
         TabIndex        =   163
         Top             =   3990
         Visible         =   0   'False
         Width           =   5565
         _Version        =   851968
         _ExtentX        =   9816
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton custodia 
         Height          =   285
         Index           =   0
         Left            =   -66970
         TabIndex        =   165
         Tag             =   "CodigoCliente"
         Top             =   3990
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit vmarcainternaHasta 
         Height          =   285
         Left            =   -60700
         TabIndex        =   169
         Top             =   3660
         Visible         =   0   'False
         Width           =   2175
         _Version        =   851968
         _ExtentX        =   3836
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit vbusca 
         Height          =   315
         Left            =   3630
         TabIndex        =   172
         Top             =   390
         Width           =   10935
         _Version        =   851968
         _ExtentX        =   19288
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   3
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox vendoso 
         Height          =   315
         Left            =   -62350
         TabIndex        =   183
         Top             =   1500
         Visible         =   0   'False
         Width           =   3885
         _Version        =   851968
         _ExtentX        =   6853
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin VB.Label lblFicha 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Endoso automático"
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   8
         Left            =   -64540
         TabIndex        =   184
         Top             =   1560
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label27 
         Caption         =   "Buscar por marca interna o Nro comprobante: "
         Height          =   255
         Left            =   120
         TabIndex        =   173
         Top             =   450
         Width           =   3765
      End
      Begin XtremeSuiteControls.Label lblSituación 
         Height          =   225
         Index           =   5
         Left            =   -61240
         TabIndex        =   171
         Top             =   3690
         Visible         =   0   'False
         Width           =   585
         _Version        =   851968
         _ExtentX        =   1032
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "hasta:"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblSituación 
         Height          =   225
         Index           =   4
         Left            =   -67120
         TabIndex        =   170
         Top             =   3660
         Visible         =   0   'False
         Width           =   585
         _Version        =   851968
         _ExtentX        =   1032
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "desde:"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblSituación 
         Height          =   225
         Index           =   3
         Left            =   -69730
         TabIndex        =   164
         Top             =   4020
         Visible         =   0   'False
         Width           =   1065
         _Version        =   851968
         _ExtentX        =   1879
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Custodia:"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblSituación 
         Height          =   225
         Index           =   1
         Left            =   -69760
         TabIndex        =   157
         Top             =   3660
         Visible         =   0   'False
         Width           =   1065
         _Version        =   851968
         _ExtentX        =   1879
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Marca Interna:"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblSituación 
         Height          =   225
         Index           =   0
         Left            =   -69640
         TabIndex        =   155
         Top             =   3330
         Visible         =   0   'False
         Width           =   1245
         _Version        =   851968
         _ExtentX        =   2196
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Nro. Cheque:"
         Transparent     =   -1  'True
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Interno Hasta:"
         Height          =   195
         Index           =   7
         Left            =   -66400
         TabIndex        =   92
         Top             =   1200
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label lblFicha 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   225
         Index           =   0
         Left            =   -69820
         TabIndex        =   89
         Top             =   480
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblFicha 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor:"
         Height          =   225
         Index           =   1
         Left            =   -69820
         TabIndex        =   88
         Top             =   840
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblFicha 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Int Desde:"
         Height          =   225
         Index           =   2
         Left            =   -69820
         TabIndex        =   87
         Top             =   1200
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblFicha 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Firmante "
         Height          =   225
         Index           =   3
         Left            =   -69820
         TabIndex        =   86
         Top             =   1530
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblFicha 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Receptor :"
         Height          =   225
         Index           =   4
         Left            =   -69820
         TabIndex        =   85
         Top             =   1920
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblFicha 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Estado :"
         Height          =   225
         Index           =   5
         Left            =   -69820
         TabIndex        =   84
         Top             =   2280
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblFicha 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Banco:"
         Height          =   225
         Index           =   6
         Left            =   -69790
         TabIndex        =   83
         Top             =   2640
         Visible         =   0   'False
         Width           =   1065
      End
      Begin XtremeSuiteControls.Label lblSituación 
         Height          =   225
         Index           =   2
         Left            =   -69820
         TabIndex        =   68
         Top             =   2970
         Visible         =   0   'False
         Width           =   1245
         _Version        =   851968
         _ExtentX        =   2196
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Observaciones :"
         Transparent     =   -1  'True
      End
   End
   Begin MSAdodcLib.Adodc bcheques 
      Height          =   360
      Left            =   14280
      Top             =   0
      Visible         =   0   'False
      Width           =   2250
      _ExtentX        =   3969
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
      Left            =   11460
      Top             =   1320
      Visible         =   0   'False
      Width           =   2250
      _ExtentX        =   3969
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
   Begin MSAdodcLib.Adodc bccliente 
      Height          =   330
      Left            =   14220
      Top             =   1410
      Visible         =   0   'False
      Width           =   2250
      _ExtentX        =   3969
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
   Begin TabDlg.SSTab TabCheques 
      Height          =   7575
      Left            =   17610
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   13361
      _Version        =   393216
      TabOrientation  =   1
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Ing. Movimiento"
      TabPicture(0)   =   "frmCheques.frx":4ABD
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame8"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "vfecha"
      Tab(0).Control(3)=   "cmdNuevo"
      Tab(0).Control(4)=   "cmdGuardar"
      Tab(0).Control(5)=   "Label5"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Modo Consulta"
      TabPicture(1)   =   "frmCheques.frx":4AD9
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label18"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label19"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label8"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cvfhasta"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cvfdesde"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdBuscar"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "f1"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Frame7"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Listado de Cheques"
      TabPicture(2)   =   "frmCheques.frx":4AF5
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame Frame8 
         Height          =   555
         Left            =   -73320
         TabIndex        =   59
         Top             =   30
         Width           =   11775
         Begin VB.OptionButton opModo 
            Caption         =   "Recibir Cheque"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   2
            Left            =   30
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   150
            Width           =   6435
         End
         Begin VB.OptionButton opModo 
            Caption         =   "Emitir Cheque"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   1
            Left            =   6480
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   150
            Value           =   -1  'True
            Width           =   5205
         End
      End
      Begin VB.Frame Frame7 
         Height          =   765
         Left            =   240
         TabIndex        =   56
         Top             =   390
         Width           =   2085
         Begin VB.OptionButton opModo 
            Caption         =   "Emitir Cheque"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   90
            TabIndex        =   58
            Top             =   480
            Width           =   1605
         End
         Begin VB.OptionButton opModo 
            Caption         =   "Recibir Cheque"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   90
            TabIndex        =   57
            Top             =   150
            Width           =   1755
         End
      End
      Begin VB.CheckBox f1 
         Caption         =   "Anular Fechas"
         Height          =   225
         Left            =   11370
         TabIndex        =   47
         Top             =   750
         Value           =   1  'Checked
         Width           =   1365
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   615
         Left            =   210
         Picture         =   "frmCheques.frx":4B11
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Ejecutar consulta"
         Top             =   5280
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Height          =   4155
         Left            =   120
         TabIndex        =   23
         Top             =   2400
         Width           =   14085
         Begin VB.TextBox Text1 
            BackColor       =   &H80000016&
            Height          =   345
            Left            =   2760
            TabIndex        =   53
            Top             =   3540
            Width           =   11115
         End
         Begin VB.ComboBox cvestado 
            BackColor       =   &H80000016&
            Height          =   315
            ItemData        =   "frmCheques.frx":4C13
            Left            =   2730
            List            =   "frmCheques.frx":4C26
            TabIndex        =   51
            Text            =   "Cualquier estado"
            Top             =   3090
            Width           =   1965
         End
         Begin VB.CheckBox f2 
            Caption         =   "Anular Fechas"
            Height          =   225
            Left            =   9330
            TabIndex        =   48
            Top             =   1950
            Value           =   1  'Checked
            Width           =   1365
         End
         Begin VB.TextBox cvihasta 
            Alignment       =   2  'Center
            BackColor       =   &H80000016&
            Height          =   345
            Left            =   7260
            TabIndex        =   45
            Top             =   2460
            Width           =   1665
         End
         Begin VB.TextBox cvidesde 
            BackColor       =   &H80000016&
            Height          =   315
            Left            =   2760
            TabIndex        =   28
            Top             =   2580
            Width           =   1635
         End
         Begin VB.TextBox cvfirmante 
            BackColor       =   &H80000016&
            Height          =   345
            Left            =   1530
            TabIndex        =   27
            Top             =   1470
            Width           =   12315
         End
         Begin VB.TextBox cvncheque 
            BackColor       =   &H80000016&
            Height          =   345
            Left            =   1530
            TabIndex        =   26
            Top             =   1020
            Width           =   5625
         End
         Begin VB.ComboBox cvsucursal 
            BackColor       =   &H80000016&
            Height          =   315
            Left            =   8370
            TabIndex        =   25
            Top             =   630
            Width           =   5535
         End
         Begin VB.ComboBox cvbanco 
            BackColor       =   &H80000016&
            Height          =   315
            Left            =   1530
            TabIndex        =   24
            Top             =   630
            Width           =   5625
         End
         Begin VB.TextBox cvnombre 
            BackColor       =   &H80000016&
            Height          =   315
            Left            =   1530
            TabIndex        =   9
            Top             =   180
            Width           =   12345
         End
         Begin MSComCtl2.DTPicker cvddesde 
            Height          =   285
            Left            =   2820
            TabIndex        =   29
            Top             =   1980
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   25755649
            CurrentDate     =   38029
         End
         Begin MSComCtl2.DTPicker cvdhasta 
            Height          =   285
            Left            =   7290
            TabIndex        =   30
            Top             =   1950
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   25755649
            CurrentDate     =   38029
         End
         Begin VB.Label Label38 
            Caption         =   "> Nombre del receptor del cheque :"
            Height          =   285
            Left            =   150
            TabIndex        =   54
            Top             =   3630
            Width           =   3165
         End
         Begin VB.Label Label22 
            Caption         =   "> Estado del cheque  :"
            Height          =   225
            Left            =   180
            TabIndex        =   49
            Top             =   3180
            Width           =   1845
         End
         Begin VB.Label Label21 
            Caption         =   "> Hasta el Importe :"
            ForeColor       =   &H00808080&
            Height          =   225
            Left            =   4650
            TabIndex        =   46
            Top             =   2610
            Width           =   1755
         End
         Begin VB.Label Label17 
            Caption         =   "> Fecha de Depósito hasta :"
            ForeColor       =   &H00808080&
            Height          =   225
            Left            =   4620
            TabIndex        =   38
            Top             =   1980
            Width           =   2415
         End
         Begin VB.Label Label16 
            Caption         =   "> Desde el Importe :"
            ForeColor       =   &H00808080&
            Height          =   225
            Left            =   150
            TabIndex        =   37
            Top             =   2580
            Width           =   1905
         End
         Begin VB.Label Label15 
            Caption         =   "> Firmante :"
            Height          =   285
            Left            =   150
            TabIndex        =   36
            Top             =   1560
            Width           =   1275
         End
         Begin VB.Label Label14 
            Caption         =   "> Fecha de Depósito del cheque  :"
            ForeColor       =   &H00808080&
            Height          =   225
            Left            =   150
            TabIndex        =   35
            Top             =   2010
            Width           =   2625
         End
         Begin VB.Label Label13 
            Caption         =   "> Sucursal :"
            Height          =   285
            Left            =   7410
            TabIndex        =   34
            Top             =   690
            Width           =   1155
         End
         Begin VB.Label Label12 
            Caption         =   "> Nº del Cheque  :"
            Height          =   375
            Left            =   120
            TabIndex        =   33
            Top             =   1080
            Width           =   1395
         End
         Begin VB.Label Label11 
            Caption         =   "> Banco  :"
            Height          =   315
            Left            =   120
            TabIndex        =   32
            Top             =   690
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "> Nombre :"
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4335
         Left            =   -73320
         TabIndex        =   12
         Top             =   900
         Width           =   11775
         Begin VB.TextBox txtNroInterno 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   6930
            TabIndex        =   3
            Top             =   1650
            Width           =   1785
         End
         Begin VB.Frame Frame4 
            Height          =   45
            Left            =   -30
            TabIndex        =   55
            Top             =   1470
            Width           =   11805
         End
         Begin VB.TextBox txtNombreEndoso 
            Height          =   315
            Left            =   2790
            TabIndex        =   6
            Top             =   3870
            Width           =   8775
         End
         Begin VB.ComboBox txtEstado 
            Height          =   315
            ItemData        =   "frmCheques.frx":4C6C
            Left            =   2820
            List            =   "frmCheques.frx":4C7C
            TabIndex        =   8
            Text            =   "No Acreditado"
            Top             =   3450
            Width           =   1785
         End
         Begin VB.TextBox txtNombre 
            Height          =   315
            Left            =   1020
            TabIndex        =   0
            Top             =   210
            Width           =   4215
         End
         Begin VB.ComboBox cboBanco 
            Height          =   315
            Left            =   1020
            TabIndex        =   1
            Top             =   630
            Width           =   10515
         End
         Begin VB.ComboBox cboSucursal 
            Height          =   315
            Left            =   1020
            TabIndex        =   2
            Top             =   1050
            Width           =   10515
         End
         Begin VB.TextBox txtFirmante 
            Height          =   315
            Left            =   2790
            TabIndex        =   4
            Top             =   2070
            Width           =   5925
         End
         Begin VB.TextBox txtImporte 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   2820
            TabIndex        =   5
            Top             =   2970
            Width           =   1785
         End
         Begin MSComCtl2.DTPicker dtpDeposito 
            Height          =   285
            Left            =   2820
            TabIndex        =   13
            Top             =   2520
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   503
            _Version        =   393216
            Format          =   25755649
            CurrentDate     =   38029
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "> Nº Interno  :"
            Height          =   195
            Left            =   5040
            TabIndex        =   62
            Top             =   1695
            Width           =   1590
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            Caption         =   "> Nombre del receptor del cheque :"
            Height          =   195
            Left            =   25
            TabIndex        =   52
            Top             =   3930
            Width           =   2700
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Caption         =   "> Estado del cheque  :"
            Height          =   195
            Left            =   25
            TabIndex        =   50
            Top             =   3480
            Width           =   2700
         End
         Begin VB.Label Label3 
            Caption         =   "> Nombre :"
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "> Banco  :"
            Height          =   315
            Left            =   120
            TabIndex        =   19
            Top             =   690
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "> Nº del Cheque  :"
            Height          =   195
            Left            =   30
            TabIndex        =   18
            Top             =   1690
            Width           =   2700
         End
         Begin VB.Label Label2 
            Caption         =   "> Sucursal :"
            Height          =   285
            Left            =   120
            TabIndex        =   17
            Top             =   1110
            Width           =   1155
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "> Fecha de Depósito del cheque  :"
            Height          =   195
            Left            =   25
            TabIndex        =   16
            Top             =   2550
            Width           =   2700
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "> Firmante :"
            Height          =   195
            Left            =   25
            TabIndex        =   15
            Top             =   2130
            Width           =   2700
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "> Importe del cheque :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   240
            Left            =   25
            TabIndex        =   14
            Top             =   3020
            Width           =   2700
         End
      End
      Begin MSComCtl2.DTPicker vfecha 
         Height          =   315
         Left            =   -71400
         TabIndex        =   21
         Top             =   600
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   25755649
         CurrentDate     =   38029
      End
      Begin MSComCtl2.DTPicker cvfdesde 
         Height          =   285
         Left            =   9420
         TabIndex        =   39
         Top             =   570
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   25755649
         CurrentDate     =   38029
      End
      Begin MSComCtl2.DTPicker cvfhasta 
         Height          =   285
         Left            =   9420
         TabIndex        =   41
         Top             =   840
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   25755649
         CurrentDate     =   38029
      End
      Begin VB.CommandButton cmdNuevo 
         Appearance      =   0  'Flat
         Caption         =   "Nuevo"
         Height          =   495
         Left            =   -72540
         Picture         =   "frmCheques.frx":4CB0
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Nuevo movimiento"
         Top             =   5280
         UseMaskColor    =   -1  'True
         Width           =   825
      End
      Begin VB.CommandButton cmdGuardar 
         Appearance      =   0  'Flat
         Caption         =   "Grabar"
         Height          =   495
         Left            =   -73350
         Picture         =   "frmCheques.frx":4DB2
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Guardar movimiento"
         Top             =   5280
         UseMaskColor    =   -1  'True
         Width           =   825
      End
      Begin VB.Label Label5 
         Caption         =   "> Fecha del movimiento  :"
         Height          =   165
         Left            =   -73320
         TabIndex        =   22
         Top             =   660
         Width           =   1935
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Modo consulta de cheques"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   2640
         TabIndex        =   44
         Top             =   120
         Width           =   9495
      End
      Begin VB.Label Label19 
         Caption         =   "> Fecha de confección del cheque hasta :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5400
         TabIndex        =   42
         Top             =   810
         Width           =   3795
      End
      Begin VB.Label Label18 
         Caption         =   "> Fecha de confección del cheque desde :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5400
         TabIndex        =   40
         Top             =   540
         Width           =   3885
      End
   End
   Begin XtremeSuiteControls.PushButton PBFiltrar 
      Height          =   435
      Left            =   60
      TabIndex        =   133
      Top             =   7110
      Width           =   17295
      _Version        =   851968
      _ExtentX        =   30506
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "Filtrar datos para consultar"
      BackColor       =   -2147483644
      UseVisualStyle  =   -1  'True
      Picture         =   "frmCheques.frx":4EB4
      ImageAlignment  =   6
   End
End
Attribute VB_Name = "frmCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public dedonde As String
Public vToBack As String
Public vViene
Dim vcbanco, vcsucursal  As Integer
Public vRemitoCheques As Long
Public vModifica As Boolean
Dim vIdCheques As Long
Dim vdscheques As dsCheques
Dim vidbanco As Long

Dim vninternos As String

Dim vsql As String
Dim vfiltro, vfiltrobanco As String
Private Sub BuscarCliente()
On Error Resume Next

    Dim rsFiltro As New ADODB.Recordset, sqlFiltro As String

    
    If opModo(1).Value = True Then
        sqlFiltro = "SELECT * FROM proveedores WHERE (nombre LIKE '%" + Trim(txtNombre.Text) + "%') or (codigo LIKE '%" + Trim(txtNombre.Text) + "%')"
    Else
        sqlFiltro = "SELECT * FROM clientes WHERE (nombre like '%" + Trim(txtNombre.Text) + "%') or (codigo LIKE '%" + Trim(txtNombre.Text) + "%')"
    End If

    With rsFiltro
        .CursorLocation = adUseClient
        Call .Open(sqlFiltro, ConnDDBB, adOpenStatic, adLockPessimistic)

        If .EOF Then
    
            If opModo(1).Value = True Then
                'frmBuscarProveedor.Show
                'frmBuscarProveedor.txtProveedor = txtNombre.Text
                'frmBuscarProveedor.txtProveedor.SetFocus
                'frmBuscarProveedor.o = 5
            Else
                'frmBuscarCliente.Show
                'frmBuscarCliente.txtClientes.Text = txtNombre.Text
                'frmBuscarCliente.txtClientes.SetFocus
                'frmBuscarCliente.o = 5
            End If

        Else

            txtNombre.Tag = .Fields("Codigo").Value
            txtNombre.Text = .Fields("Nombre").Value
            'cboBanco.SetFocus
        End If
    
    End With
    
    sqlFiltro = ""
    
    If rsFiltro.State = 1 Then
        rsFiltro.Close
        Set rsFiltro = Nothing
    End If
    
If Err Then GrabarLog "BuscarCliente", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub BuscarRemito(vnroremito As Long)
On Error Resume Next

    With bcheques
        .RecordSource = "SELECT * FROM cheques WHERE (remito = " & vnroremito & ")"
        .Refresh
        
        If Not .Recordset.EOF = True Then .Recordset.MoveFirst
    
    End With

If Err Then GrabarLog "BuscarRemito", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cboCambiarEstado_Click()
On Error Resume Next

    cboCambiarEstado.Tag = Val(TraerDato("EstadoCheque", "Descripcion = '" & Trim(cboCambiarEstado.Text) & "'", "idEstadoCheque"))

If Err Then GrabarLog "cboCambiarEstado_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cboCambiarEstado_GotFocus()
On Error Resume Next

    Call CargarComboNew("EstadoCheque", "Descripcion", cboCambiarEstado, True)

If Err Then GrabarLog "cboCambiarEstado_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cmdCambiarEstado_Click(Index As Integer)

On Error Resume Next
    
    
    
'if me.vtipoOperacion = "Depósito en Banco"
    
    
    
    
    
    Select Case Index
    
        Case 0
            If Not Val(lblNroCheque.Tag) = 0 Then
                If MsgBox("Esta seguro que desea cambiar el estado del Cheque Nº " & Val(lblNroCheque.Caption) & "?", vbInformation + vbYesNo, "Mensaje ...") = vbYes Then
                    ' Alfredo: acá guardo
                    Call EjecutarScript("INSERT INTO HistoricoEstadosCheques (idCheques, idEstadoAnterior, idEstadoActual, FechaCambio, propietario) VALUES (" & Val(lblNroCheque.Tag) & "," & Val(KlexCheques.TextMatrix(KlexCheques.Row, 6)) & "," & Val(cboCambiarEstado.Tag) & ",'" & strfechaMySQL(dtpCambioFecha.Value) & ")'," + Str(Me.vpropietario) + ")")
                    
                    Call EjecutarScript("UPDATE Cheques SET idEstadoCheque = " & Val(cboCambiarEstado.Tag) & ",propietario='" + Me.vpropietario + "' WHERE (idCheques = " & Val(lblNroCheque.Tag) & ")")

                    If Err.Number = 0 Then MsgBox "Estado Cambiado Correctamente!!", vbInformation, "Mensaje ..."
                    
                End If
                        
            Else
                MsgBox "No existe el Cheque que ha seleccionado !!!", vbExclamation, "Mensaje ..."
            End If
        
        Case 1
            If MsgBox("Esta seguro que desea cambiar el estado de todos los cheques ?", vbInformation + vbYesNo, "Mensaje ...") = vbYes Then
                Call EjecutarScript("UPDATE Cheques SET idEstadoCheque = " & Val(cboCambiarEstado.Tag) & "")
            End If
        
        Case 2
            lblNroCheque.Caption = ""
            GbEstado.Visible = False
            KlexCheques.Enabled = True
            'PicInferior.Enabled = True
    
    End Select

Call PBFiltrar_Click

If Err Then GrabarLog "cmdCambiarEstado_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdAcciones_Click(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 0
            With KlexCheques
            
                frmChequesAlta.ModificarCheque (.TextMatrix(.Row, 1))
            End With
            
        
        Case 1
            With KlexCheques
                
                If MsgBox("Está seguro de borrar este cheque sin borrar los otros movimientos realcionados ?", vbYesNo, "Cuidado...") = vbNo Then Exit Sub
                
                Call BorrarBase("Cheques WHERE idCheques = " & Val(.TextMatrix(.Row, 1)) & "", pathDBMySQL)
                If .Rows = 2 Then
                    FormatoGrilla (1)
                Else
                    Call .RemoveItem(.Row)
                End If

            End With
            
        Case 2
                    
            vVieneImpresion = Me.Name
            'frmImprimir.Show
            reportesCheques
            ' acà va la impresión de cheques
             'frmImprimir 'acá estamos
        
        
        Case 3
            vVieneImpresion = Me.Name
            frmImprimir.chkActivarFecha(4).Value = xtpUnchecked
            frmImprimir.chkDiferido.Value = xtpChecked
            frmImprimir.Show
            
            'Call Imprimir(Index)
        
        Case 4
            'Cambiar Estado
            
            With KlexCheques
                If Not .Rows = 1 And Not Val(.TextMatrix(.Row, 1) = 0) Then
                    
                    cboCambiarEstado.Tag = EsNulo(.TextMatrix(.Row, 6))
                    cboCambiarEstado.Text = EsNulo(.TextMatrix(.Row, 7))
                    lblNroCheque.Tag = EsNulo(.TextMatrix(.Row, 1))
                    lblNroCheque.Caption = EsNulo(.TextMatrix(.Row, 2))
                    GbEstado.Visible = True
                    GbEstado.Top = 1320
                    GbEstado.Left = 3240
                    KlexCheques.Enabled = False
                    'PicInferior.Enabled = False
                End If
            End With
        Case 5
        
        Case 6
            Unload Me
            
        Case 7
        ' nuevo cheque
        If Me.vViene = "cobro" Then frmChequesAlta.vViene = "cobro"
        If Me.vViene = "pago" Then frmChequesAlta.vViene = "pago"
        
        frmChequesAlta.Show
        
        Case 8
        
        frmBancoCajaDetalle.vnrocheque = gbldsCheques.Ncheque
        Call frmBancoCajaDetalle.cmdFiltrar_Click
    
    End Select
    
If Err Then GrabarLog "cmdAcciones_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub reportesCheques()
Dim vshare As String

' vfiltrobanco se encarga de filtrar solamente los bancos

vshare = "SHAPE {SELECT B.idBancos, B.Descripcion, BC.idBancosCuentas, BC.Cuenta, BC.Descripcion, BC.CuentaContableAsociada, BC.idTipoCuentaBanco" & _
" FROM Bancos B INNER JOIN BancosCuentas BC ON B.idBancos = BC.idBancos where (1=1 " + vfiltrobanco + " ) ;}  AS ValoresDiferidosEncabezado" & _
" APPEND ({SELECT * FROM VistaCheques where 1=1 " & vfiltro & _
" }  AS ValoresDiferidosDetalle RELATE 'idBancos' TO 'idBancos','idBancosCuentas' TO 'idBancosCuentas') AS ValoresDiferidosDetalle"


     With Mantenimiento.rsValoresDiferidosEncabezado
                                If .State = 1 Then .Close
                                .Source = vshare
                                If .State = 0 Then .Open
                                .Close
                                .Open
    End With

drCheques.Show
'reportesGrupos ' verifica y realiza reportes de grupos
'reportesSimples ' verifica y realiza reportes de grupos
End Sub


Private Sub reportesGrupos()
' verifico que tipo de reporte degrupo es
If Left(vgrupo.Text, 2) = "0." Then Exit Sub ' no es un reporte de grupos
If Left(vgrupo.Text, 2) = "1." Then reporteGruposBanco
' Alfredo: continuar con los otros reportes
End Sub

Private Sub reporteGruposBanco()
Dim sqlGrupo As String
'sqlGrupo = fsqlGruposBanco() ' armo el filtro para los campos del grupos

Dim sqlDetalle As String
'sqlDetalle = fsqlDetalleBanco() ' armo los filtro para los campos del detalle
End Sub

Private Sub fsqlDetalleBanco()
'vSQLDetalle = " WHERE (Fecha >= '" & strfechaMySQL(dtpBancoCajaMovimiento(0).Value) & "' AND Fecha <= '" & strfechaMySQL(dtpBancoCajaMovimiento(1).Value) & "')"

End Sub

Private Sub cmdGuardar_Click()
    On Error Resume Next
    
    If Val(txtImporte.Text) = 0 Then
        MsgBox "Debe ingresar un IMPORTE para el Cheque!", vbExclamation, "Mensdaje ..."
        Exit Sub
    End If

    Dim rsGuardar As New ADODB.Recordset, sqlCheques As String
    
    With rsGuardar
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        
        If Not vModifica = True Then
            sqlCheques = "SELECT * FROM Cheques WHERE 1=2"
        Else
            sqlCheques = "SELECT * FROM Cheques WHERE (idCheques = " & vIdCheques & ")"
        End If
        
        Call .Open(sqlCheques, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If .EOF Then .AddNew
   
        .Fields("Fecha").Value = strfechaMySQL(vfecha.Value)
        .Fields("Nombre").Value = EsNulo(txtNombre.Text)
        .Fields("Codigo").Value = EsNulo(txtNombre.Tag)
        .Fields("Banco").Value = EsNulo(cboBanco.Text)
        .Fields("Sucursal").Value = EsNulo(cboSucursal.Text)
        .Fields("Monto").Value = Val(txtImporte.Text)
        .Fields("Endoso").Value = EsNulo(txtNombreEndoso.Text)
        .Fields("Remito").Value = EsNulo(vRemitoCheques)
    
        If opModo(2).Value = True Then
            .Fields("cp").Value = "c"
        Else
            .Fields("cp").Value = "p"
        End If
    
        .Fields("Firmante").Value = EsNulo(txtFirmante.Text)
        .Fields("FechaDeposito").Value = strfechaMySQL(dtpDeposito.Value)
        '.Fields("NCheque").Value = EsNulo(txtNroCheque.Text)
        .Fields("estado").Value = EsNulo(txtEstado.Text)
        .Fields("NroInterno").Value = Val(txtNroInterno)


        .Update

        Select Case dedonde
        
            Case "pctacte"
                frmCtaCteP.o1.Value = True
                frmCtaCteP.txtImporte.Text = txtImporte.Text
                'frmCtaCteP.txtComentario.Text = " Acreditación cheque Nº " & Trim(txtNroCheque.Text)
                frmCtaCteP.txtComentario.Tag = .Fields("idCheques").Value
        End Select

        If opModo(2).Value Then
            wccorrientes
        Else
            wpccorrientes
        End If

        vModifica = False

        LimpiarCampos

         'If vConfigGral.vIncluyeContabilidad = True Then
         '   With frmAsientosAlta
         '       .Show
         '       .ZOrder (0)
         '       .txtCuentaVieneDe.Text = Me.Caption
         '   End With
        'End If
        
    End With

    sqlCheques = ""
    
    If rsGuardar.State = 1 Then
        rsGuardar.Close
        Set rsGuardar = Nothing
    End If
    
    If Err Then
        GrabarLog "cmdGuardar_Click", Err.Number & " " & Err.Description, Me.Name
        MsgBox "Error!. Revisar operaciones.", vbCritical
    End If

End Sub
Private Sub Imprimir(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 2
            Unload Mantenimiento
            Load Mantenimiento
            
            With Mantenimiento.rsCheques
                If Not .State = 0 Then .Close
                
                .Source = bcheques.RecordSource
                
                If Not .State = 1 Then .Open
                .Close
                .Open
            
            End With
        
            With drCheques
                .Show
            End With
                
        Case 3
            Unload Mantenimiento
            Load Mantenimiento
            
            With Mantenimiento.rsCheques
                If Not .State = 0 Then .Close
                
                .Source = bcheques.RecordSource & " WHERE (FechaDeposito > '" & strfechaMySQL(Date) & "')"
                
                If Not .State = 1 Then .Open
                .Close
                .Open
            
            End With
        
            With drCheques
                .Sections("TituloEmpresa").Controls("Label1").Caption = "Listado de Cheques Diferidos"
                .Show
            End With
            'With drChequesPorDia
            '    .Show
            'End With
                
        Case 4
            'ImpresionDeUnChequeEnElPapel
            
        
    End Select
    
If Err Then GrabarLog "cmdImprimir_Click", Err.Number & "  " & Err.Description, Me.Name
End Sub
Private Sub Modificar()
    On Error Resume Next
    
    TabCheques.tab = 0
            
    With bcheques
    
        vfecha.Value = .Recordset("fecha").Value
        txtNombre.Tag = EsNulo(.Recordset("codigo").Value)
        txtNombre.Text = EsNulo(.Recordset("nombre").Value)
        cboBanco.Text = EsNulo(.Recordset("banco").Value)
        cboSucursal.Text = EsNulo(.Recordset("sucursal").Value)
        
       ' txtNroCheque.Text = EsNulo(.Recordset("ncheque").Value)
        txtNroInterno.Text = EsNulo(.Recordset("NroInterno").Value)
        
        txtFirmante.Text = EsNulo(.Recordset("firmante").Value)
        dtpDeposito.Value = EsNulo(.Recordset("FechaDeposito").Value)
        txtImporte.Text = EsNulo(.Recordset("Monto").Value)
        txtEstado.Text = EsNulo(.Recordset("estado").Value)
        txtNombreEndoso.Text = EsNulo(.Recordset("endoso").Value)
        
        'vvncheque = EsNulo(.Recordset("ncheque").Value)


        If .Recordset("cp").Value = "c" Then
            opModo(2).Value = True
        Else
            opModo(1).Value = True
        End If

        vIdCheques = .Recordset("idCheques").Value
        vModifica = True

    End With
    
    txtNombre.SetFocus

    If Err Then GrabarLog "Modificar", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdNuevo_Click()
    On Error Resume Next
    
    LimpiarCampos
    
    If Err Then GrabarLog "cmdNuevo_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdSalir_Click()
On Error Resume Next

    Select Case vToBack

        Case "remito"
            Unload Me
           frmRemito.RecargarForm
    
        Case "compras"
            Unload Me
            frmCompras.RecargarForm
                
        Case "pctacte"
            Unload Me
            frmCtaCteP.txtComentario.SetFocus
            
        Case Else
            Unload Me

    End Select

If Err Then GrabarLog "cmdSalir_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub cmdBuscar_Click()
On Error Resume Next
    
    Dim vsql As String
    
    vsql = ""

    If Not cvnombre = "" Then
        vsql = vsql & " AND (nombre like '%" + Trim(cvnombre.Text) + "%')"
    End If

    If Not cvfirmante = "" Then
        vsql = vsql & " AND (firmante = '" & Trim(cvfirmante.Text) & "')"
    End If

    If Not cvncheque = "" Then
        vsql = vsql & " AND (ncheque = '" & Trim(cvncheque.Text) & "')"
    End If

    If Not cvbanco = "" Then
        vsql = vsql + " AND (banco LIKE '%" + Trim(cvbanco.Text) & "%')"
    End If

    If Not cvsucursal = "" Then
        vsql = vsql & " AND (sucursal = '" & Trim(cvsucursal.Text) & "')"
    End If

    If Not (cvidesde = "" Or cvihasta = "") Then
        vsql = vsql & " AND (monto >= " & Val(cvidesde.Text) & " AND monto <= " & Val(cvihasta.Text) & ")"
    End If

    If f1.Value = 0 Then
        vsql = vsql & " AND (fecha >= '" & strfechaMySQL(cvfdesde.Value) & "' AND fecha <= '" & strfechaMySQL(cvfhasta.Value) + "')"
    End If

    If f2.Value = 0 Then
        vsql = vsql & " AND (FechaDeposito >= '" & strfechaMySQL(cvddesde.Value) + "' AND FechaDeposito <= '" & strfechaMySQL(cvdhasta.Value) + "')"
    End If

    If Not Trim(cvestado.Text) = "Cualquier estado" Then
        vsql = vsql & " AND (estado = '" & Trim(cvestado.Text) & "')"
    End If

    Dim i As Integer
    For i = 3 To 4
        If opModo(i).Value = True Then
            If i = 3 Then
                vsql = vsql & " AND (cp = 'c')"
            Else
                vsql = vsql & " AND (cp = 'p')"
            End If
        End If
    Next
    
    With bcheques
        .RecordSource = "SELECT * FROM cheques WHERE 1=1 " & vsql & ""
        .Refresh
        
        If Not .Recordset.EOF = True Then .Recordset.MoveFirst
    
    End With
    
    TabCheques.tab = 2

If Err Then GrabarLog "cmdBuscar_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub custodia_Click(Index As Integer)
Call fbuscarGrilla("Bancos", "Descripcion", "idBancos", Me.vdcustodia.Name, Me)  ' ema:
End Sub

Private Sub cvbanco_GotFocus()
On Error Resume Next

    Call CargarCombo("BANCO", "Nombre", cvbanco, False)

If Err Then GrabarLog "cvbanco_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub

Public Sub cvnombre_KeyPress(KeyAscii As Integer)
On Error Resume Next
    
    'If KeyAscii = 13 Then cbuscacli
    
If Err Then GrabarLog "cvnombre_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cvsucursal_GotFocus()
On Error Resume Next

    CargarCombo "COD_POS", "Localidad", cvsucursal, False

If Err Then GrabarLog "cvsucursal_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub f1_Click()
On Error Resume Next

    cvfdesde.Enabled = CBool(f1.Value - 1)
    cvfhasta.Enabled = CBool(f1.Value - 1)

If Err Then GrabarLog "f1_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub f2_Click()
On Error Resume Next
        
    cvddesde.Enabled = CBool(f2.Value - 1)
    cvdhasta.Enabled = CBool(f2.Value - 1)

If Err Then GrabarLog "f2_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
    
        SendKeys "{TAB}"
    
    End If
        

If Err Then GrabarLog "", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Load()
On Error Resume Next

    init


'    With Me
'          .Show
'         .Top = 0
'         .Left = 1000
'        .Width = 17460
'        .Height = 7590
'    End With
    
    txtFecha(0).Value = Date
    txtFecha(1).Value = Date
    
    vfecha.Value = Date
    dtpDeposito.Value = Date
    cvfdesde = Date
    cvfhasta = Date

    llenarComListado
    llenarComSituacion
    llenarCombOrdenamiento
    
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2 - 1000

    
    FormatoGrilla (1)
    'sb.Panels(2).Text = "Sistema preparado para la emisión de un cheque"

    'VerificarAcreditaciones

If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub init()

Me.TabBusqueda.SelectedItem = 0


'-------------------- vorden ------------------------------
Me.vorden.AddItem "Fecha"
Me.vorden.AddItem "IdCustodia"
Me.vorden.AddItem "MarcaInterna"
Me.vorden.AddItem "FechaAcreditacion"
Me.vorden.Text = "Fecha"
' ---------------------------------------------------------


' --------------- vgrupo --------------------------------
vgrupo.Clear
vgrupo.AddItem ("No efectuar agrupaciones")
vgrupo.AddItem ("Agrupar por Bancos y Sucursales")
vgrupo.AddItem ("Agrupar por Persona")
vgrupo.AddItem ("Estados")
vgrupo.AddItem ("Fechas de depósito")
vgrupo.AddItem ("Fechas de recepción del cheque")
vgrupo.Text = "No efectuar agrupaciones"
'----------------------------------------------------------
End Sub
Private Sub llenarComListado()
Me.ComListado.AddItem "(Todos)"
Me.ComListado.AddItem "En Cartera"
Me.ComListado.AddItem "Propios"
Me.ComListado.AddItem "Entregados"
Me.ComListado.AddItem "Depositados"
Me.ComListado.AddItem "Rechazados"
Me.ComListado.AddItem "Diferidos"
End Sub
Private Sub llenarComSituacion()
Me.CombOrdenamiento.AddItem "Cliente - Fecha de Vencimiento"
Me.CombOrdenamiento.AddItem "Fecha de Vencimiento - Cliente"
Me.CombOrdenamiento.AddItem "Fecha de Vencimiento - Importe"
Me.CombOrdenamiento.AddItem "Fecha de Vencimiento - Clering - Banco"
Me.CombOrdenamiento.AddItem "Fecha de Efectivación - Banco"
Me.CombOrdenamiento.AddItem "Tipo de valor  - Fecha de vencimiento"
End Sub
Private Sub llenarCombOrdenamiento()

End Sub

Public Sub filtro()
On Error Resume Next


    vfiltro = ""
    vfiltrobanco = ""
    
     If Not Me.txtFicha(9) = "" Then
     
        vfiltro = vfiltro + " and idbancos='" + Trim(Me.txtFicha(9)) + "'"
        vfiltrobanco = " and B.idbancos='" + Trim(Me.txtFicha(9)) + "'"
    End If
    
     If Not Me.vccustodia = "" Then vfiltro = vfiltro + " and idCustodia='" + Me.vccustodia + "'"
    
     If Not Me.vmarcaInterna = "" Then
     
     
        If chknroExacto.Value = xtpChecked Then
            vfiltro = vfiltro + " and marcainterna =" + Me.vmarcaInterna
        Else
            vfiltro = vfiltro + " and marcainterna >=" + Me.vmarcaInterna
        End If
        
        Me.CombOrdenamiento.Text = "marcainterna"
        vorden.Text = "marcainterna"
     End If
     
     If Not Me.vbusca = "" Then
     
        If chknroExacto.Value = xtpChecked Then
                vfiltro = vfiltro + " and marcainterna =" + Me.vbusca
        Else
                vfiltro = vfiltro + " and marcainterna=" + Me.vbusca.Text
                vfiltro = vfiltro + " or ncheque like '%" + Trim(Me.vbusca.Text) + "%'"
        End If
    
    End If
     
     If Not Me.vnrocheque = "" Then vfiltro = vfiltro + " and ncheque='" + Me.vnrocheque + "'"
      
     If Not Me.txtFicha(0) = "" Then vfiltro = vfiltro + " and cheques.codigo='" + Trim(Me.txtFicha(0)) + "'"
    
    ' If Not Me.txtFicha(2) = "" Then vfiltro = vfiltro + " and ClienteProveedor='" + Trim(Me.txtFicha(2)) + "'"
     
     If Not Me.txtFicha(2) = "" Then vfiltro = vfiltro + " and cheques.codigo='" + Trim(Me.txtFicha(2)) + "'"
    
     If Not Me.txtFicha(4) = "" Then vfiltro = vfiltro + " and cheques.NroInterno>=" + Str(Me.txtFicha(4)) + " and cheques.NroInterno <=" + Str(Me.txtFicha(4))
    
     If Not vFirmante.Text = "" Then vfiltro = vfiltro + " and cheques.Firmante='" + vFirmante.Text + "'"
     
     If Not Me.txtFicha(6) = "" Then vfiltro = vfiltro + " and cheques.endoso='" + Str(Me.txtFicha(6)) + "'"
     
     If Not Me.txtFicha(7) = "" Then vfiltro = vfiltro + " and cheques.idEstadoCheque=" + Str(Me.txtFicha(7))
     
     If Not Me.txtBuscar(2).Text = "" Then vfiltro = vfiltro + " and Observaciones like '%" + Trim(Me.txtBuscar(2).Text) + "%'"
     
     If Me.chkActivarFecha Then vfiltro = vfiltro + " and cheques.fecha >='" + strfechaMySQL(Me.txtFecha(0)) + "' and cheques.fecha <='" + strfechaMySQL(Me.txtFecha(1)) + "'"
     If Me.chkActivarFDeposito Then vfiltro = vfiltro + " and cheques.FechaDeposito >='" + strfechaMySQL(Me.txtFecha(2)) + "' and cheques.FechaDeposito <='" + strfechaMySQL(Me.txtFecha(3)) + "'"
     
     If Not (Me.vImporteDesde + Me.vImporteHasta) = "" Then vfiltro = vfiltro + " and monto >=" + Str(Me.vImporteDesde.Text) + " and  monto <= " + Str(Me.vImporteHasta.Text)
     
     
  ' controlo el lugar donde està el cheque. Por ejemplo en Cartera
   'If Trim(Me.ComListado.Text) = "En Cartera" Then vFiltro = vFiltro + " and (t.EsCaja='S') and bancosmovimientos.debito>0 and not (cheques.idCustodia='098')"
   If Trim(Me.ComListado.Text) = "En Cartera" Then vfiltro = vfiltro + " and (t.EsCaja='S') and not (cheques.idCustodia='098') or (cheques.idCustodia='') or (cheques.idCustodia is  null) "
  
   
   If Trim(Me.ComListado.Text) = "Depositados" Then vfiltro = vfiltro + " and (t.EsCaja='N') and bancosmovimientos.debito>0"
   If Trim(Me.ComListado.Text) = "Entregados" Then vfiltro = vfiltro + " and (cheques.idCustodia='098')"

    If Trim(Me.ComListado.Text) = "Diferidos" Then vfiltro = vfiltro + " and (cheques.FechaDeposito > '" + strfechaMySQL(Date) + "' and propietario= 'Propio' )"
   
   
    If Me.rbpropios.Value Then
            vfiltro = vfiltro + " and (propietario = 'propio') "
    End If
    
     
    If Me.rdterceros.Value Then
            vfiltro = vfiltro ' + " and (not (propietario = 'propio') or  propietario is null) "
    End If
     
    If Me.chkChkSinCustodia.Value Then
            vfiltro = vfiltro + " and idcustodia is null or idcustodia = 0 "
    End If
     
     
     
   If Trim(vfiltro) = "" Then
        If MsgBox("No se han seleccionado datos para filtrar. Está seguro que quiere ver todos los datos ?", vbYesNo, "Filtro...") = vbNo Then Exit Sub
   End If
   
    
    
    
   Dim vvistacheques As String
   
   'vvistacheques = " bancos " & _
   '                " INNER JOIN cheques ON (bancos.idBancos=cheques.idBancos) " & _
   '                " INNER JOIN estadocheque ON (cheques.idEstadoCheque=estadocheque.idEstadoCheque)"
   
   
   
   If Me.RadResumido Then
    'vsql = ChequesFiltros(vFiltro + vninternos, "resumen")
    vsql = ChequesFiltros(vfiltro, Me.vorden.Text, "resumen")

   End If
   
   If Me.RadDetallado Then
    'vsql = ChequesFiltros(vFiltro + vninternos, "historial")
    vsql = ChequesFiltros(vfiltro, vorden.Text, "historial")
   End If
   
    
  
     
        With bcheques
        .ConnectionString = pathDBMySQL
        If dedonde = "pctacte" Then
            .RecordSource = "SELECT * FROM VistaCheques WHERE (idEstadoCheque = 2)"
            .Refresh
        Else
            .RecordSource = vsql
            .Refresh
        End If
            
            
       If Not .Recordset.RecordCount > 0 Then
        Me.Caption = "No hay cheques para mostar"
        Me.KlexCheques.Visible = False
        Exit Sub
       End If
       
       ' CargarCheques ' Alfredo: acá se va a llenar la grilla con los datos de la consulta
       
       
       
     With Mantenimiento.rscheque2
                                If .State = 1 Then .Close
                                .Source = vsql
                                If .State = 0 Then .Open
                                .Close
                                .Open
    End With

    
    
       
       Call LlenarGrilla("cheques", Me.KlexCheques, vsql, "")
    
    End With

If Err Then
    GrabarLog "Filtro", Err.Number & " " & Err.Description, Me.Caption
    Exit Sub
End If
End Sub


Private Sub FormatoGrilla(vCantidadRenglones As Integer)
On Error Resume Next

    Dim i As Integer

    With KlexCheques
        .FixedRows = 1
        .FixedCols = 1
    
        .Cols = 20
        .Rows = vCantidadRenglones + 1
        
        If vCantidadRenglones = 1 Then
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
                .ColWidth(i) = 0
            Next
        End If
        
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 400
        
        .TextMatrix(0, 1) = "idCheques"
        .ColWidth(1) = 0
               
        .TextMatrix(0, 2) = "Nº Cheque"
        .ColWidth(2) = 1000
        
        .TextMatrix(0, 3) = "Nº Int."
        .ColWidth(3) = 1000
        
        .TextMatrix(0, 4) = "Fecha"
        .ColWidth(4) = 1250
        
        .TextMatrix(0, 5) = "FechaDeposito"
        .ColWidth(5) = 0
        
        .TextMatrix(0, 6) = "idEstadoCheque"
        .ColWidth(6) = 0
        
        .TextMatrix(0, 7) = "Estado"
        .ColWidth(7) = 1500
        .ColAlignment(6) = vbAlignRight
        
        .TextMatrix(0, 8) = "Codigo"
        .ColWidth(8) = 1500
        
        .TextMatrix(0, 9) = "Nombre"
        .ColWidth(9) = 1500
        
        .TextMatrix(0, 10) = "CP"
        .ColWidth(10) = 750
        
        .TextMatrix(0, 11) = "Banco"
        .ColWidth(11) = 1000
        
        .TextMatrix(0, 12) = "Sucursal"
        .ColWidth(12) = 1000
        
        .TextMatrix(0, 13) = "Endoso"
        .ColWidth(13) = 0
        
        .TextMatrix(0, 14) = "Firmante"
        .ColWidth(14) = 0
        
        .TextMatrix(0, 15) = "Fecha Acreditacion"
        .ColWidth(15) = 1250
        
        .TextMatrix(0, 16) = "Remito"
        .ColWidth(16) = 0
        
        .TextMatrix(0, 17) = "Importe"
        .ColWidth(17) = 1500
        .ColDisplayFormat(17) = "#0.000"
        .ColAlignment(17) = vbAlignRight

        .TextMatrix(0, 18) = "Observaciones"
        .ColWidth(18) = 2500
        
        .TextMatrix(0, 19) = "Propietarios"
        .ColWidth(18) = 2500
        


        .BackColorAlternate = 14737632
    End With
    
If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub CargarCheques()
    On Error Resume Next
    
    Dim i As Integer
    
    With bcheques
        If Not .Recordset.EOF = True Then .Recordset.MoveFirst
            
        FormatoGrilla (.Recordset.RecordCount)  '(Val(GenerarDato("SELECT COUNT(idCheques) as CantidadDeRegistros FROM VistaCheques", "CantidadDeRegistros")))
        
        i = 1
        Do Until .Recordset.EOF = True
        
            KlexCheques.TextMatrix(i, 0) = ""
            KlexCheques.TextMatrix(i, 1) = EsNulo(.Recordset("idCheques").Value)
            KlexCheques.TextMatrix(i, 2) = EsNulo(.Recordset("NCheque").Value)
            KlexCheques.TextMatrix(i, 3) = EsNulo(.Recordset("NroInterno").Value)
            KlexCheques.TextMatrix(i, 4) = EsNulo(.Recordset("Fecha").Value)
            KlexCheques.TextMatrix(i, 5) = EsNulo(.Recordset("FechaDeposito").Value)
            KlexCheques.TextMatrix(i, 6) = EsNulo(.Recordset("idEstadoCheque").Value)
            KlexCheques.TextMatrix(i, 7) = EsNulo(.Recordset("Descripcion").Value)
            KlexCheques.TextMatrix(i, 8) = EsNulo(.Recordset("Codigo").Value)
            KlexCheques.TextMatrix(i, 9) = EsNulo(.Recordset("Nombre").Value)
            KlexCheques.TextMatrix(i, 10) = EsNulo(.Recordset("CP").Value)
            KlexCheques.TextMatrix(i, 11) = EsNulo(.Recordset("Banco").Value)
            KlexCheques.TextMatrix(i, 12) = EsNulo(.Recordset("Sucursal").Value)
            KlexCheques.TextMatrix(i, 13) = EsNulo(.Recordset("Endoso").Value)
            KlexCheques.TextMatrix(i, 14) = EsNulo(.Recordset("Firmante").Value)
            KlexCheques.TextMatrix(i, 15) = EsNulo(.Recordset("FechaAcreditacion").Value)
            KlexCheques.TextMatrix(i, 16) = EsNulo(.Recordset("Remito").Value)
            KlexCheques.TextMatrix(i, 17) = EsNulo(.Recordset("Monto").Value)
            KlexCheques.TextMatrix(i, 18) = EsNulo(.Recordset("Observaciones").Value)
            KlexCheques.TextMatrix(i, 19) = EsNulo(.Recordset("Propietario").Value)
            
            .Recordset.MoveNext
        
            i = i + 1
        Loop
        
        .Refresh
        
        If .Recordset.EOF = True Then .Recordset.MoveLast
    End With
    
If Err Then GrabarLog "CalcularSaldo", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub LimpiarCampos()
    On Error Resume Next
    
    vIdCheques = 0
    vRemitoCheques = 0
    dtpDeposito.Value = Date
    vfecha.Value = Date
    txtNombre.Text = ""
    txtNombre.Tag = ""
    cboBanco.Text = ""
    cboSucursal.Text = ""
   ' txtNroCheque.Text = ""
    txtNroInterno.Text = ""
    txtFirmante.Text = ""
    txtImporte.Text = ""
    txtNombreEndoso.Text = ""
    
    txtNombre.SetFocus
    
    If Err Then GrabarLog "LimpiarCampos", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub DgCheques_DblClick()
On Error Resume Next

    With bcheques
        If dedonde = "pctacte" Then
            .Recordset("estado").Value = "Endosado"
            .Recordset("endoso").Value = frmCtaCteP.txtProveedor.Text
            
            .Recordset.Update
            
            frmCtaCteP.txtImporte.Text = Val(frmCtaCteP.txtImporte) + Val(.Recordset("monto").Value)
            frmCtaCteP.txtComentario.Text = frmCtaCteP.txtComentario.Text & " Pago con cheque nº " & .Recordset("ncheque").Value
            Unload Me
            frmCtaCteP.cmdGuardar.Enabled = False
        End If
    End With
    
If Err Then GrabarLog "DgCheques_DblClick", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub KlexCheques_Click()
cargarDsCheques
End Sub

Private Sub KlexCheques_DblClick()
Dim vsql As String

' selecciona un cheque
cargarDsCheques


With gbldsCheques

Select Case Me.vViene

    Case "cobro", "pagos"
    
    
       frmCobros.txtBancoCheque(0) = .idBancos
       
       vsql = "select * from bancos where idbancos ='" + Trim(.idBancos) + "'"
       frmCobros.txtBancoCheque(1).Text = traerDatos2(vsql, "Descripcion", pathDBMySQL)
       
       frmCobros.vfechaCheque = .fecha
       ' .FechaAcreditacion
       frmCobros.dtpDepositoCheque = .FechaDeposito
       frmCobros.txtImporteCheque = .monto
       frmCobros.txtNroCheque = .Ncheque
       ' .NroInterno
       ' .TipoMovimiento
       frmCobros.txtFirmanteCheque = .Firmante
       frmCobros.vidcheque = .vid
       frmCobros.txtBancoCheque(4) = .idCustodia
       
       
       frmCobros.vDCajaDestino = "Entregado a proveedor"
       frmCobros.vDCajaDestino.Tag = "098"
       frmCobros.vCCajaDestino = "098"
       
       frmCobros.vmarcaInterna = .marcainterna
       
       frmCobros.vsucursal = .sucursal
       
       
       
       Call frmCobros.txtBancoCheque_KeyPress(4, 13) ' hago cargar la caja
      ' Call frmCobros.txtBancoCheque_KeyPress(0, 13) ' hago cargar el banco
      ' Call frmCobros.txtBancoCheque_KeyPress(2, 13) ' hago cargar el sucursal
       
       
       frmCobros.vchequeComentario.SetFocus
       frmCobros.WindowState = vmaximizar
       
       Unload Me

    Case "frmIngresosEgresos"
                
        frmIngresosEgresos.txtAlta(3) = "CH"
        frmIngresosEgresos.txtAlta6 = .idCustodia
        'Call frmIngresosEgresos.txtAlta_KeyPress(6, 13)
       ' frmIngresosEgresos.txtAlta(8) =  ' falta en el ds el nro de cuenta
        ' .txtAlta(10) = ' cta contable
        frmIngresosEgresos.txtAlta(5) = .Ncheque
        frmIngresosEgresos.txtAlta(12) = .monto
        frmIngresosEgresos.txtAlta(13) = .Observaciones
        'frmIngresosEgresos.txtAlta(0) =
      '  frmIngresosEgresos.dtpFecha = .fecha
        frmIngresosEgresos.dtpValor = .FechaDeposito
        frmIngresosEgresos.vIdCheques = .vid
                
                
        If frmIngresosEgresos.RBDebeHaber(1).Value Then
                
             frmIngresosEgresos.VNuevaCustodiaNombre.Text = "Entrega a proveedor"
             frmIngresosEgresos.vNuevaCustodiaCodigo.Text = "098"
            
        End If
       
        If frmIngresosEgresos.RBDebeHaber(0).Value Then
             frmIngresosEgresos.VNuevaCustodiaNombre.Tag = .idCustodia
             frmIngresosEgresos.vNuevaCustodiaCodigo.Text = .idCustodia
             
             
             vsql = "select Descripcion as c from bancos where idbancos='" + .idCustodia + "'"
             frmIngresosEgresos.VNuevaCustodiaNombre.Text = traerDatos2(vsql, "c", pathDBMySQL)
            
        End If
       
       
       
       '-----------------------panic ! -------------------------------------------------
        frmIngresosEgresos.VNuevaCustodiaNombre.Tag = .idCustodia
        frmIngresosEgresos.vNuevaCustodiaCodigo.Text = .idCustodia
             
             
        vsql = "select Descripcion as c from bancos where idbancos='" + .idCustodia + "'"
        frmIngresosEgresos.VNuevaCustodiaNombre.Text = traerDatos2(vsql, "c", pathDBMySQL)
       
       '-----------------------------------------------------------------------
       
        frmIngresosEgresos.vDesBanco.Tag = .idBancos
        frmIngresosEgresos.vCodBanco.Text = .idBancos
        
        frmIngresosEgresos.vDesBanco.Text = traerDatos2("select * from bancos where idbancos=" + .idBancos, "descripcion", pathDBMySQL)
        
        frmIngresosEgresos.VchequesDisplay.Caption = fdsChequesToString
        
        Unload Me
        

   Case ""
   
   verDatosChequeSeleccionado
   
End Select

End With

Me.vViene = ""
'----------------------------
End Sub
Private Sub verDatosChequeSeleccionado()
On Error Resume Next
Dim vmensaje As String

With gbldsCheques
       vmensaje = vmensaje + Chr(13) + .idBancos
        vmensaje = vmensaje + Chr(13) + " > Fecha: " + Str(.fecha)
       ' .FechaAcreditacion
        vmensaje = vmensaje + Chr(13) + " > Diferida: " + Str(.FechaDeposito)
        vmensaje = vmensaje + Chr(13) + " > Importe: " + Str(.monto)
        vmensaje = vmensaje + Chr(13) + " > Nro. Cheque: " + Str(.Ncheque)
       ' .NroInterno
       ' .TipoMovimiento
        vmensaje = vmensaje + Chr(13) + " > Firmante: " + .Firmante
End With

MsgBox vmensaje

If Err Then Exit Sub
End Sub

Private Sub cargarDsCheques()
On Error Resume Next
Dim i As Integer
i = Me.KlexCheques.Row

'0  `idCheques`          int(10) AUTO_INCREMENT NOT NULL,
'1  `idEstadoCheque`     int(10) UNSIGNED NOT NULL,
'2  `Fecha`              date,
'3  `Codigo`             varchar(50),
'4  `Nombre`             varchar(250),
'5  `idBancos`           varchar(3),
'6  `idBancosCuentas`    int(10) UNSIGNED,
'7  `Ncheque`            varchar(20) NOT NULL,
'8  `Firmante`           varchar(50),
'9  `cp`                 varchar(1),
'10  `FechaDeposito`      date,
'11  `Monto`              double(15,3),
'12  `Endoso`             varchar(255),
'13  `Remito`             int(10) UNSIGNED,
'14  `NroInterno`         int(10) UNSIGNED,
'15  `Observaciones`      varchar(255),
'16  `FechaAcreditacion`  date,
'17  `Foto`               longblob,
'18  `TipoMovimiento`     varchar(2),
'19  `TimeStamp`          timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP,


With gbldsCheques
        .vid = Val(Me.KlexCheques.TextMatrix(i, 1))
        
        .Codigo = Me.KlexCheques.TextMatrix(i, 4)
        .Nombre = Me.KlexCheques.TextMatrix(i, 5)
        
        .idBancos = traerDatos2("select idBancos from cheques where idCheques=" + Str(.vid), "idBancos", pathDBMySQL)
        .idBancosCuentas = traerDatos2("select idBancosCuentas from cheques where idCheques=" + Str(.vid), "idBancosCuentas", pathDBMySQL)
        
        .fecha = traerDatos2("select fecha from cheques where idCheques=" + Str(.vid), "fecha", pathDBMySQL)
        
        '.FechaAcreditacion = Me.KlexCheques.TextMatrix(i, 16)
        .FechaDeposito = CDate(traerDatos2("select FechaDeposito from cheques where idCheques=" + Str(.vid), "FechaDeposito", pathDBMySQL))
        
        .monto = traerDatos2("select monto from cheques where idCheques=" + Str(.vid), "monto", pathDBMySQL)
        
        .Ncheque = traerDatos2("select Ncheque from cheques where idCheques=" + Str(.vid), "Ncheque", pathDBMySQL)
        .NroInterno = traerDatos2("select Nrointerno from cheques where idCheques=" + Str(.vid), "Nrointerno", pathDBMySQL)
        
        .idEstadoCheque = traerDatos2("select idEstadoCheque from cheques where idCheques=" + Str(.vid), "idEstadoCheque", pathDBMySQL)
        
        '.NroInterno = Me.KlexCheques.TextMatrix(i, 14)
        '.TipoMovimiento = Me.KlexCheques.TextMatrix(i, 19)
        '.Firmante = Me.KlexCheques.TextMatrix(i, 9)
        
        .idCustodia = traerDatos2("select idCustodia from cheques where idCheques=" + Str(.vid), "idCustodia", pathDBMySQL)

        .sucursal = traerDatos2("select sucursal from cheques where idCheques=" + Str(.vid), "sucursal", pathDBMySQL)

        .marcainterna = traerDatos2("select marcainterna from cheques where idCheques=" + Str(.vid), "marcainterna", pathDBMySQL)
        
        .Endoso = traerDatos2("select endoso from cheques where idCheques=" + Str(.vid), "endoso", pathDBMySQL)

        .Firmante = traerDatos2("select Firmante from cheques where idCheques=" + Str(.vid), "Firmante", pathDBMySQL)


End With
If Err Then Exit Sub
End Sub
Private Sub OpModo_Click(Index As Integer)
On Error Resume Next

    If Index = 1 Then
        'sb.Panels(2).Text = "Sistema preparado para la emisión de un cheque"
    Else
        'sb.Panels(2).Text = "Sistema preparado para recibir un cheque"
    End If

If Err Then GrabarLog "op_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub pbCarga_Click(Index As Integer)
On Error Resume Next

    vVuelveBusqueda = Me.Name
    vVieneBusqueda = pbCarga(Index).Tag

    Select Case Index

        Case 0 To 10
            frmBusqueda.Show
    End Select
    
    If Err Then GrabarLog "pbCarga_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Public Sub PBFiltrar_Click()
On Error Resume Next
    
    Validar
    filtro
    verDatos
    

If Err Then GrabarLog "PBFiltrar_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Validar()
Me.Caption = "Listado de cheques "
End Sub
Private Sub verDatos()
Me.TabBusqueda.SelectedItem = 2
Me.KlexCheques.TopRow = KlexCheques.Rows - 1
End Sub
Private Sub limpiarcampo()
Dim i As Integer

For i = 0 To Me.txtFicha.Count - 3
    Me.txtFicha(i).Text = ""
Next

Me.vImporteDesde = ""
Me.vImporteHasta = ""

Me.Caption = "Listado de cheques"

Me.KlexCheques.Visible = True

End Sub

Private Sub PushButton1_Click()
On Error Resume Next
drChequesGral.WindowState = 2
drChequesGral.Show
If Err Then Exit Sub
End Sub

Private Sub PushButton2_Click()
        frmTransaccionMantenimiento.vnrointerno = gbldsCheques.NroInterno
        frmTransaccionMantenimiento.Show
End Sub

Private Sub PushButton3_Click()
On Error Resume Next
    
  Call grillaToExcel2(Me.KlexCheques)

If Err Then Exit Sub

End Sub

Private Sub PusLimpiarTodos_Click()
limpiarcampo
End Sub

Private Sub PusSemanaPróxima_Click(Index As Integer)

End Sub

Private Sub PusVerHistorial_Click(Index As Integer)
On Error Resume Next
frmChequesAlta.Show
frmChequesAlta.TabAlta.TabIndex = 4


Call LlenarGrilla("historicoestadocheque", frmCheques.KlexCheques, "select * from historicoestadoscheques where idCheques=" + Str(Me.KlexCheques.TextMatrix(Me.KlexCheques.Row, 1)), "")
If Err Then Exit Sub
End Sub

Private Sub TabBusqueda_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
If Item.Caption = "Ver Datos" Then
   ' filtro
End If
End Sub

Public Sub TabCheques_Click(PreviousTab As Integer)
On Error Resume Next

    Select Case TabCheques.tab
    
        Case 0
            opModo(1).Value = True
        Case 1
            opModo(3).Value = True
            cvnombre.SetFocus
        Case 2
    
    End Select

If Err Then GrabarLog "TabCheques_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub VerificarAcreditaciones()
On Error Resume Next

    Call EjecutarScript("UPDATE Cheques SET idEstadoCheque = 1 WHERE (idEstadoCheque = 2) AND (FechaDeposito <= '" & strfechaMySQL(Date) + "')")
    
    bcheques.Refresh
    
    
If Err Then GrabarLog "VerificarAcreditaciones", Err.Number & " " & Err.Description, Me.Name
End Sub

Public Sub txtNombre_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 And Not Trim(txtNombre.Text) = "" Then
        
        BuscarCliente
    
    End If
    
If Err Then GrabarLog "txtNombre_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub wccorrientes() 'Asienta el mov. en las ctactes del Cliente
    On Error Resume Next

    Exit Sub
    With bccliente
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM CuentasCorrientes"
        .Refresh
    
        If vModifica = True Then
            'ESSTO ESTA MAAAAAAAAAAAAAAAAAAALLLLL
            '.Recordset.Find ("codigo = '" + Str(vCodigo) + "' and ncheque = '" + vvncheque + "'")
            If .Recordset.EOF Then .Recordset.AddNew
        Else
            .Recordset.AddNew
        End If
    
        .Recordset("fecha").Value = vfecha.Value
        .Recordset("codigo").Value = txtNombre.Tag
        .Recordset("nombre").Value = txtNombre.Text  ' nombre del cliente
       ' .Recordset("comentario").Value = "Acreditación cheque nro.  " + txtNroCheque.Text
       ' .Recordset("ncheque").Value = txtNroCheque.Text
    
        .Recordset("credito").Value = Val(txtImporte.Text)
        .Recordset("debito") = 0
    
        .Recordset.Update
    
    End With
    
    If Err Then GrabarLog "wccorrientes", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub wpccorrientes() ''Asienta el mov. en las ctactes del Proveedor
    On Error Resume Next
    
    Exit Sub
    
    With bpccliente
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM pcuentascorrientes"
        .Refresh
   
        If vModifica = True Then
            '.Recordset.Find ("ncheque = '" + Trim(vvncheque) + "'")
            'If .Recordset.EOF Then .Recordset.AddNew
        Else
            .Recordset.AddNew
        End If
    
        .Recordset("fecha").Value = strfechaMySQL(vfecha.Value)
        .Recordset("codigo").Value = EsNulo(txtNombre.Tag)
        .Recordset("nombre").Value = EsNulo(txtNombre.Text)
        .Recordset("comentario").Value = "Acreditación cheque nro.  " + txtNroCheque.Text
        .Recordset("ncheque").Value = Val(txtNroCheque.Text)
    
        .Recordset("credito").Value = Val(txtImporte.Text)
        .Recordset("debito").Value = 0
        
        .Recordset.Update

    End With
    
    If Err Then GrabarLog "wpccorrientes", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub vbusca_Change()
Dim vsql As String

Call PBFiltrar_Click

'vsql = " and marcainterna=" + Str(Val(vbusca.Text))


'vsql = ChequesFiltros(vsql, "marcainterna", "resumen")

'Call LlenarGrilla("cheques", Me.KlexCheques, vsql, "")
End Sub

Private Sub vbusca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.KlexCheques.Row = 1
    Call KlexCheques_DblClick
End If
End Sub

Private Sub vdcustodia_Change()
Me.vccustodia.Text = Me.vdcustodia.Tag
End Sub



Public Sub fInvariantes()
Dim vsql As String

' cheques sin custodias

vsql = "select count(*) as c  from cheques where idcustodia is null or idcustodia = 0"

If traerDatos2(vsql, "c", pathDBMySQL) > 0 Then

    If MsgBox("Hay cheques que se encuentran sin custodias." + Chr(13) + "Quiere verlos ?", vbYesNo) = vbYes Then
        
        frmCheques.chkChkSinCustodia.Value = xtpChecked
       Call frmCheques.PBFiltrar_Click
    
    End If
    
End If


End Sub


Function fchequeOrder()



End Function
