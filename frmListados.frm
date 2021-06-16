VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{63BEADB1-20E1-478A-9B40-DDDAFBF3624F}#1.0#0"; "bsGradientLabel.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.ShortcutBar.v13.0.0.Demo.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "Copia de KlexGrid.ocx"
Object = "{9746E3DA-06E1-4D26-9CE4-D9F6411A9C70}#1.0#0"; "SMGA_OcxTxt2008.ocx"
Begin VB.Form frmBancoCajaDetalle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta y Listado de Movimientos de Bancos y Cajas"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   15990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   15990
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   15
      Left            =   30
      TabIndex        =   15
      Top             =   0
      Width           =   15915
      _Version        =   851968
      _ExtentX        =   28072
      _ExtentY        =   -26
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton PusAyuda 
         Height          =   345
         Left            =   11430
         TabIndex        =   36
         Top             =   60
         Width           =   1485
         _Version        =   851968
         _ExtentX        =   2619
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Ayuda"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   195
         Left            =   60
         TabIndex        =   26
         Top             =   390
         Width           =   12975
         _Version        =   851968
         _ExtentX        =   22886
         _ExtentY        =   344
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
      End
      Begin XtremeSuiteControls.PushButton cmdCerrar 
         Height          =   345
         Left            =   8370
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Cerrar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmListados.frx":0000
      End
      Begin XtremeSuiteControls.PushButton PusRealizarArqueos 
         Height          =   345
         Left            =   60
         TabIndex        =   27
         Top             =   60
         Visible         =   0   'False
         Width           =   1725
         _Version        =   851968
         _ExtentX        =   3043
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Realizar Arqueos"
         UseVisualStyle  =   -1  'True
         RightToLeft     =   -1  'True
         Picture         =   "frmListados.frx":0400
      End
   End
   Begin XtremeSuiteControls.TabControl tabbc 
      Height          =   8565
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   15945
      _Version        =   851968
      _ExtentX        =   28125
      _ExtentY        =   15108
      _StockProps     =   68
      ItemCount       =   3
      SelectedItem    =   2
      Item(0).Caption =   "Filtrar"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "GBRangoDeFechas"
      Item(0).Control(1)=   "cmdFiltrar"
      Item(1).Caption =   "Datos"
      Item(1).ControlCount=   7
      Item(1).Control(0)=   "GroupBox1"
      Item(1).Control(1)=   "cmdImprimir(0)"
      Item(1).Control(2)=   "cmdImprimir(1)"
      Item(1).Control(3)=   "PushButton1"
      Item(1).Control(4)=   "PushButton4"
      Item(1).Control(5)=   "PusReimprimirComprobante"
      Item(1).Control(6)=   "PushButton13"
      Item(2).Caption =   "Cierres de Caja"
      Item(2).ControlCount=   43
      Item(2).Control(0)=   "grilla"
      Item(2).Control(1)=   "PusDesactivarCierre"
      Item(2).Control(2)=   "PusCerrarCaja"
      Item(2).Control(3)=   "vfecha"
      Item(2).Control(4)=   "lblIndicarUna"
      Item(2).Control(5)=   "vestado"
      Item(2).Control(6)=   "lblIndiqueUn"
      Item(2).Control(7)=   "PusMostrarTodo"
      Item(2).Control(8)=   "grilla2"
      Item(2).Control(9)=   "GroSaldosDe"
      Item(2).Control(10)=   "PusImprimir"
      Item(2).Control(11)=   "PushButton2"
      Item(2).Control(12)=   "PusPaso3"
      Item(2).Control(13)=   "PushButton6"
      Item(2).Control(14)=   "PusFinDe"
      Item(2).Control(15)=   "paraBalance"
      Item(2).Control(16)=   "ShoPaso1"
      Item(2).Control(17)=   "ShoPaso2"
      Item(2).Control(18)=   "ShortcutCaption1"
      Item(2).Control(19)=   "ShoPaso4"
      Item(2).Control(20)=   "GroListadosDiarios"
      Item(2).Control(21)=   "GroupBox7"
      Item(2).Control(22)=   "GroOtrosListados"
      Item(2).Control(23)=   "lblPuedeRealizar"
      Item(2).Control(24)=   "Label7"
      Item(2).Control(25)=   "Label8"
      Item(2).Control(26)=   "Label9"
      Item(2).Control(27)=   "barra"
      Item(2).Control(28)=   "grilla3"
      Item(2).Control(29)=   "ShoCierreMensual"
      Item(2).Control(30)=   "ShortcutCaption2"
      Item(2).Control(31)=   "PushButton7"
      Item(2).Control(32)=   "PushButton8"
      Item(2).Control(33)=   "PushButton9"
      Item(2).Control(34)=   "PushButton10"
      Item(2).Control(35)=   "PusNroComp"
      Item(2).Control(36)=   "log"
      Item(2).Control(37)=   "gnonro"
      Item(2).Control(38)=   "vvbarra"
      Item(2).Control(39)=   "cmdExp"
      Item(2).Control(40)=   "PushButton11"
      Item(2).Control(41)=   "ShortcutCaption4"
      Item(2).Control(42)=   "cmdReimprimir"
      Begin VB.CommandButton cmdReimprimir 
         Caption         =   "Reimprimir"
         Height          =   240
         Left            =   5715
         TabIndex        =   135
         Top             =   7650
         Width           =   1590
      End
      Begin VB.CommandButton cmdExp 
         Caption         =   "Exp"
         Height          =   795
         Left            =   6870
         TabIndex        =   129
         Top             =   6720
         Width           =   495
      End
      Begin MSComctlLib.ProgressBar vvbarra 
         Height          =   135
         Left            =   150
         TabIndex        =   128
         Top             =   8340
         Width           =   11625
         _ExtentX        =   20505
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gnonro 
         Height          =   795
         Left            =   420
         TabIndex        =   127
         Top             =   6720
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   1402
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin XtremeSuiteControls.ListBox log 
         Height          =   2355
         Left            =   11640
         TabIndex        =   126
         Top             =   4140
         Width           =   4005
         _Version        =   851968
         _ExtentX        =   7064
         _ExtentY        =   4154
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Niagara Engraved"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.PushButton PusNroComp 
         Height          =   345
         Left            =   6690
         TabIndex        =   125
         Top             =   7920
         Width           =   705
         _Version        =   851968
         _ExtentX        =   1244
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Nro Comp"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton10 
         Height          =   345
         Left            =   450
         TabIndex        =   121
         Top             =   1710
         Width           =   2775
         _Version        =   851968
         _ExtentX        =   4895
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Todos"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton9 
         Height          =   345
         Left            =   450
         TabIndex        =   120
         Top             =   1350
         Width           =   1305
         _Version        =   851968
         _ExtentX        =   2302
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Cierres diarios"
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton7 
         Height          =   375
         Left            =   150
         TabIndex        =   118
         Top             =   390
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cierre diarios"
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroListadosDiarios 
         Height          =   375
         Left            =   120
         TabIndex        =   99
         Top             =   7500
         Width           =   6705
         _Version        =   851968
         _ExtentX        =   11827
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Listados Diarios"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
         Begin VB.CheckBox chkComprobantesNo 
            Caption         =   "Solo comprobantes no impresos"
            Height          =   225
            Left            =   2655
            TabIndex        =   124
            Top             =   135
            Width           =   2775
         End
      End
      Begin XtremeSuiteControls.GroupBox paraBalance 
         Height          =   1695
         Left            =   7380
         TabIndex        =   89
         Top             =   5850
         Width           =   4995
         _Version        =   851968
         _ExtentX        =   8811
         _ExtentY        =   2990
         _StockProps     =   79
         Caption         =   "Infique el més del balance que desea generar."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Begin VB.TextBox vidanomes_label 
            BackColor       =   &H0000FFFF&
            Height          =   375
            Left            =   2340
            TabIndex        =   133
            Top             =   450
            Width           =   1635
         End
         Begin Project1.bsGradientLabel BsGDebeImprimir 
            Height          =   255
            Left            =   120
            Top             =   900
            Width           =   4755
            _ExtentX        =   8387
            _ExtentY        =   450
            Caption         =   "Debe imprimir el Balance y la composición de Saldo"
            BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Colour1         =   4210752
            Colour2         =   4210752
            CaptionAlignment=   1
         End
         Begin XtremeSuiteControls.PushButton PusComposiciónDe 
            Height          =   375
            Left            =   960
            TabIndex        =   108
            Top             =   1230
            Width           =   1695
            _Version        =   851968
            _ExtentX        =   2990
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Composición de Saldo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit vvmes 
            Height          =   285
            Left            =   1530
            TabIndex        =   93
            Top             =   270
            Width           =   720
            _Version        =   851968
            _ExtentX        =   1270
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.PushButton PusListo 
            Height          =   375
            Left            =   2730
            TabIndex        =   90
            Top             =   1230
            Width           =   915
            _Version        =   851968
            _ExtentX        =   1614
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Balance"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit vvano 
            Height          =   285
            Left            =   1530
            TabIndex        =   94
            Top             =   570
            Width           =   720
            _Version        =   851968
            _ExtentX        =   1270
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.PushButton PushButton12 
            Height          =   375
            Left            =   3750
            TabIndex        =   132
            Top             =   1230
            Width           =   1155
            _Version        =   851968
            _ExtentX        =   2037
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Imputaciones"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label lblMesA 
            Caption         =   "Mes a mostar: AAAAMM"
            Height          =   330
            Left            =   2340
            TabIndex        =   134
            Top             =   270
            Width           =   1725
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
            Height          =   375
            Left            =   180
            TabIndex        =   123
            Top             =   1230
            Width           =   765
            _Version        =   851968
            _ExtentX        =   1349
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Paso #1."
            ForeColor       =   1375373
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientColorLight=   128
            GradientColorDark=   255
            ForeColor       =   1375373
         End
         Begin XtremeSuiteControls.Label lblIngEl 
            Height          =   405
            Left            =   270
            TabIndex        =   92
            Top             =   480
            Width           =   975
            _Version        =   851968
            _ExtentX        =   1720
            _ExtentY        =   714
            _StockProps     =   79
            Caption         =   "Ing. el año:"
         End
         Begin XtremeSuiteControls.Label Label6 
            Height          =   405
            Left            =   270
            TabIndex        =   91
            Top             =   180
            Width           =   975
            _Version        =   851968
            _ExtentX        =   1720
            _ExtentY        =   714
            _StockProps     =   79
            Caption         =   "Ing. el mes:"
         End
      End
      Begin XtremeSuiteControls.PushButton PusImprimir 
         Height          =   375
         Left            =   930
         TabIndex        =   39
         Top             =   7890
         Width           =   2445
         _Version        =   851968
         _ExtentX        =   4313
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   " Imprimir Composición Saldos"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmListados.frx":099A
      End
      Begin XtremeSuiteControls.GroupBox GroSaldosDe 
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   3840
         Width           =   15405
         _Version        =   851968
         _ExtentX        =   27173
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Saldos de cajas:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
      End
      Begin XtremeSuiteControls.FlatEdit vestado 
         Height          =   315
         Left            =   3570
         TabIndex        =   33
         Top             =   840
         Width           =   2655
         _Version        =   851968
         _ExtentX        =   4683
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   375
         Left            =   -66460
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   2415
         _Version        =   851968
         _ExtentX        =   4260
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Ver transacciones - Borrar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GBRangoDeFechas 
         Height          =   7095
         Left            =   -69850
         TabIndex        =   1
         Top             =   570
         Visible         =   0   'False
         Width           =   13635
         _Version        =   851968
         _ExtentX        =   24051
         _ExtentY        =   12515
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         BorderStyle     =   2
         Begin XtremeSuiteControls.GroupBox GroupBox8 
            Height          =   585
            Left            =   5250
            TabIndex        =   110
            Top             =   4980
            Width           =   5085
            _Version        =   851968
            _ExtentX        =   8969
            _ExtentY        =   1032
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Begin XtremeSuiteControls.RadioButton RadSoloRecibos 
               Height          =   255
               Left            =   1380
               TabIndex        =   111
               Top             =   210
               Width           =   1245
               _Version        =   851968
               _ExtentX        =   2196
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Solo Recibos"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RadSoloPagos 
               Height          =   315
               Left            =   3000
               TabIndex        =   112
               Top             =   180
               Width           =   1995
               _Version        =   851968
               _ExtentX        =   3519
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "Solo Ordenes de Pagos"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RadTodos2 
               Height          =   315
               Left            =   180
               TabIndex        =   113
               Top             =   180
               Width           =   825
               _Version        =   851968
               _ExtentX        =   1455
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "Todos"
               UseVisualStyle  =   -1  'True
               Value           =   -1  'True
            End
         End
         Begin XtremeSuiteControls.PushButton PusSalos 
            Height          =   315
            Left            =   10740
            TabIndex        =   107
            Top             =   1260
            Width           =   885
            _Version        =   851968
            _ExtentX        =   1561
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Salos"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PusAyer 
            Height          =   315
            Left            =   10740
            TabIndex        =   85
            Top             =   120
            Width           =   885
            _Version        =   851968
            _ExtentX        =   1561
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Ayer"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.GroupBox GroupBox6 
            Height          =   375
            Left            =   7500
            TabIndex        =   79
            Top             =   3660
            Width           =   3255
            _Version        =   851968
            _ExtentX        =   5741
            _ExtentY        =   661
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Begin XtremeSuiteControls.RadioButton RadTodos 
               Height          =   195
               Left            =   90
               TabIndex        =   80
               Top             =   150
               Width           =   765
               _Version        =   851968
               _ExtentX        =   1349
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "Todos"
               UseVisualStyle  =   -1  'True
               Value           =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RadCancelados 
               Height          =   195
               Left            =   840
               TabIndex        =   81
               Top             =   150
               Width           =   1185
               _Version        =   851968
               _ExtentX        =   2090
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "Cancelados"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RadPendientes 
               Height          =   195
               Left            =   2040
               TabIndex        =   82
               Top             =   150
               Width           =   1155
               _Version        =   851968
               _ExtentX        =   2037
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "Pendientes"
               UseVisualStyle  =   -1  'True
            End
         End
         Begin XtremeSuiteControls.CheckBox chkFecha 
            Height          =   255
            Left            =   210
            TabIndex        =   67
            Top             =   150
            Width           =   1965
            _Version        =   851968
            _ExtentX        =   3466
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Todas las Fechas"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.GroupBox GroupBox5 
            Height          =   285
            Left            =   210
            TabIndex        =   54
            Top             =   690
            Width           =   10245
            _Version        =   851968
            _ExtentX        =   18071
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Tipos de listados: "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            BorderStyle     =   1
         End
         Begin XtremeSuiteControls.GroupBox GroupBox4 
            Height          =   165
            Left            =   270
            TabIndex        =   53
            Top             =   5580
            Width           =   10005
            _Version        =   851968
            _ExtentX        =   17648
            _ExtentY        =   291
            _StockProps     =   79
            Caption         =   "Filtros secundarios: "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            BorderStyle     =   1
         End
         Begin XtremeSuiteControls.GroupBox GroFiltros 
            Height          =   1455
            Left            =   240
            TabIndex        =   40
            Top             =   2280
            Width           =   10185
            _Version        =   851968
            _ExtentX        =   17965
            _ExtentY        =   2566
            _StockProps     =   79
            Caption         =   "Filtros:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            BorderStyle     =   1
            Begin VB.TextBox vcliprovee 
               Height          =   345
               Left            =   3960
               TabIndex        =   42
               Top             =   330
               Width           =   6105
            End
            Begin VB.PictureBox Picture3 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   2190
               ScaleHeight     =   315
               ScaleWidth      =   495
               TabIndex        =   41
               Top             =   300
               Width           =   495
            End
            Begin XtremeSuiteControls.FlatEdit vctipovalor 
               Height          =   315
               Left            =   2100
               TabIndex        =   44
               Top             =   690
               Width           =   1215
               _Version        =   851968
               _ExtentX        =   2143
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.PushButton btv 
               Height          =   285
               Left            =   3420
               TabIndex        =   45
               Tag             =   "TipoValor"
               Top             =   690
               Width           =   405
               _Version        =   851968
               _ExtentX        =   714
               _ExtentY        =   503
               _StockProps     =   79
               Caption         =   "..."
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit vtipovalor 
               Height          =   345
               Left            =   3930
               TabIndex        =   46
               Top             =   690
               Width           =   6135
               _Version        =   851968
               _ExtentX        =   10821
               _ExtentY        =   609
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.PushButton bp 
               Height          =   285
               Left            =   3420
               TabIndex        =   48
               Tag             =   "TipoValor"
               Top             =   360
               Width           =   405
               _Version        =   851968
               _ExtentX        =   714
               _ExtentY        =   503
               _StockProps     =   79
               Caption         =   "..."
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit vcbancocaja 
               Height          =   315
               Left            =   2100
               TabIndex        =   49
               Top             =   1050
               Width           =   1215
               _Version        =   851968
               _ExtentX        =   2143
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.PushButton bcb 
               Height          =   285
               Left            =   3420
               TabIndex        =   50
               Tag             =   "CajaBanco"
               Top             =   1080
               Width           =   405
               _Version        =   851968
               _ExtentX        =   714
               _ExtentY        =   503
               _StockProps     =   79
               Caption         =   "..."
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit vdbancocaja 
               Height          =   345
               Left            =   3930
               TabIndex        =   51
               Top             =   1050
               Width           =   6135
               _Version        =   851968
               _ExtentX        =   10821
               _ExtentY        =   609
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.Label lblAltaCaja 
               Height          =   195
               Index           =   7
               Left            =   180
               TabIndex        =   52
               Top             =   1080
               Width           =   1815
               _Version        =   851968
               _ExtentX        =   3201
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "Caja / Banco: "
               ForeColor       =   255
               Alignment       =   1
               Transparent     =   -1  'True
            End
            Begin XtremeSuiteControls.Label lblAltaCaja 
               Height          =   195
               Index           =   10
               Left            =   510
               TabIndex        =   47
               Top             =   720
               Width           =   1485
               _Version        =   851968
               _ExtentX        =   2619
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "Tipo de Valores:"
               ForeColor       =   255
               Alignment       =   1
               Transparent     =   -1  'True
            End
            Begin VB.Label lblAsientos 
               Alignment       =   1  'Right Justify
               Caption         =   "Personas / Entidades:"
               ForeColor       =   &H000000FF&
               Height          =   285
               Index           =   11
               Left            =   -240
               TabIndex        =   43
               Top             =   390
               Width           =   2265
            End
         End
         Begin XtremeSuiteControls.FlatEdit vnrocheque 
            Height          =   315
            Left            =   1350
            TabIndex        =   25
            Top             =   6690
            Width           =   8985
            _Version        =   851968
            _ExtentX        =   15849
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vmarcainterna 
            Height          =   315
            Left            =   1350
            TabIndex        =   23
            Top             =   6300
            Width           =   8985
            _Version        =   851968
            _ExtentX        =   15849
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Enabled         =   0   'False
         End
         Begin VB.TextBox vnrointerno 
            Height          =   315
            Left            =   1365
            TabIndex        =   21
            Top             =   5940
            Width           =   8985
         End
         Begin VB.ComboBox cboAgrupado 
            Height          =   315
            ItemData        =   "frmListados.frx":0F34
            Left            =   1560
            List            =   "frmListados.frx":0F36
            TabIndex        =   17
            Text            =   "No agrupar"
            Top             =   1860
            Width           =   8835
         End
         Begin VB.ComboBox cboOrdenado 
            Height          =   315
            ItemData        =   "frmListados.frx":0F38
            Left            =   1560
            List            =   "frmListados.frx":0F3A
            TabIndex        =   11
            Text            =   "Fecha"
            Top             =   1440
            Width           =   8835
         End
         Begin VB.ComboBox vtipolistado 
            Height          =   315
            ItemData        =   "frmListados.frx":0F3C
            Left            =   1560
            List            =   "frmListados.frx":0F3E
            TabIndex        =   2
            Text            =   "Todos"
            Top             =   1020
            Width           =   8805
         End
         Begin Aplisoft_CajasDeTexto.TxF dtpFecha 
            Height          =   345
            Index           =   0
            Left            =   4470
            TabIndex        =   3
            Top             =   120
            Width           =   1980
            _ExtentX        =   3493
            _ExtentY        =   609
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
         Begin Aplisoft_CajasDeTexto.TxF dtpFecha 
            Height          =   345
            Index           =   1
            Left            =   7830
            TabIndex        =   4
            Top             =   120
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   609
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
         Begin XtremeSuiteControls.FlatEdit vctm 
            Height          =   315
            Left            =   2340
            TabIndex        =   56
            Top             =   3750
            Width           =   1215
            _Version        =   851968
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.PushButton PushButton3 
            Height          =   285
            Left            =   3690
            TabIndex        =   57
            Tag             =   "CajaBanco"
            Top             =   3750
            Width           =   405
            _Version        =   851968
            _ExtentX        =   714
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit vdtm 
            Height          =   345
            Left            =   4170
            TabIndex        =   58
            Top             =   3720
            Width           =   3285
            _Version        =   851968
            _ExtentX        =   5794
            _ExtentY        =   609
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vctasCodigo 
            Height          =   315
            Left            =   2340
            TabIndex        =   69
            Top             =   4140
            Width           =   2745
            _Version        =   851968
            _ExtentX        =   4842
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.PushButton PushButton5 
            Height          =   285
            Left            =   5130
            TabIndex        =   70
            Tag             =   "CajaBanco"
            Top             =   4170
            Width           =   405
            _Version        =   851968
            _ExtentX        =   714
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit vCtasDescrip 
            Height          =   345
            Left            =   5610
            TabIndex        =   71
            Top             =   4110
            Width           =   4725
            _Version        =   851968
            _ExtentX        =   8334
            _ExtentY        =   609
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vcomentario 
            Height          =   345
            Left            =   2340
            TabIndex        =   76
            Top             =   4590
            Width           =   3900
            _Version        =   851968
            _ExtentX        =   6879
            _ExtentY        =   609
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vnrocomprobante 
            Height          =   345
            Left            =   2340
            TabIndex        =   83
            Top             =   4980
            Width           =   2775
            _Version        =   851968
            _ExtentX        =   4895
            _ExtentY        =   609
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.PushButton PusSemanal 
            Height          =   315
            Left            =   10740
            TabIndex        =   86
            Top             =   510
            Width           =   885
            _Version        =   851968
            _ExtentX        =   1561
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Semanal"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PusMensual 
            Height          =   315
            Left            =   10740
            TabIndex        =   87
            Top             =   870
            Width           =   885
            _Version        =   851968
            _ExtentX        =   1561
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Mensual"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit vcomentario2 
            Height          =   345
            Left            =   6255
            TabIndex        =   136
            Top             =   4590
            Width           =   4125
            _Version        =   851968
            _ExtentX        =   7276
            _ExtentY        =   609
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.Label lblParteDel 
            Height          =   195
            Index           =   0
            Left            =   -90
            TabIndex        =   84
            Top             =   5010
            Width           =   2415
            _Version        =   851968
            _ExtentX        =   4260
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Nro de comprobante: "
            ForeColor       =   255
            Alignment       =   1
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblParteDel 
            Height          =   195
            Index           =   2
            Left            =   -60
            TabIndex        =   77
            Top             =   4620
            Width           =   2385
            _Version        =   851968
            _ExtentX        =   4207
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Escriba parte de un comentario: "
            ForeColor       =   255
            Alignment       =   1
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblAltaCaja 
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   72
            Top             =   4170
            Width           =   1815
            _Version        =   851968
            _ExtentX        =   3201
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Ctas Contable:"
            ForeColor       =   255
            Alignment       =   1
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblAltaCaja 
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   59
            Top             =   3780
            Width           =   1815
            _Version        =   851968
            _ExtentX        =   3201
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Tipo Movimiento:"
            ForeColor       =   255
            Alignment       =   1
            Transparent     =   -1  'True
         End
         Begin VB.Label Label5 
            Caption         =   "Nro. Cheque:"
            Height          =   225
            Left            =   210
            TabIndex        =   24
            Top             =   6690
            Width           =   1035
         End
         Begin VB.Label Label4 
            Caption         =   "Marca interna:"
            Height          =   225
            Left            =   150
            TabIndex        =   22
            Top             =   6330
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Nro interno:"
            Height          =   225
            Left            =   330
            TabIndex        =   20
            Top             =   6000
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "Agrupado por:"
            Height          =   225
            Left            =   330
            TabIndex        =   18
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Ordenado por:"
            Height          =   225
            Left            =   330
            TabIndex        =   12
            Top             =   1530
            Width           =   1095
         End
         Begin XtremeSuiteControls.Label lblFechas 
            Height          =   195
            Index           =   0
            Left            =   3300
            TabIndex        =   7
            Top             =   180
            Width           =   1035
            _Version        =   851968
            _ExtentX        =   1826
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Fecha Desde:"
         End
         Begin XtremeSuiteControls.Label lblFechas 
            Height          =   195
            Index           =   1
            Left            =   6660
            TabIndex        =   6
            Top             =   210
            Width           =   1020
            _Version        =   851968
            _ExtentX        =   1799
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Fecha Hasta :"
         End
         Begin VB.Label lblTipoListado 
            Caption         =   "Tipo Listado:"
            Height          =   225
            Left            =   450
            TabIndex        =   5
            Top             =   1050
            Width           =   1125
         End
      End
      Begin XtremeSuiteControls.PushButton cmdFiltrar 
         Height          =   495
         Left            =   -69640
         TabIndex        =   8
         Top             =   7710
         Visible         =   0   'False
         Width           =   10155
         _Version        =   851968
         _ExtentX        =   17912
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Filtrar Movimientos"
         UseVisualStyle  =   -1  'True
         BorderGap       =   10
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   7545
         Left            =   -69910
         TabIndex        =   9
         Top             =   900
         Visible         =   0   'False
         Width           =   15705
         _Version        =   851968
         _ExtentX        =   27702
         _ExtentY        =   13309
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.PushButton PusExportarA 
            Height          =   405
            Left            =   13470
            TabIndex        =   114
            Top             =   5460
            Width           =   2145
            _Version        =   851968
            _ExtentX        =   3784
            _ExtentY        =   714
            _StockProps     =   79
            Caption         =   "Exportar a Excel"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit1 
            Height          =   315
            Left            =   810
            TabIndex        =   73
            Top             =   150
            Width           =   6945
            _Version        =   851968
            _ExtentX        =   12250
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.GroupBox fcancelaValores 
            Height          =   1395
            Left            =   7800
            TabIndex        =   61
            Top             =   90
            Visible         =   0   'False
            Width           =   7905
            _Version        =   851968
            _ExtentX        =   13944
            _ExtentY        =   2461
            _StockProps     =   79
            Caption         =   "Ingrese la Caja o Banco donde quiere transferir los Vales seleccionados: "
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Begin XtremeSuiteControls.PushButton PusEjecutar 
               Height          =   375
               Left            =   5400
               TabIndex        =   65
               Top             =   900
               Width           =   2325
               _Version        =   851968
               _ExtentX        =   4101
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "Ejecutar"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit vccaja 
               Height          =   315
               Left            =   120
               TabIndex        =   62
               Top             =   450
               Width           =   1335
               _Version        =   851968
               _ExtentX        =   2355
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.PushButton pbCarga11 
               Height          =   285
               Left            =   1500
               TabIndex        =   63
               Tag             =   "CajaBanco"
               Top             =   480
               Width           =   405
               _Version        =   851968
               _ExtentX        =   714
               _ExtentY        =   503
               _StockProps     =   79
               Caption         =   "..."
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit vdcaja 
               Height          =   315
               Left            =   1950
               TabIndex        =   64
               Top             =   480
               Width           =   5775
               _Version        =   851968
               _ExtentX        =   10186
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
            End
         End
         Begin Grid.KlexGrid KlexDetalle 
            Height          =   4875
            Left            =   60
            TabIndex        =   10
            Top             =   540
            Width           =   15585
            _ExtentX        =   27490
            _ExtentY        =   8599
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
            MouseIcon       =   "frmListados.frx":0F40
            SelectionMode   =   1
         End
         Begin Grid.KlexGrid gdetalle 
            Height          =   1545
            Left            =   60
            TabIndex        =   75
            Top             =   5910
            Width           =   15585
            _ExtentX        =   27490
            _ExtentY        =   2725
            EnterKeyBehaviour=   0
            BackColorAlternate=   16777215
            GridLinesFixed  =   2
            AllowUserResizing=   1
            BackColor       =   16777215
            BackColorBkg    =   4210752
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
            ForeColor       =   4210752
            GridColorFixed  =   8421504
            MouseIcon       =   "frmListados.frx":0F5C
            SelectionMode   =   1
            WordWrap        =   -1  'True
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid g 
            Height          =   4335
            Left            =   150
            TabIndex        =   122
            Top             =   960
            Width           =   15375
            _ExtentX        =   27120
            _ExtentY        =   7646
            _Version        =   393216
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin XtremeSuiteControls.Label lblCambieEl 
            Height          =   135
            Left            =   120
            TabIndex        =   106
            Top             =   5490
            Width           =   7125
            _Version        =   851968
            _ExtentX        =   12568
            _ExtentY        =   238
            _StockProps     =   79
            Caption         =   "Cambie el estado de los VALES haciendo doble clic sobre el Vale seleccionado"
            ForeColor       =   255
         End
         Begin XtremeSuiteControls.Label lblBuscar 
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   210
            Width           =   1125
            _Version        =   851968
            _ExtentX        =   1984
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Buscar:"
         End
      End
      Begin XtremeSuiteControls.PushButton cmdImprimir 
         Height          =   375
         Index           =   0
         Left            =   -68170
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Imprimir Listado"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdImprimir 
         Height          =   375
         Index           =   1
         Left            =   -69880
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla 
         Height          =   1755
         Left            =   420
         TabIndex        =   28
         Top             =   2040
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   3096
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   5
         BackColorSel    =   65280
         ForeColorSel    =   0
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).BandIndent=   9
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin XtremeSuiteControls.PushButton PusDesactivarCierre 
         Height          =   405
         Left            =   6660
         TabIndex        =   29
         Top             =   900
         Width           =   2385
         _Version        =   851968
         _ExtentX        =   4207
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Abrir Caja"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PusCerrarCaja 
         Height          =   375
         Left            =   3570
         TabIndex        =   30
         Top             =   1230
         Width           =   1305
         _Version        =   851968
         _ExtentX        =   2302
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cerrar Caja"
         UseVisualStyle  =   -1  'True
      End
      Begin MSComCtl2.DTPicker vfecha 
         Height          =   375
         Left            =   3570
         TabIndex        =   31
         Top             =   420
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   661
         _Version        =   393216
         Format          =   71106561
         CurrentDate     =   41887
      End
      Begin XtremeSuiteControls.PushButton PusMostrarTodo 
         Height          =   375
         Left            =   6660
         TabIndex        =   35
         Top             =   450
         Width           =   2385
         _Version        =   851968
         _ExtentX        =   4207
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Mostrar todo"
         UseVisualStyle  =   -1  'True
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla2 
         Height          =   2535
         Left            =   420
         TabIndex        =   37
         Top             =   4110
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   4471
         _Version        =   393216
         BackColorSel    =   65280
         ForeColorSel    =   0
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).BandIndent=   9
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   285
         Left            =   13260
         TabIndex        =   55
         Top             =   8160
         Width           =   2505
         _Version        =   851968
         _ExtentX        =   4419
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Imprimir Movimientos de Ctas."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   375
         Left            =   -62350
         TabIndex        =   60
         Top             =   450
         Visible         =   0   'False
         Width           =   3585
         _Version        =   851968
         _ExtentX        =   6324
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Dar de baja los VALES seleccionados"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PusReimprimirComprobante 
         Height          =   375
         Left            =   -58480
         TabIndex        =   66
         Top             =   450
         Visible         =   0   'False
         Width           =   2295
         _Version        =   851968
         _ExtentX        =   4048
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Reimprimir Comprobante"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PusPaso3 
         Height          =   285
         Left            =   13260
         TabIndex        =   68
         Top             =   7740
         Width           =   2475
         _Version        =   851968
         _ExtentX        =   4366
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Imprimir movimientos de Caja"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton6 
         Height          =   345
         Left            =   4275
         TabIndex        =   78
         Top             =   7920
         Width           =   2355
         _Version        =   851968
         _ExtentX        =   4154
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Imprimir movimientos del día"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PusFinDe 
         Height          =   405
         Left            =   7890
         TabIndex        =   88
         Top             =   7830
         Width           =   3915
         _Version        =   851968
         _ExtentX        =   6906
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "(Fin de mes) - Impresión de comprobantes mensuales"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox7 
         Height          =   375
         Left            =   7470
         TabIndex        =   100
         Top             =   7560
         Width           =   4515
         _Version        =   851968
         _ExtentX        =   7964
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Listados Fin de mes"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
      End
      Begin XtremeSuiteControls.GroupBox GroOtrosListados 
         Height          =   375
         Left            =   12420
         TabIndex        =   101
         Top             =   6900
         Width           =   3225
         _Version        =   851968
         _ExtentX        =   5689
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Otros Listados"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla3 
         Height          =   1605
         Left            =   11610
         TabIndex        =   115
         Top             =   2160
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   2831
         _Version        =   393216
         BackColor       =   4210752
         ForeColor       =   14737632
         Cols            =   5
         BackColorSel    =   65280
         ForeColorSel    =   0
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).BandIndent=   9
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin XtremeSuiteControls.PushButton PushButton8 
         Height          =   345
         Left            =   1740
         TabIndex        =   119
         Top             =   1350
         Width           =   1485
         _Version        =   851968
         _ExtentX        =   2619
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Cierre mensuales"
         BackColor       =   16761024
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ProgressBar barra 
         Height          =   255
         Left            =   12810
         TabIndex        =   109
         Top             =   6600
         Width           =   2745
         _Version        =   851968
         _ExtentX        =   4842
         _ExtentY        =   450
         _StockProps     =   93
         Appearance      =   3
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton PushButton11 
         Height          =   315
         Left            =   13260
         TabIndex        =   130
         Top             =   7320
         Width           =   2475
         _Version        =   851968
         _ExtentX        =   4366
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Libro de Imputaciones"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton13 
         Height          =   375
         Left            =   -55645
         TabIndex        =   137
         Top             =   450
         Visible         =   0   'False
         Width           =   1305
         _Version        =   851968
         _ExtentX        =   2302
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Reimp. Todos"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption4 
         Height          =   375
         Left            =   12390
         TabIndex        =   131
         Top             =   7260
         Width           =   765
         _Version        =   851968
         _ExtentX        =   1349
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Paso #5."
         ForeColor       =   1375373
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   128
         GradientColorDark=   255
         ForeColor       =   1375373
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   255
         Left            =   4110
         TabIndex        =   117
         Top             =   1740
         Width           =   7395
         _Version        =   851968
         _ExtentX        =   13044
         _ExtentY        =   450
         _StockProps     =   14
         Caption         =   "Cierres diario:"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   0
         ForeColor       =   16777215
      End
      Begin XtremeShortcutBar.ShortcutCaption ShoCierreMensual 
         Height          =   285
         Left            =   11610
         TabIndex        =   116
         Top             =   1830
         Width           =   4095
         _Version        =   851968
         _ExtentX        =   7223
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "Cierres mensual:"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   0
         ForeColor       =   16777215
      End
      Begin XtremeSuiteControls.Label Label9 
         Height          =   375
         Left            =   9210
         TabIndex        =   105
         Top             =   1380
         Width           =   6555
         _Version        =   851968
         _ExtentX        =   11562
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "- Consultar un cierre de caja. Haciendo clic en la grilla y luego generar los listados que desea."
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   435
         Left            =   9210
         TabIndex        =   104
         Top             =   1050
         Width           =   4725
         _Version        =   851968
         _ExtentX        =   8334
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "- Cierre de caja a fin de mes. Pasos 1,2,3,4, clic en el botón <Fin de mes>"
      End
      Begin XtremeSuiteControls.Label Label7 
         Height          =   375
         Left            =   9210
         TabIndex        =   103
         Top             =   780
         Width           =   3555
         _Version        =   851968
         _ExtentX        =   6271
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "- Cierre de caja diaria: Pasos 1,2,3,4"
      End
      Begin XtremeSuiteControls.Label lblPuedeRealizar 
         Height          =   435
         Left            =   9240
         TabIndex        =   102
         Top             =   360
         Width           =   5505
         _Version        =   851968
         _ExtentX        =   9710
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Puede realizar algunas de estras tres acciones:"
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
      End
      Begin XtremeShortcutBar.ShortcutCaption ShoPaso4 
         Height          =   375
         Left            =   3480
         TabIndex        =   98
         Top             =   7860
         Width           =   765
         _Version        =   851968
         _ExtentX        =   1349
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Paso #4."
         ForeColor       =   1375373
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   128
         GradientColorDark=   255
         ForeColor       =   1375373
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Left            =   120
         TabIndex        =   97
         Top             =   7860
         Width           =   765
         _Version        =   851968
         _ExtentX        =   1349
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Paso #3."
         ForeColor       =   1375373
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   128
         GradientColorDark=   255
         ForeColor       =   1375373
      End
      Begin XtremeShortcutBar.ShortcutCaption ShoPaso2 
         Height          =   375
         Left            =   5460
         TabIndex        =   96
         Top             =   1200
         Width           =   765
         _Version        =   851968
         _ExtentX        =   1349
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Paso #2."
         ForeColor       =   1375373
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   128
         GradientColorDark=   255
         ForeColor       =   1375373
      End
      Begin XtremeShortcutBar.ShortcutCaption ShoPaso1 
         Height          =   375
         Left            =   5460
         TabIndex        =   95
         Top             =   420
         Width           =   765
         _Version        =   851968
         _ExtentX        =   1349
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Paso #1."
         ForeColor       =   1375373
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   128
         GradientColorDark=   255
         ForeColor       =   1375373
      End
      Begin XtremeSuiteControls.Label lblIndiqueUn 
         Height          =   345
         Left            =   150
         TabIndex        =   34
         Top             =   780
         Width           =   3255
         _Version        =   851968
         _ExtentX        =   5741
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Indique un comentario para el cierre de caja: "
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblIndicarUna 
         Height          =   345
         Left            =   30
         TabIndex        =   32
         Top             =   450
         Width           =   3315
         _Version        =   851968
         _ExtentX        =   5847
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Fecha del cierre de  CAJA :"
         Alignment       =   1
      End
   End
End
Attribute VB_Name = "frmBancoCajaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim impDirecto, cierremensual As Boolean
Dim vgid, vidcierecaja As Long
Dim vsd, vsc As Double
Dim vFiltro2, vaccion As String
Dim ars(20) As String
Dim vr As Integer
Dim vidcajacierre As Integer
Dim unicavez As Integer
Public vGidhastabm, vGiddesdebm, vGidasientos, vidasientosdesde, vidasientoshasta As Long
Dim vsqlFecha, vvviene As String
Dim vidanomes, vidanomesAnterior As Long
Dim vidH, vidd, vidAdesde, vidAhasta As Long
Dim tsd, tsc As Double
Dim vtGralD, vtGralC As Double

Dim reimprimirTodo As Integer


Private Sub bcb_Click()
Call fbuscarGrilla("(select * from bancos where not EsCaja='B') as b", "Descripcion", "idBancos", Me.vdbancocaja.Name, Me, , False)
End Sub

Private Sub bp_Click()
Call fbuscarGrilla("proveedores", "Nombre", "Codigo", Me.vcliprovee.Name, Me, , False)
End Sub

Private Sub btv_Click()
Call fbuscarGrilla("tipovalor", "TipoValor", "idTipoValor", Me.vtipovalor.Name, Me, , False)
End Sub


Private Sub ImprimeCajaConDetalle()
On Error Resume Next

With frmMovientosDiarios

    .chkFechas.Value = 0
    
    .dtpCuentas(0).Value = Me.vfecha
    .dtpCuentas(1).Value = Me.vfecha
    
    .vGidasientosdesde = grilla.TextMatrix(vr, 4)
    .vGidasientoshasta = grilla.TextMatrix(vr, 7)
    
     
    
    Call .cmdEjecutar_Click
    

End With


Unload frmMovientosDiarios

If Err Then Exit Sub
End Sub
    

Private Sub cmdExp_Click()
Call grillaToExcel(Me.gnonro)
End Sub

Public Sub cmdImprimir_Click(Index As Integer)
On Error Resume Next
    
    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox "Prepare la Impresora !!!", vbInformation, "Mensaje ..."

    If Index = 0 Then
    
    Else
    
        With Mantenimiento.rsSaldoBancos
                 
            If .State = 1 Then .Close
        
           ' .Source = "SHAPE {SELECT * FROM BancosMovimientos WHERE (Fecha >= '" & strfechaMySQL(dtpFecha(0).Value) & "' AND Fecha <= '" & strfechaMySQL(dtpFecha(1).Value) & "') GROUP BY NroInterno, idBancos, idBancosCuentas;}  AS ListadoBancosCaja APPEND ({SELECT AsientosDetalle.*, Cuentas.Cuenta FROM AsientosDetalle LEFT JOIN Cuentas ON AsientosDetalle.CodigoCuenta=Cuentas.CodigoCuenta}  AS DetalleBancosCaja RELATE 'NroAsiento' TO 'Numero') AS DetalleBancosCaja"
        
           ' .Source = fsqlreporte()
            
            If .State = 0 Then .Open
            .Close
            .Open
        End With
    
    
    
    If Left(Me.vtipolistado.Text, 12) = "Agrupado por" Then
    
            With drBancosSaldo
            
                If impDirecto Then
                    Call .PrintReport(False, rptRangeAllPages)
                    Unload .object
                    impDirecto = False
                Else
                    
                    If vidanomes > 0 Then
                      .Sections("TituloEmpresa").Controls("etitulo").Caption = "Saldo de composición de cajas correspondiente al mes: " + Format(vidanomes, "0000/00")
                    Else
                        .Sections("TituloEmpresa").Controls("etitulo").Caption = "Saldos de Cajas - Bancos. Período: " + Str(Me.dtpFecha(0).Value) + " - " + Str(Me.dtpFecha(1).Value)
                    End If
                    .Show
                End If
                
            End With
    End If
    
    
    If Me.vtipolistado.Text = "Con detalles de Movimientos contables" Then
    
    With Mantenimiento.rsListadoBancosCaja
                 
            
            If .State = 1 Then .Close
        
            '.Source = "SHAPE {SELECT * FROM BancosMovimientos WHERE (Fecha >= '" & strfechaMySQL(dtpFecha(0).Value) & "' AND Fecha <= '" & strfechaMySQL(dtpFecha(1).Value) & "') GROUP BY NroInterno, idBancos, idBancosCuentas;}  AS ListadoBancosCaja APPEND ({SELECT AsientosDetalle.*, Cuentas.Cuenta FROM AsientosDetalle LEFT JOIN Cuentas ON AsientosDetalle.CodigoCuenta=Cuentas.CodigoCuenta}  AS DetalleBancosCaja RELATE 'NroAsiento' TO 'Numero') AS DetalleBancosCaja"
        
             .Source = fsqlreportedetalle
           '  .DataSource = "ListadoBancosCaja"
        
            If .State = 0 Then .Open
                .Close
                .Open
            End With
    

            With drBancoCajaDetalle
            
              If impDirecto Then
                    Call .PrintReport(False, rptRangeAllPages)
                    Unload .object
                    impDirecto = False
                Else
                    .Show
                End If
            End With
    End If
    

    '--------------------------------- Todos ---------------------
    If Me.vtipolistado.Text = "Todos" Then
    
    With Mantenimiento.rsMoviBC
            If .State = 1 Then .Close
               ' .Source = fBancoCaja(Me.dtpFecha(0).Value, Me.dtpFecha(1))
        
            If .State = 0 Then .Open
                .Close
                .Open
            End With
    

            With drBancosCaja
                
                .Sections("TituloEmpresa").Controls("lblSaldoAnterior").Caption = Format(Me.KlexDetalle.TextMatrix(1, 9), "###,###,##0.00")
                .Sections("TituloEmpresa").Controls("lblFechaDesde").Caption = Me.dtpFecha(0).Value
                .Sections("TituloEmpresa").Controls("lblFechaHasta").Caption = Me.dtpFecha(1).Value
                .Show
            End With
    End If
    ' -------------------------------------------------------------
    
    
    
    End If

If Err Then GrabarLog "cmdImprimir_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Function fsqlreportedetalle() As String
fsqlreportedetalle = "SHAPE {SELECT * FROM BancosMovimientos inner join bancos on bancosmovimientos.idBancos = bancos.idBancos Where " & _
" (bancosmovimientos.Fecha >= '" & strfechaMySQL(dtpFecha(0).Value) & "' AND bancosmovimientos.Fecha <= '" & strfechaMySQL(dtpFecha(1).Value) & "') " & _
" GROUP BY NroInterno;}  AS ListadoBancosCaja APPEND ({SELECT AsientosDetalle.*, Cuentas.Cuenta FROM AsientosDetalle LEFT JOIN Cuentas ON AsientosDetalle.CodigoCuenta=Cuentas.CodigoCuenta}  AS DetalleBancosCaja RELATE 'NroAsiento' TO 'Numero') AS DetalleBancosCaja"
End Function

Function fsqlreporte() As String
Dim vgrupo, vcb As String

'If Me.vtipolistado.Text = "Agrupado por Fecha" Then
'    vgrupo = "bancos.fecha"
'Else
'    vgrupo = "bancos.idbancos"
'End If


'-------------------
vcb = ""

vgrupo = "bancos.idbancos"

If Me.vtipolistado.Text = "Agrupado por Fecha" Then vgrupo = "bancos.fecha"


If Me.vtipolistado.Text = "Agrupado por Bancos" Then
    vgrupo = "bancos.idbancos"
    vcb = " and bancos.escaja='N' "
End If

If Me.vtipolistado.Text = "Agrupado por Cajas" Then
    vgrupo = "bancos.idbancos"
    vcb = " and bancos.escaja='S' "
End If


If Me.vtipolistado.Text = "Con detalles de Movimientos contables" Then
    vgrupo = "bancos.idbancos"
    vcb = ""
End If

'----------------

fsqlreporte = "SELECT" & _
  " sum(bancosmovimientos.Debito) AS d," & _
  " sum(bancosmovimientos.Credito) AS c," & _
  " sum(bancosmovimientos.Debito - bancosmovimientos.Credito) AS saldo," & _
  " bancosmovimientos.idBancos , " & _
  " bancos.descripcion , " & _
  " bancosmovimientos.fecha " & _
" From" & _
 " bancosmovimientos " & _
 " INNER JOIN bancos ON (bancosmovimientos.idBancos=bancos.idBancos)" & _
" Where " & _
  " (bancosmovimientos.Fecha >= '" & strfechaMySQL(dtpFecha(0).Value) & "' AND bancosmovimientos.Fecha <= '" & strfechaMySQL(dtpFecha(1).Value) & "') " + vcb & _
" Group By " + vgrupo
 ' " bancos.idbancos "
End Function
Private Sub cmdCerrar_Click()
On Error Resume Next
    
    Unload Me

If Err Then GrabarLog "cmdCerrar_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub


Function fBancoCaja(vfdesde As Date, vfhasta As Date) As String

Dim vcb, vgrupo, vsql1, vsql2  As String

vsqlFecha = ""


If Me.chkFecha Then
    vsqlFecha = " 1=1 " + vFiltro2 + " "

Else
    vsqlFecha = " (Fecha >= '" & strfechaMySQL(vfdesde) & "' AND Fecha <= '" & strfechaMySQL(vfhasta) & "') " + vFiltro2 + " "
End If

'-------------------
vcb = ""


If Me.RadSoloPagos.Value = True Then
    vsqlFecha = vsqlFecha + " and credito > 0 "
End If

If Me.RadSoloRecibos = True Then
    vsqlFecha = vsqlFecha + " and debito > 0 "
End If



If Me.vtipolistado.Text = "Agrupado por Fecha" Then vgrupo = "bb.fecha"


If Me.vtipolistado.Text = "Agrupado por Bancos" Then
    vgrupo = "bb.idbancos"
    vcb = " and bb.escaja='N' "
End If

If Me.vtipolistado.Text = "Agrupado por Cajas" Then
    vgrupo = "bb.idbancos"
    vcb = " and bb.escaja='S' "
End If


If Me.vtipolistado.Text = "Con detalles de Movimientos contables" Then
    vgrupo = "bb.idbancos"
    vcb = ""
End If

Dim sqlComentario As String

sqlComentario = ""

If Not vcomentario = "" Then

    sqlComentario = " and (comentario like '%" + vcomentario + "%' or comentario2 like '%" + Me.vcomentario + "%')"

End If



'----------------

If Not Me.vtipolistado.Text = "Todos" Then


       ' MsgBox "Este listado se muestra solo en formato impresión", vbInformation
        
        Exit Function


        fBancoCaja = "select  bm.TipoMovimiento, bm.idBancosMovimientos, bm.nrocomprobante, bm.codpersona,  bm.`Fecha`,   bm.`idBancos`,   bb.descripcion,   bm.`idBancosCuentas`," & _
         "   sum(bm.`Debito`) as D ,   sum(bm.`Credito`) as C,   (sum(bm.`Debito`) -  sum(bm.`Credito`)) as saldo, bm.`NroInterno`,  bm.`TipoMovimiento`, bm.`NroCheque`, bm.comentario, bm.conciliado  from ba" & _
         "ncosmovimientos  bm   left join bancoscuentas b on      bm.idbancoscuentas = b.i" & _
         "dbancoscuentas   inner join bancos bb on     bm.`idBancos` = bb.`idBancos`      " & _
         " where 1=1 and " + _
         vsqlFecha & _
         vcb + " group by " + vgrupo + " order by bm.`Fecha` , bm.idbancosmovimientos "
         
         
Else

       Select Case cboAgrupado.Text
       
        Case "No agrupar"
       
       
        If Me.vctasCodigo.Text = "" Then
        
                    fBancoCaja = "select  pp.nombre, bm.TipoMovimiento,bm.idBancosMovimientos, bm.nrocomprobante, bm.codpersona, bm.`Fecha`,   bm.`idBancos`,   bb.descripcion,   bm.`idBancosCuentas`," & _
                     "   bm.`Debito`,   bm.`Credito` ,   bm.saldo, bm.`NroInterno`,  bm.`TipoMovimiento`, bm.`NroCheque`, bm.comentario , bm.conciliado from ba" & _
                     "ncosmovimientos  bm   left join bancoscuentas b on      bm.idbancoscuentas = b.i" & _
                     "dbancoscuentas  left join bancos bb on     bm.`idBancos` = bb.`idBancos`      " & _
                     " left join  proveedores pp on  pp.`codigo` = bm.`codPersona`      " & _
                     " where  " + _
                    vsqlFecha & _
                    sqlComentario & _
                    " order by bm.`Fecha`, bm.idbancosmovimientos  "
                    
         Else
         
         
  vsql1 = " select  " + _
  " pp.nombre, " + _
  " bm.tipomovimiento, " + _
  " bm.idbancosmovimientos, " + _
  " bm.nrocomprobante, " + _
  " bm.codpersona, " + _
  " bm.`Fecha`, " + _
  " bm.`idBancos`, " + _
  " bb.descripcion, " + _
  " bm.`idBancosCuentas`, " + _
  " bm.`Debito`, " + _
  " bm.`Credito` , " + _
  " bm.saldo, " + _
  " bm.`NroInterno`, " + _
  " bm.`TipoMovimiento`, " + _
  " bm.`NroCheque`, " + _
  " bm.comentario , " + _
  " bm.conciliado "

vsql2 = "  from bancosmovimientos bm " + _
  "  left join bancoscuentas b on " + _
   " bm.idBancosCuentas = b.idBancosCuentas " + _
  "  left join bancos bb on " + _
  "   bm.`idBancos` = bb.`idBancos` " + _
  " left join proveedores pp on " + _
   "  pp.`codigo` = bm.`codPersona` " + _
  " Inner Join " + _
  " ( select nrointerno from `asientos`  " + _
   "  inner join `asientosdetalle` on " + _
   "    (`asientos`.`Numero` = `asientosdetalle`.`Numero`) " + _
  " Where " + _
   "  asientosdetalle.codigocuenta = '" + Me.vctasCodigo.Text + "' " + _
  " ) aa " + _
  "   on bm.NroInterno = aa.NroInterno " + _
  "   where " + vsqlFecha + _
  "   group by idbancosMovimientos"
         
  fBancoCaja = vsql1 + vsql2
         
         End If
         

        Case "Personas"
       
        fBancoCaja = "select  pp.nombre, bm.TipoMovimiento,bm.idBancosMovimientos, bm.nrocomprobante, bm.codpersona, bm.`Fecha`,   bm.`idBancos`,   bb.descripcion,   bm.`idBancosCuentas`," & _
         "  sum(bm.`Debito`) as Debito,  sum(bm.`Credito`) as Credito ,   bm.saldo, bm.`NroInterno`,  bm.`TipoMovimiento`, bm.`NroCheque`, bm.comentario , bm.conciliado from ba" & _
         "ncosmovimientos  bm   left join bancoscuentas b on      bm.idbancoscuentas = b.i" & _
         "dbancoscuentas   inner join bancos bb on     bm.`idBancos` = bb.`idBancos`      " & _
         "inner join proveedores pp on  pp.`codigo` = bm.`codPersona`      " & _
         " where  " + _
        vsqlFecha + " group by codPersona " & _
         " order by pp.`nombre`"

              
       End Select
        
End If

End Function


Public Sub cmdFiltrar_Click()
    On Error Resume Next

    Dim vgrupo As String
    Dim rsBancoCajaDetalle As New ADODB.Recordset, sqlBancoCajaDetalle As String
    Dim vsaldo As Double
    Dim vsqlFecha As String
    
    Dim vcondi2 As String
    
    
    Me.tabbc.SelectedItem = 1
    
    vaccion = ""
    
    If Me.vtipovalor = "VALE" Then vaccion = "Consulta.Vales"
    
    If Not Me.vcbancocaja = "" Then vcondi2 = " and (idbancos = '" + Str(vcbancocaja) + "')"
    
    
    
    
    If Me.chkFecha Then
        vsqlFecha = " "
    Else
        vsqlFecha = " WHERE (Fecha >= '" & strfechaMySQL(dtpFecha(0).Value) & "' AND Fecha <= '" & strfechaMySQL(dtpFecha(1).Value) & "')"
        
    End If
    
    If vvviene = "movidiario" Then
      vsqlFecha = " WHERE (idbancosmovimientos >= " & Str(vidd) & " AND idbancosmovimientos <= " & Str(vidH) & ")"
    End If
    
        
    sqlBancoCajaDetalle = "SELECT Bancosmovimientos.*, BancosCuentas.Cuenta FROM Bancosmovimientos LEFT JOIN BancosCuentas ON Bancosmovimientos.idBancosCuentas=BancosCuentas.idBancosCuentas " + _
    vsqlFecha
    
     
     If Not (vcliprovee = "") Or Not (Me.vcbancocaja.Text = "") Then vFiltro2 = componerFiltro2()
     
    'If Not Me.vctasCodigo.Text = "" Then vFiltro2 = componerFiltro2()
    
       
     sqlBancoCajaDetalle = fBancoCaja(Me.dtpFecha(0).Value, Me.dtpFecha(1).Value)
          
    If Not Me.vnrointerno = "" Then sqlBancoCajaDetalle = "SELECT Bancosmovimientos.*, BancosCuentas.Cuenta FROM Bancosmovimientos LEFT JOIN BancosCuentas ON Bancosmovimientos.idBancosCuentas=BancosCuentas.idBancosCuentas WHERE Bancosmovimientos.nrointerno = " + Str(Me.vnrointerno.Text)
    
    If Val(Me.vnrocheque.Text) > 0 Then sqlBancoCajaDetalle = filtraNroCheque(vnrocheque.Text)
    
        
 
    Select Case Me.vtipolistado
    
    
        Case "Movimientos Caja - Banco con movimientos Contables"
    
            
            Me.tabbc.TabIndex = 1
            
            Wait (1000)
            
            prenderCartel
            
            
            Call Reporte_MoviCBA(dtpFecha(0).Value, dtpFecha(1).Value)
            
            apagarCartel
            
            Exit Sub
    
            
        Case "Saldos"
        
            Call generarSaldosTemp(Me.dtpFecha(0).Value, Me.dtpFecha(1).Value, "todos")
            
            Unload Mantenimiento
            Load Mantenimiento
             
            MsgBox "Prepare la impresora ...", vbInformation, "Mensaje ..."
            
            dtpFecha(0).SetFocus
            
            Unload Mantenimiento
            Load Mantenimiento
                 
           
            drBancosSaldo.Sections("TituloEmpresa").Controls("efiltro").Caption = "Movimientos con fecha desde: " + Str(dtpFecha(0).Value) + " hasta: " + Str(dtpFecha(1).Value)
            drBancosSaldo.Show
            
         
            Exit Sub
        
            
        Case "Agrupado por Cajas-Bancos"
        
            Call generarSaldosTemp(Me.dtpFecha(0).Value, Me.dtpFecha(1).Value, "")
            
            Unload Mantenimiento
            Load Mantenimiento
             
            MsgBox "Prepare la impresora ...", vbInformation, "Mensaje ..."
            
            dtpFecha(0).SetFocus
            
            Unload Mantenimiento
            Load Mantenimiento
                
           
            drBancosSaldo.Sections("TituloEmpresa").Controls("efiltro").Caption = "Movimientos con fecha desde: " + Str(dtpFecha(0).Value) + " hasta: " + Str(dtpFecha(1).Value)
            drBancosSaldo.Show
            
         
            Exit Sub
    
    Case "Con detalles de Movimientos contables"
    
        Call cmdImprimir_Click(1)
    
        Exit Sub
    
    Case "Personas"
    
        
    End Select
 
    Me.KlexDetalle.SetFocus
 
    LlenarGrilla (sqlBancoCajaDetalle)
End Sub

Function componerFiltro2() As String
Dim vsql As String

vsql = ""

If Not Trim(Me.vcbancocaja.Text) = "" Then vsql = vsql + " and  (bm.idBancos = '" + Me.vcbancocaja.Text + "') "


If Not Trim(Me.vcliprovee.Tag) = "" Then vsql = vsql + " and  (bm.codpersona = '" + Me.vcliprovee.Tag + "') "


If Not Trim(Me.vctipovalor.Text) = "" Then vsql = vsql + " and  (bm.idTipoValor = '" + Me.vctipovalor.Text + "') "

If Not Trim(Me.vctm.Text) = "" Then
    vsql = vsql + " and  (bm.TipoMovimiento = '" + Me.vctm.Text + "') "
    vaccion = "Consulta.Vales"
End If


If Not Me.vcomentario.Text = "" Then vsql = vsql + "  and  ( (bm.comentario like '%" + Me.vcomentario.Text + "%') or (bm.comentario2 like '%" + Me.vcomentario.Text + "%'))"


If Not Me.vnrocomprobante.Text = "" Then vsql = vsql + " and bm.nrocomprobante = " + Trim(Me.vnrocomprobante.Text)

If Not Me.vctasCodigo.Text = "" Then vsql = vsql + addCondiCtas(Me.vctasCodigo)

If Me.RadCancelados Then vsql = vsql + " and conciliado = '-'"

If Me.RadPendientes Then vsql = vsql + " and (conciliado is null  or  not conciliado = '-')"

If vGidhastabm > 0 Then
    vsql = vsql + " and idbancosmovimientos > " + Str(vGiddesdebm) + " and idbancosmovimientos <= " + Str(vGidhastabm)
    vsqlFecha = " "
End If



componerFiltro2 = vsql


' ------------ limpiar campos --------

vcliprovee.Tag = ""
Me.vdbancocaja.Tag = ""
Me.vcbancocaja.Tag = ""
Me.vctipovalor.Tag = ""


Me.vcliprovee.Tag = ""
Me.vtipovalor.Tag = ""

Me.vdbancocaja.Text = ""
Me.vcbancocaja.Text = ""
Me.vcliprovee.Text = ""
Me.vtipovalor.Text = ""
Me.vctipovalor.Text = ""

vdtm.Tag = ""
vdtm.Text = ""
vctm.Text = ""

Me.vcomentario.Text = ""
Me.vctasCodigo.Text = ""
Me.vCtasDescrip.Tag = ""
Me.vCtasDescrip.Text = ""


'-------------------------------------

End Function

Function addCondiCtas(vcod As String) As String
Dim vsql As String

addCondiCtas = " and bm.NroInterno in (SELECT asientos.NroInterno From  `asientos` INNER JOIN `asientosdetalle` ON (`asientos`.`Numero` = `asientosdetalle`.`Numero`) " + _
" where asientosdetalle.CodigoCuenta = '" + vcod + "')"

End Function


Function filtraNroCheque(vnrocheque As String) As String
On Error Resume Next

 filtraNroCheque = "select    cheques.Ncheque,   bancosmovimientos.*,    bancoscuentas.cuenta  from " & _
 "bancosmovimientos    left join bancoscuentas on      bancosmovimientos.idbancosc" & _
 "uentas = bancoscuentas.idbancoscuentas    inner join cheques on     bancosmovimie" & _
 "ntos.idcheques = cheques.idCheques  where cheques.Ncheque = " + vnrocheque

If Err Then Exit Function
End Function


Function SaldoAnterior(vfd As Date, Optional ByVal vcondi2 As String) As Double
Dim vsql As String

'If vcondi2 = "" Then vcondi = " 1=1 "

vsql = "select    (sum(bm.`Debito`) - sum(bm.`Credito`)) as saldo  from bancosmovimiento" & _
"s bm where Fecha < '" & strfechaMySQL(vfd) + "'" + vcondi2

SaldoAnterior = Val(EsNulo(traerDatos2(vsql, "saldo", pathDBMySQL)))

End Function
Private Sub LlenarGrilla(vsqlBancoCajaDetalle As String)
On Error Resume Next
Dim vsaldo As Double
Dim vcampos, vvalores, vsql, vvcp As String
Dim rsBancoCajaDetalle2 As New ADODB.Recordset


vsql = "delete from movibc"
Call EjecutarScript(vsql, PathDBListados)  ' vacio la tabla
 
vcampos = "fecha,banco,cuenta,debito,credito,saldo,nrocheque,comentarios,cp"


With rsBancoCajaDetalle2
        Dim i As Integer
        
        .CursorLocation = adUseClient
        
        Call .Open(vsqlBancoCajaDetalle, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .RecordCount > 0 Then
            MsgBox "No hay datos para mostar", vbInformation
            Exit Sub
        End If

        
        
        FormatoGrilla (.RecordCount)
        
        
        i = 1
        
        
         If Not Me.chkFecha.Value = xtpChecked Then
            KlexDetalle.TextMatrix(i, 9) = SaldoAnterior(Me.dtpFecha(0).Value, vFiltro2)
        Else
            KlexDetalle.TextMatrix(i, 9) = 0
        End If
        
        vsaldo = KlexDetalle.TextMatrix(i, 9)
        
        barra.Value = 0
        
        barra.Max = .RecordCount
    
        Do Until .EOF = True
        
            vvalores = ""
            
            barra.Value = barra.Value + 1
            
            i = i + 1
            KlexDetalle.Rows = i + 1
            
            KlexDetalle.TextMatrix(i, 0) = EsNulo(.Fields("tipomovimiento").Value)
            KlexDetalle.TextMatrix(i, 1) = EsNulo(.Fields("idBancosMovimientos").Value)
            KlexDetalle.TextMatrix(i, 2) = EsNulo(.Fields("Fecha").Value)
            
            vvalores = vvalores + "'" + strfecha2(EsNulo(.Fields("Fecha").Value)) + "',"
            
            KlexDetalle.TextMatrix(i, 3) = EsNulo(.Fields("nrocheque").Value)
            KlexDetalle.TextMatrix(i, 4) = EsNulo(.Fields("nrocomprobante").Value)
            
            KlexDetalle.TextMatrix(i, 5) = "[" & EsNulo(.Fields("descripcion").Value) & "]"
            
            vvalores = vvalores + "'" + EsNulo(.Fields("descripcion").Value) + "',"
            
            vvalores = vvalores + "'" + EsNulo(.Fields("idBancosCuentas").Value) + "',"
            
            KlexDetalle.TextMatrix(i, 6) = EsNulo(.Fields("Cuenta").Value)
            
            vvalores = vvalores + "'" + EsNulo(.Fields("Cuenta").Value) + "',"
            
            
            KlexDetalle.TextMatrix(i, 7) = EsNulo(.Fields("Debito").Value)
            
            vvalores = vvalores + EsNulo(.Fields("Debito").Value) + ","
             
            
            KlexDetalle.TextMatrix(i, 8) = EsNulo(.Fields("Credito").Value)
            
            vvalores = vvalores + EsNulo(.Fields("Credito").Value) + ","

            
            vsaldo = vsaldo + EsNulo(.Fields("Debito").Value) - EsNulo(.Fields("Credito").Value)
            
            KlexDetalle.TextMatrix(i, 9) = vsaldo
            
            KlexDetalle.TextMatrix(i, 10) = EsNulo(.Fields("conciliado").Value)
            
            
            KlexDetalle.TextMatrix(i, 11) = EsNulo(.Fields("Comentario").Value)
            
            vvalores = vvalores + EsNulo(vsaldo) + ","
            
            vvalores = vvalores + "'" + EsNulo(.Fields("nrocheque").Value) + "',"
            
            vvalores = vvalores + "'" + EsNulo(.Fields("Comentario").Value) + "'"
            
            
            
            '--------------
            vvcp = ""
            
           ' vsql = "select nombre from cuentascorrientes where nrointerno=" + EsNulo(.Fields("NroInterno").Value)
             
             vsql = "select nombre from proveedores where codigo ='" + EsNulo(.Fields("codpersona").Value) + "'"


            vvcp = traerDatos2(vsql, "nombre", pathDBMySQL)
            
            vsql = "select nombre from pcuentascorrientes where nrointerno=" + EsNulo(.Fields("NroInterno").Value)

            vvcp = vvcp + traerDatos2(vsql, "nombre", pathDBMySQL)
            
            '----------------
            
            vvalores = vvalores + ",'" + vvcp + "'"
            
            
            
            .MoveNext
        
            Call GuardarTemp(vcampos, vvalores)
        
        
        Loop
    
        
    End With

    KlexDetalle.TopRow = KlexDetalle.Rows - 1
    
    vsqlBancoCajaDetalle = ""

    If rsBancoCajaDetalle2.State = 1 Then
        rsBancoCajaDetalle2.Close
        Set rsBancoCajaDetalle2 = Nothing
    End If
    
    Me.tabbc.SelectedItem = 1
    
If Err Then GrabarLog "", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub GuardarTemp(ByVal vcampos As String, ByVal vvalores As String)
On Error Resume Next
Dim vsql As String

vsql = "insert into MoviBC (" + vcampos + ") values (" + vvalores + ")"
Call EjecutarScript(vsql, PathDBListados)
            
If Err Then Exit Sub
End Sub


Function fgrupo() As String
Dim vsql As String


'sqlBancoCajaDetalle = "SELECT Bancosmovimientos.*, BancosCuentas.Cuenta FROM Bancosmovimientos LEFT JOIN BancosCuentas ON Bancosmovimientos.idBancosCuentas=BancosCuentas.idBancosCuentas WHERE (Fecha >= '" & strfechaMySQL(dtpFecha(0).Value) & "' AND Fecha <= '" & strfechaMySQL(dtpFecha(1).Value) & "')"


End Function

Private Sub cmdReimprimir_Click()

On Error Resume Next

  Unload Mantenimiento
  Load Mantenimiento
    
        
    drBCMTotal.Sections("TituloEmpresa").Controls("lblFechaDesde").Caption = frmBancoCajaDetalle.dtpFecha(0).Value
    drBCMTotal.Sections("TituloEmpresa").Controls("lblFechaHasta").Caption = frmBancoCajaDetalle.dtpFecha(1).Value
        
        
    'If Not vsolofaltantes Then
        drBCMTotal.Sections("PieInforme").Controls("totaldebe").Caption = Format(vtGralD, "###,###,##0.00")
        drBCMTotal.Sections("PieInforme").Controls("totalhaber").Caption = Format(vtGralC, "###,###,##0.00")
    'End If
    
    drBCMTotal.Show

If Err Then Exit Sub

End Sub

Private Sub dtpFecha_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
    
    If KeyAscii = 13 Then
        Select Case Index
    
            Case 0
                dtpFecha(Index + 1).SetFocus
        
            Case 1
                cmdFiltrar.SetFocus
    
        End Select
    End If
If Err Then GrabarLog "dtpFecha_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub Form_Load()
On Error Resume Next

    dtpFecha(0).Value = Date - 30
    dtpFecha(1).Value = Date

    With Me
        .Show
    End With
    
    Me.tabbc.SelectedItem = 0
    
    FormatoGrilla (1)
    
    init
    
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub setcierremensual()
On Error Resume Next
Dim vValor As Long
Dim vsqllocal As String

' verifico que no esté cerrando un mes cerrado

vsqllocal = "select max(idanomes) as c from t_cajacierre"
vValor = Val(traerDatos2(vsqllocal, "c", pathDBMySQL))

vidanomesAnterior = vValor

'If vidanomes <= vValor Then
'    MsgBox "Hay un cierre de mes posterior. " + Chr(13) + _
    "Último mes cerrado: " + Format(vValor, "yyy/mm")
'End If


If Err < 0 Then
    vidanomesAnterior = 0
End If

End Sub

Private Sub init()
vidanomesAnterior = 0
If Not vidanomes > 0 Then vidanomes = 0

tsd = 0
tsc = 0

cierremensual = False


Call setcierremensual

Me.paraBalance.Visible = False

vgid = 0

configurarGrid

unicavez = 0

Me.fcancelaValores.Visible = False

Me.vtipolistado.Text = "Todos"
Me.vtipolistado.AddItem "Todos", 0
Me.vtipolistado.AddItem "Agrupado por Cajas-Bancos", 1
Me.vtipolistado.AddItem "Agrupado por Fecha", 2
Me.vtipolistado.AddItem "Agrupado por Cajas", 3
Me.vtipolistado.AddItem "Agrupado por Bancos", 4
Me.vtipolistado.AddItem "Con detalles de Movimientos contables", 5
Me.vtipolistado.AddItem "Personas", 6
Me.vtipolistado.AddItem "Movimientos Caja - Banco con movimientos Contables", 7
Me.vtipolistado.AddItem "Saldos", 8


Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2 - 1000

impDirecto = False

Me.dtpFecha(0).Value = Date
Me.dtpFecha(1).Value = Date

vfecha.Value = Date


Call Buscar("", "")

formatearGrilla2


End Sub

Private Sub FormatoGrilla(vCantidadRenglones As Integer)
On Error Resume Next

    Dim i As Integer

    With KlexDetalle
        .FixedRows = 1
        .FixedCols = 1
    
        .Cols = 12
        .Rows = vCantidadRenglones + 1
        
        If vCantidadRenglones = 1 Then
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
                .ColWidth(i) = 0
            Next
        End If
        
        .TextMatrix(0, 0) = "TM"
        .ColWidth(0) = 600
        
        .TextMatrix(0, 1) = "idBancosMovimientos"
        .ColWidth(1) = 0
               
        .TextMatrix(0, 2) = "Fecha"
        .ColWidth(2) = 1000
        
        .TextMatrix(0, 3) = "N.Cheque"
        .ColWidth(3) = 1000
        
        .TextMatrix(0, 4) = "Nro.Comp"
        .ColWidth(4) = 1000
        
        .TextMatrix(0, 5) = "Banco/Caja"
        .ColWidth(5) = 1500
        
        .TextMatrix(0, 6) = "Cuenta Banco"
        .ColWidth(6) = 100
        
        .TextMatrix(0, 7) = "Ingreso"
        .ColWidth(7) = 1500
        .ColDisplayFormat(7) = "###,###,##0.000"
        .ColAlignment(7) = 6



        .TextMatrix(0, 8) = "Egresos"
        .ColWidth(8) = 1500
        .ColDisplayFormat(8) = "###,###,##0.000"
        .ColAlignment(8) = 6
        
        
        
        .TextMatrix(0, 9) = "Saldo"
        .ColWidth(9) = 1500
        .ColDisplayFormat(9) = "###,###,##0.000"
        .ColAlignment(9) = 6
        
        .TextMatrix(0, 10) = "Estado"
        .ColWidth(10) = 1000
  
        .TextMatrix(0, 11) = "Obs."
        .ColWidth(11) = 5000
  

        .BackColorAlternate = 14737632
    End With
    
If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub TabStrip1_Click(Index As Integer)

End Sub

Private Sub PicInferior_Click()

End Sub

Private Sub TabControl1_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

End Sub

Private Sub grilla_Click()
On Error Resume Next

vr = grilla.Row
vgid = grilla.TextMatrix(vr, 1)
vidcierecaja = grilla.TextMatrix(vr, 3)

vidd = traerDatos2("select iddesde from t_cajacierre where idcajacierre = " + Str(vidcierecaja), "iddesde", pathDBMySQL)

vidH = traerDatos2("select idhasta from t_cajacierre where idcajacierre = " + Str(vidcierecaja), "idhasta", pathDBMySQL)


If Val(grilla.TextMatrix(vr, 8)) > 0 Then
    vidAdesde = Val(traerDatos2("select idasientosdesde from t_cajacierre where idcajacierre = " + Str(vidcierecaja), "idasientosdesde", pathDBMySQL))
    vidAhasta = Val(traerDatos2("select idasientoshasta from t_cajacierre where idcajacierre = " + Str(vidcierecaja), "idasientoshasta", pathDBMySQL))
Else
    vidAdesde = 0
    vidAhasta = 0
End If

If Err Then Exit Sub
End Sub

Private Sub grilla_DblClick()
Call Buscar("", " where t_cajacierresaldos.idcajacierre = " + Str(vidcierecaja))

MsgBox "El cierre de caja fue seleccionado"

End Sub

Private Sub pbCarga_Click(Index As Integer)

End Sub

Private Sub KlexDetalle_Click()

' Call VerDetalle(Me.KlexDetalle.TextMatrix(vr, 1), Me.gdetalle)
    
End Sub

Private Sub KlexDetalle_DblClick()
On Error Resume Next
Dim r As Integer
Dim vsql, valor, vtipo As String
Dim vid As Long


r = Me.KlexDetalle.Row

valor = Me.KlexDetalle.TextMatrix(r, 10)

vid = Me.KlexDetalle.TextMatrix(r, 1)

vtipo = Me.KlexDetalle.TextMatrix(r, 0)


'If Me.KlexDetalle.TextMatrix(r, 10) = "-" Then
'
'        If Not MsgBox("Este VALE ya fue cancela. Quiere seleccionarlo de todas manera ?", vbYesNo) = vbYes Then
'            Exit Sub
'        End If
'End If

If vtipo = "VL" Then
    If valor = "Cancelado" Then
    
        Me.KlexDetalle.TextMatrix(r, 10) = "-"
        vsql = "update bancosmovimientos set Conciliado='-' where idBancosMovimientos=" + Str(vid)
         Me.KlexDetalle.Col = 10
         Me.KlexDetalle.CellForeColor = vbBlack
         Me.KlexDetalle.CellBackColor = vbBlue
         
         'Me.KlexDetalle.ForeColorSel = vbBlack
    End If
    
    If valor = "" Then
        vsql = "update bancosmovimientos set Conciliado='Cancelado' where idBancosMovimientos=" + Str(vid)
        Me.KlexDetalle.TextMatrix(r, 10) = "Cancelado"
            
          Me.KlexDetalle.Col = 10
         
         Me.KlexDetalle.CellForeColor = vbBlack
         Me.KlexDetalle.CellBackColor = vbRed
        
    End If
    
    
    If valor = "-" Then
        vsql = "update bancosmovimientos set Conciliado='' where idBancosMovimientos=" + Str(vid)
        Me.KlexDetalle.TextMatrix(r, 10) = ""
            
          Me.KlexDetalle.Col = 10
          
         Me.KlexDetalle.CellForeColor = vbBlack
         Me.KlexDetalle.CellBackColor = vbWhite
         
    End If
    
    
End If



If vtipo = "CH" Then
    If valor = "Conciliado" Then
        Me.KlexDetalle.TextMatrix(r, 10) = ""
        vsql = "update bancosmovimientos set Conciliado='' where ibancosmovimientos=" + Str(vid)
        Me.KlexDetalle.BackColorSel = vbWhite
        Me.KlexDetalle.ForeColorSel = vbBlack
        
    Else
        vsql = "update bancosmovimientos set Conciliado='Cancelado' where ibancosmovimientos=" + Str(vid)
        Me.KlexDetalle.TextMatrix(r, 10) = "Cancelado"
        Me.KlexDetalle.BackColorSel = vbRed
        Me.KlexDetalle.ForeColorSel = vbBlack
    End If
End If



If Not vsql = "" Then Call EjecutarScript(vsql, pathDBMySQL)


If Err Then Exit Sub
End Sub

Private Sub KlexDetalle_SelChange()
On Error Resume Next
Dim vsql As String
 
vr = Me.KlexDetalle.Row
 


vsql = " SELECT " + _
"  `bancosmovimientos`.`Fecha`, " + _
"  `asientosdetalle`.`CodigoCuenta`, " + _
" cuentas.cuenta, " + _
"  `asientosdetalle`.`Debe`, " + _
"  `asientosdetalle`.`Haber`, `bancosmovimientos`.`NroCheque`, " + _
"  `bancosmovimientos`.`codpersona` as CPersona, proveedores.nombre as Persona, " + _
"  `proveedores`.`Nombre`, bancosmovimientos.comentario2, bancosmovimientos.comentario  " + _
" FROM " + _
"  `bancosmovimientos` " + _
"  left JOIN `asientos` ON (`asientos`.`NroInterno` = `bancosmovimientos`.`NroInterno`) " + _
"  left JOIN `asientosdetalle` ON (`asientos`.`Numero` = `asientosdetalle`.`Numero`) " + _
"  left JOIN `cuentas` ON (`asientosdetalle`.`CodigoCuenta` = `cuentas`.`CodigoCuenta`) " + _
"  left JOIN `proveedores` ON (`bancosmovimientos`.`ClienteProveedor` = `proveedores`.`Codigo`) " + _
" where bancosmovimientos.idbancosMovimientos = " + Str(KlexDetalle.TextMatrix(vr, 1))


    
    Call LlenarGrilla2(Me.gdetalle, vsql, 7, pathDBMySQL)

If Err Then Exit Sub
End Sub

Private Sub pbCarga11_Click()
Call fbuscarGrilla("bancos", "Descripcion", "idBancos", Me.vdcaja.Name, Me)
End Sub

Private Sub PusAyer_Click()
dtpFecha(0).Value = Date - 1
dtpFecha(1).Value = Date - 1
End Sub

Private Sub PusAyuda_Click()
Dim x
'X = Shell(vbpath + "ayuda.bat")
'donde mipagina.cl colocas la url que quieras
End Sub

Private Sub PusCerrarCaja_Click()
Dim vsql, vidS As String
Dim vid, viddesde, vidhasta, vidAsientos As Long

vsql = "select * from t_cajacierre  where fecha = '" + strfechaMySQL(vfecha) + "'"
vidS = traerDatos2(vsql, "idcajacierre", pathDBMySQL)

vid = Val(vidS)

If Val(vid) > 0 And Not vidanomes > 0 Then
   If MsgBox("Cuidado, está cerrando en el mismo día de otro cierre. " + Chr(13) + "Continúa de todas maneras ? ", vbYesNo) = vbNo Then
        Exit Sub
   End If
End If
    
    If vidanomes > 0 Then
        ' en el cao que sea un cierre mensual busco el último cierre mensual
        vsql = "select max(idhasta) as c  from t_cajacierre where idanomes > 0 "
    Else
        ' cierre diario
        vsql = "select max(idhasta) as c   from t_cajacierre where (not idanomes > 0 or idanomes is null)"
    End If
    
   ' vsql = "select idhasta as c  from t_cajacierre order by idhasta desc limit 1 "
    viddesde = Val(traerDatos2(vsql, "c", pathDBMySQL))  ' paso
    
    vsql = "select idBancosMovimientos as c  from bancosmovimientos order by idBancosMovimientos desc limit 1 "
    vidhasta = Val(traerDatos2(vsql, "c", pathDBMySQL)) ' paso
     
    If vidanomes > 0 Then
        vsql = "select max(idasientoshasta) as c  from t_cajacierre where idanomes > 0 "
    Else
        vsql = "select max(idasientoshasta) as c  from t_cajacierre where (not idanomes > 0 or idanomes is null)"
    End If

    vidasientosdesde = Val(traerDatos2(vsql, "c", pathDBMySQL)) ' paso
    
      If vidasientosdesde = 0 Then
            vsql = "select max(idasientos) as c  from asientos where fecha < '" + strfechaMySQL(Me.vfecha) + "'"
            vidasientosdesde = Val(traerDatos2(vsql, "c", pathDBMySQL))
      End If
    
    vsql = "select max(idasientos) as c  from asientos "
    vidasientoshasta = Val(traerDatos2(vsql, "c", pathDBMySQL))
    
    vsd = saldoIngresos(viddesde)  ' paso
    vsc = saldoEgresos(viddesde)   ' paso
    
    tsd = tsd + vsd
    tsc = tsc + vsc
    
    ' ****************************** validar ***********************************************************
    If valComparaSalCajaAsiento(vidasientosdesde, vidasientoshasta, viddesde, vidhasta, vidanomes) Then
        Exit Sub
    End If
    ' **************************************************************************************************
    
    If Not validarCerrarCaja(viddesde, vidhasta) Then Exit Sub
    
   If Not verificocierredemes(vidanomes) Then Exit Sub
   
  
    ' insert
    vsql = "insert into t_cajacierre (fecha,estado, iddesde, idhasta, idasientosdesde,idasientoshasta,idanomes) values ('" + strfechaMySQL(vfecha) + "','" + vestado.Text + "'," + Str(viddesde) + "," + Str(vidhasta) + "," + Str(vidasientosdesde) + "," + Str(vidasientoshasta) + "," + Trim(vidanomes) + ")"
    Call EjecutarScript(vsql, pathDBMySQL)
    
   
    'If vidanomes > 0 Then
    '    vsql = "select idcajacierre as c  from t_cajacierre where idanomes > 0  order by idcajacierre desc limit 1 "
   ' Else
   '     vsql = "select idcajacierre as c  from t_cajacierre where (not idanomes > 0 or idanomes is null) order by idcajacierre  desc limit 1 "
   ' End If
   
   vsql = "select max(idcajacierre) as c  from t_cajacierre "
    vid = Val(traerDatos2(vsql, "c", pathDBMySQL))
   
    'vsql = "select idcajacierre as c  from t_cajacierre order by idcajacierre desc limit 1 "
   

    generarSaldosCajaCierre (vid) ' paso 3
    
vgid = 0

Call Buscar("", "")

vgid = grilla.TextMatrix(1, 1)
vidcierecaja = grilla.TextMatrix(1, 3)



MsgBox "El cierre de caja se realizó correctamente. ", vbInformation

End Sub

Function verificocierredemes(ByRef vidanomes) As Boolean
On Error Resume Next
Dim vsqllocal As String
Dim vValor As Long
verificocierredemes = True
vsqllocal = "select idanomes as c from t_cajacierre where idanomes =" + Str(vidanomes)
vValor = Val(traerDatos2(vsqllocal, "c", pathDBMySQL))

If Not vValor = 0 Then
    If MsgBox("El periódos " + Format(vidanomes, "0000/00") + _
    " está cerrado." + Chr(13) + "Quiere continuar de todas maneras ? ", vbYesNo) = vbYes Then
    
    Else
        verificocierredemes = False
        vidanomes = 0
    End If
    
   
    
End If
If Err Then Exit Function
End Function

Function validarCerrarCaja(ByVal vid As Long, ByVal vih As Long) As Boolean
Dim vmensaje As String

vmensaje = ""

validarCerrarCaja = True


If vid = vih Then
    vmensaje = vmensaje + " No hay movimiento nuevo desde el último cierre" + Chr(13)
    vidanomes = 0
    validarCerrarCaja = False
End If


If Not vmensaje = "" Then
    MsgBox vmensaje, vbCritical
End If
End Function


Private Sub generarSaldosTemp(vfd, vfh As Date, Optional vtodos As String)
Dim vsa, vsp, vs, vd, vh, vss As Double
Dim vidBancos As String
Dim vsql, vcampos, vvalores, vnombre  As String

Dim rs As New ADODB.Recordset


'------------ borrar la base temporar saldos -----
vsql = "delete from saldos "
Call EjecutarScript(vsql, PathDBListados)
'------------------------------------------------

If vtodos = "todos" Then
    vsql = "select * from bancos where  not idbancos = '098' and not EsCaja = 'B' order by idBancos"
Else
    vsql = "select * from bancos where tipodisponibilidad = 'Disponible' and not EsCaja = 'B' order by idBancos"
End If

Call rs.Open(vsql, ConnDDBB, adOpenStatic, adLockPessimistic)

barra.Value = 0

barra.Max = rs.RecordCount

Do Until rs.EOF
    
    barra.Value = barra.Value + 1
    
    vidBancos = rs.Fields("idbancos")
    vnombre = rs.Fields("descripcion")
    
    vsa = sacierreTemp(vidBancos, vfd)
    
    
  '  If vidBancos = "1010" Then MsgBox ""
    
    Call spcierreTemp(vidBancos, vfd, vfh, vd, vh, vsp)
    
    
    vss = vsa + vsp


    vcampos = "(codigo,nombre,sa,d,c,sp,saldo)"
    vvalores = "('" + (vidBancos) + "','" + (vnombre) + "'," + Str(vsa) + "," + Str(vd) + "," + Str(vh) + "," + Str(vsp) + "," + Str(vss) + ")"

    vsql = "insert into saldos " + vcampos + " values " + vvalores

    If Abs(vss) > 0 Then
            Call EjecutarScript(vsql, PathDBListados)
    End If
    
    rs.MoveNext

Loop

End Sub


Private Sub generarSaldosCajaCierre(vid As Long)
Dim vsa, vsp, vs, vsd, vsc, vsacumulado As Double
Dim vidBancos As String
Dim vsql, vcampos, vvalores As String

Dim rs As New ADODB.Recordset


' vid es ùltimo id de la tabla t_cajacierre que se llama idcajacierre

vsql = "select * from bancos where not EsCaja = 'B'"

Call rs.Open(vsql, ConnDDBB, adOpenStatic, adLockPessimistic)

Do Until rs.EOF

    
    vidBancos = rs.Fields("idbancos")
    
    Dim vingresos As Double
    Dim vegresos As Double
    
    
    If Not vidBancos = "098" Then
    
               ' If vidBancos = "1010" Or vidBancos = "1000" Then MsgBox ""
                
                vsa = sacierre(vidBancos, vid)
                
                ' vsp = spcierre(vidBancos, vid) ' ahora lo voy a calcular
                
                vs = sAcumuladoCierre(vidBancos, vid)
              
                 vsp = vs - vsa
                    
                'vs = vsa + vsp
                
                vingresos = fun_totalDC(vidBancos, "Debito", vid)
                vegresos = fun_totalDC(vidBancos, "Credito", vid)
                Debug.Print (" Suma debito / credito : " + Str(vingresos) + " - " + Str(vegresos) + " >> " + Str(vidBancos))
                
            
                vcampos = "(santerior,speriodo,saldo,idbancos,idcajacierre,idanomes,ingresos,egresos)"
                vvalores = "(" + Str(vsa) + "," + Str(vsp) + "," + Str(vs) + ",'" + vidBancos + "'," + Str(vid) + "," + Str(vidanomes) + "," + Str$(vingresos) + "," + Str$(vegresos) + ")"
            
                vsql = "insert into t_cajacierresaldos " + vcampos + " values " + vvalores
            
                Debug.Print "Caja Saldos: -> " + vsql
            
            
                If Abs(vsa) > 0 Or Abs(vsp) > 0 Then
                        Call EjecutarScript(vsql, pathDBMySQL)
                End If
   End If
                   
    rs.MoveNext

Loop




End Sub


Function sacierreTemp(ByVal vidBancos As String, ByVal fd As Date) As Double
' saldo anterior al cierre

On Error Resume Next

Dim vsql As String
Dim v As Double

vsql = " select " + _
" sum(bancosmovimientos.debito - bancosmovimientos.credito) As c" + _
" From bancosmovimientos " + _
" Where " + _
"  (bancosmovimientos.fecha < '" + strfechaMySQL(fd) + "' and bancosmovimientos.idBancos ='" + vidBancos + "') " + vFiltro2 + " " + _
" Group By " + _
"  bancosmovimientos.idbancos"

v = Val(traerDatos2(vsql, "c", pathDBMySQL))

sacierreTemp = v

If Err Then
    sacierreTemp = 0
    Exit Function
End If

End Function
Function sacierre(vidBancos As String, vid As Long) As Double
' saldo anterior al cierre

On Error Resume Next

Dim vsql As String
Dim v As Double


If vidanomes = 0 Then vidanomesAnterior = 0

If vidanomesAnterior > 0 Then  ' en el caso que sea un cierre de mes tiene que ir a buscar el saldo del último cierre mensual
        
        vsql = "select saldo as c  from t_cajacierresaldos where idbancos = '" + vidBancos + "' and idanomes= " + Str(vidanomesAnterior) + " order by idbcajacierresaldos desc limit 1"
        v = Val(traerDatos2(vsql, "c", pathDBMySQL))
Else
        vsql = "select saldo as c  from t_cajacierresaldos where idbancos = '" + vidBancos + "' and (idanomes=0 or idanomes is null) order by idbcajacierresaldos desc limit 1"
        v = Val(traerDatos2(vsql, "c", pathDBMySQL))
End If

sacierre = v



If Err Then
    sacierre = 0
    Exit Function
End If

End Function


Function fun_totalDC(idbanco, vDB, vid) As Double

Dim vsql, vsql2 As String
Dim vidhasta As Long

Dim vie, vegreso As Double

vsql = "Select iddesde as c from t_cajacierre where idcajacierre = " + Str(vid)

vidhasta = traerDatos2(vsql, "c", pathDBMySQL)

vsql2 = "select sum(" + vDB + ") as c  from bancosmovimientos where  idbancos = " + Str(idbanco) + " and  idBancosMovimientos > " + Str(vidhasta)

vie = Val(traerDatos2(vsql2, "c", pathDBMySQL))

fun_totalDC = vie

End Function




Private Sub spcierreTemp(ByVal vidBancos As String, ByVal vfd As Date, ByVal vfh As Date, ByRef vd, ByRef vh, ByRef vsp)
' saldo anterior al cierre
On Error Resume Next

Dim vsql As String


vsql = " select " + _
    " sum(bancosmovimientos.debito) as d, sum(bancosmovimientos.credito) as h, " + _
    " sum(bancosmovimientos.debito - bancosmovimientos.credito) As sp " + _
    " From bancosmovimientos " + _
    " Where bancosmovimientos.fecha >= '" + strfechaMySQL(vfd) + "' and bancosmovimientos.fecha <= '" + strfechaMySQL(vfh) + "' " + _
    " and bancosmovimientos.idbancos = '" + vidBancos + "'" + vFiltro2 + " " + _
    " Group By  bancosmovimientos.idBancos "

vd = Val(traerDatos2(vsql, "d", pathDBMySQL))
vh = Val(traerDatos2(vsql, "h", pathDBMySQL))
vsp = Val(traerDatos2(vsql, "sp", pathDBMySQL))


If Err Then
    'spcierreTemp = 0
    Exit Sub
End If

End Sub
Function saldoIngresos(ByVal viddesde As Long) As Double
' saldo anterior al cierre

On Error Resume Next

Dim vsql As String
Dim v As Double


vsql = " select " + _
    " sum(bancosmovimientos.debito) As c " + _
    " From bancosmovimientos " + _
    " Where bancosmovimientos.idBancosMovimientos > " + Str(viddesde) + _
    " Group By  bancosmovimientos.idBancos "

v = Val(traerDatos2(vsql, "c", pathDBMySQL))

saldoIngresos = v

If Err Then
    saldoIngresos = 0
    Exit Function
End If

End Function
Function saldoEgresos(ByVal viddesde As Long) As Double
' saldo anterior al cierre

On Error Resume Next

Dim vsql As String
Dim v As Double


vsql = " select " + _
    " sum(bancosmovimientos.credito) As c " + _
    " From bancosmovimientos " + _
    " Where bancosmovimientos.idBancosMovimientos > " + Str(viddesde) + _
    " Group By  bancosmovimientos.idBancos "

v = Val(traerDatos2(vsql, "c", pathDBMySQL))

saldoEgresos = v

If Err Then
    saldoEgresos = 0
    Exit Function
End If

End Function


Function sAcumuladoCierre(vidBancos As String, vid As Long) As Double
' saldo anterior al cierre
On Error Resume Next

Dim vsql As String
Dim v As Double
Dim viddesde As Long


vsql = "select sum(t.Debito) - sum(t.Credito) as saldo   " + _
" from bancosmovimientos t " + _
" where not t.idBancos='098' and " + _
" t.idbancos = '" + vidBancos + "'" + _
" group by t.idBancos "

sAcumuladoCierre = traerDatos2(vsql, "saldo", pathDBMySQL)


If Err Then
    sAcumuladoCierre = 0
    Exit Function
End If
End Function


Function spcierre(vidBancos As String, vid As Long) As Double
' saldo anterior al cierre

On Error Resume Next

Dim vsql As String
Dim v As Double
Dim viddesde As Long


vsql = "select iddesde as c from t_cajacierre where idcajacierre = " + Str(vid)
viddesde = traerDatos2(vsql, "c", pathDBMySQL)


vsql = " select " + _
    " sum(bancosmovimientos.debito - bancosmovimientos.credito) As c " + _
    " From bancosmovimientos " + _
    " Where bancosmovimientos.idBancosMovimientos > " + Str(viddesde) + " and bancosmovimientos.idbancos = " + vidBancos + _
    " Group By  bancosmovimientos.idBancos "

v = Val(traerDatos2(vsql, "c", pathDBMySQL))


vsql = " select " + _
    " sum(bancosmovimientos.credito) As c " + _
    " From bancosmovimientos " + _
    " Where bancosmovimientos.idBancosMovimientos > " + Str(viddesde) + " and bancosmovimientos.idbancos = " + vidBancos + _
    " Group By  bancosmovimientos.idBancos "

 tsc = tsc + Val(traerDatos2(vsql, "c", pathDBMySQL))

vsql = " select " + _
    " sum(bancosmovimientos.debito) As c " + _
    " From bancosmovimientos " + _
    " Where bancosmovimientos.idBancosMovimientos > " + Str(viddesde) + " and bancosmovimientos.idbancos = " + vidBancos + _
    " Group By  bancosmovimientos.idBancos "

tsd = tsd + Val(traerDatos2(vsql, "c", pathDBMySQL))


spcierre = v

If Err Then
    spcierre = 0
    Exit Function
End If

End Function
Function validarCierreMensual() As Boolean
Dim vsql As String
Dim vValor As Long

validarCierreMensual = True


vsql = "select max(idanomes) as c  from t_cajacierre where idanomes > " + Str(vidanomes)
'vsql = "select max(idanomes) as c  from t_cajacierre "

vValor = Val(EsNulo(traerDatos2(vsql, "c", pathDBMySQL)))

If vValor > 0 Then
    If MsgBox("Este més ya fue cerrado. Continúa ? ", vbYesNo) = vbYes Then
        validarCierreMensual = True
    Else
        validarCierreMensual = False
    End If
End If

End Function

Private Sub composicionSaldoMensual()
On Error Resume Next
vidanomes = Val(Str(vvano) + Format(vvmes, "00"))

If Not validarCierreMensual Then Exit Sub

Call PusCerrarCaja_Click
If vidanomes = 0 Then Exit Sub
Call PusImprimir_Click

vidanomes = 0
vidanomesAnterior = 0

If Err Then Exit Sub
End Sub


Private Sub PusComposiciónDe_Click()
Call setcierremensual
Call composicionSaldoMensual

cierremensual = True


'Me.dtpFecha(0) = "01/" + Str(vvmes) + "/" + Str(vvano)
'Me.dtpFecha(1) = DiasDelMes(dtpFecha(0).Value) & "/" & AjustarMes(Month(dtpFecha(0).Value)) & "/" & Year(dtpFecha(0).Value)

'Call PusSalos_Click

'Me.paraBalance.Visible = False
'drBancosSaldo.SetFocus

End Sub

Private Sub PusDesactivarCierre_Click()
On Error Resume Next

Dim vsql, vmensaje  As String

vidcierecaja = grilla.TextMatrix(1, 3)


vmensaje = "Está seguro de cerrar borrar el cierre de caja." + Chr(13) + "Estó permitirá que modifique datos para los movimientos de esa Caja"
 
If MsgBox(vmensaje, vbYesNo) = vbNo And Val(vidcierecaja) > 0 Then Exit Sub


vsql = "delete  from t_cajacierre where  idcajacierre = " + Str(vidcierecaja)
Call EjecutarScript(vsql, pathDBMySQL)

vsql = "delete  from t_cajacierresaldos where  idcajacierre = " + Str(vidcierecaja)
Call EjecutarScript(vsql, pathDBMySQL)

vgid = 0

Call Buscar("", "")


MsgBox "La caja fue abierta", vbInformation


If Err Then Exit Sub
End Sub


Private Sub fbajarVales()
Dim vsql, vset  As String
 
'If Not MsgBox("Está seguro que desea dar de bajas los vales marcados como <Cancelado> ? ", vbYesNo) = vbYes Then Exit Sub


vsql = " insert into bancosmovimientos (TipoMovimiento,idTipoValor,fecha,idbancos,credito,comentario,conciliado) " + _
" select t.TipoMovimiento, t.idTipoValor ,fecha,idbancos,debito, comentario, '-' from bancosmovimientos t " + _
" where conciliado = 'cancelado'"


Call EjecutarScript(vsql)

vsql = "update  bancosmovimientos set conciliado = '-', idtipoValor= 'VA', tipomovimiento = 'VL' where conciliado = 'cancelado'"

Call EjecutarScript(vsql)


Me.fcancelaValores.Visible = False

vctm.Text = "VL"

Call cmdFiltrar_Click

End Sub

Private Sub fpasarVales()
On Error Resume Next
Dim i, vtr As Integer

vtr = Me.KlexDetalle.Rows

With Me.KlexDetalle

       
        For i = 0 To vtr - 1
                
                If .TextMatrix(i, 10) = "Cancelado" Then
                
                    Call cargarRenglonIE(Val(sacarComa(.TextMatrix(i, 7))), Val(.TextMatrix(i, 4)), .TextMatrix(i, 1))
                
                   ' .RemoveItem (i)
        
                End If
        
        Next

End With

If Not frmIngresosEgresos.KlexMovimientoCaja.Rows = 2 Then
    frmIngresosEgresos.KlexMovimientoCaja.Rows = frmIngresosEgresos.KlexMovimientoCaja.Rows - 1
End If
'frmIngresosEgresos.txtAlta(0).Text = "TR"
'frmIngresosEgresos.txtAlta(1).Text = "Transferencia"
'frmIngresosEgresos.vobservacion = "Cancelación de VALES."


If Err Then

    If Not frmIngresosEgresos.KlexMovimientoCaja.Rows = 2 Then
        'frmIngresosEgresos.KlexMovimientoCaja.Rows = frmIngresosEgresos.KlexMovimientoCaja.Rows - 1
    End If

Exit Sub
End If
End Sub

Function sacarComa(vValor As String)
    sacarComa = Replace(vValor, ",", "")
End Function

Function buscarValeDublicado(vid As Long) As Boolean
Dim i As Integer

For i = 0 To Me.KlexDetalle.Rows - 1
    

Next

End Function

Private Sub cargarRenglonIE(vimporte As Double, vnroorden As Long, vidVale As Long)

'If buscarValeDublicado(vidVale) Then
'    Exit Sub
'End If


vr = frmIngresosEgresos.KlexMovimientoCaja.Rows
vr = vr + 1

'If vr > 3 And Not unicavez = 1 Then

'unicavez = 1

'vr = vr + 1
'frmIngresosEgresos.KlexMovimientoCaja.Rows = vr

'End If

frmIngresosEgresos.KlexMovimientoCaja.Rows = vr

vr = vr - 2
With frmIngresosEgresos.KlexMovimientoCaja


    .TextMatrix(vr, 9) = vimporte
    
    .TextMatrix(vr, 8) = "H"
    
    .TextMatrix(vr, 2) = "VAL"
    
    .TextMatrix(vr, 4) = frmIngresosEgresos.dtpFecha.Value
    
    .TextMatrix(vr, 14) = EsNulo(frmIngresosEgresos.vcliprovee.Tag)
    
    .TextMatrix(vr, 5) = "*1001"
    
    .TextMatrix(vr, 10) = "Cancelación de Vale de la orden " + Str(vnroorden)
    
    ' guardo la idbancosmovimientos del vale para poder darlo de baja según los valoes registrados en la grilla
    .TextMatrix(vr, 3) = vidVale
    
    
End With

End Sub


Private Sub PusEjecutar_Click()
 Dim vsql, vset  As String
 
 If MsgBox("Está seguro que desea dar de bajas los vales marcados como <Cancelado> ? ", vbYesNo) = vbYes Then


If vccaja.Text = "" Then
    
    MsgBox "No hay Caja seleccionada", vbInformation
    Exit Sub

End If

vset = "idbancos ='" + vccaja.Text + "'" + _
", TipoMovimiento = 'EF'" + _
", debito = credito " + _
", idTipoValor = 'EF'"


            vsql = "update bancosmovimientos  set " + vset + " where conciliado  = 'Cancelado'"
            Call EjecutarScript(vsql, pathDBMySQL)
            
vset = "idbancos ='" + vccaja.Text + "'" + _
", TipoMovimiento = 'EF'" + _
", credito =0" + _
", conciliado  = ''"


            vsql = "update bancosmovimientos  set " + vset + " where conciliado  = 'Cancelado'"
            Call EjecutarScript(vsql, pathDBMySQL)
    
    End If

Me.fcancelaValores.Visible = False

vctm.Text = "VL"

Call cmdFiltrar_Click

End Sub

Private Sub PusExportarA_Click()
On Error Resume Next
Set Me.g.Recordset = Me.KlexDetalle.Recordset
    
  Call grillaToExcel2(Me.KlexDetalle)

If Err Then Exit Sub
End Sub

Private Sub PusFinDe_Click()


Me.paraBalance.Visible = True

vvmes = Month(Date)

vvano = Year(Date)

Me.vidanomes_label.Text = Trim(Str(vvano)) + Trim(Str(vvmes))

vvmes.SetFocus

End Sub

Private Sub PushButton1_Click()
Dim vid As Long
Dim vsql As String

vsql = "select nrointerno as c from bancosmovimientos where idbancosmovimientos=" + Me.KlexDetalle.TextMatrix(Me.KlexDetalle.Row, 1)

Call verTransacciones(traerDatos2(vsql, "c", pathDBMySQL))
End Sub

Private Sub PushButton10_Click()
Call mostrarTiposCierre("todos")
End Sub

Private Sub PushButton11_Click()
frmMovimientosCuentas.Show

frmMovimientosCuentas.dtpCuentas(0).Value = Me.vfecha
frmMovimientosCuentas.dtpCuentas(1).Value = Me.vfecha

End Sub

Private Sub PushButton12_Click()
frmMovimientosCuentas.Show

frmMovimientosCuentas.dtpCuentas(0).Value = CDate("01/" + Str(Month(Me.vfecha)) + "/" + Str(Year(Me.vfecha)))

'Call frmMovimientosCuentas.dtpCuentas_KeyPress(0, 13)

'frmMovimientosCuentas.dtpCuentas(0).Value = CDate("01/" + Str(Month(Me.vfecha)) + "/" + Str(Year(Me.vfecha)))

frmMovimientosCuentas.Show

End Sub

Private Sub PushButton13_Click()
Dim v As String
Dim i As Integer
Dim finn As Integer
Dim vcomp As String

v = InputBox("Escriba la palabra clave: dalas2" + Chr(13) + "Si quiere PDF debe predeterminar PRIMIPDF")

If v = "dalas2" Then

    reimprimirTodo = 1
    
    finn = Me.KlexDetalle.Rows - 1
    
    For i = 1 To finn
        vcomp = Me.KlexDetalle.TextMatrix(i, 4)
        ActualizarRecibo (Val(vcomp))
    Next

End If

End Sub

Private Sub PushButton2_Click()
Me.vtipolistado.Text = "Con detalles de Movimientos contables"

dtpFecha(0).Value = Me.vfecha.Value
dtpFecha(1).Value = Me.vfecha.Value



Call ImprimeCajaConDetalle


End Sub

Private Sub PushButton3_Click()
Call fbuscarGrilla(" tipomovimientos ", "TipoMovimiento", "Codigo", Me.vdtm.Name, Me)    ' ema:
End Sub

Private Sub PushButton4_Click()
Dim vsql As String


Call fpasarVales

Unload Me
'Call fbajarVales

Exit Sub


If vaccion = "Consulta.Vales" Then
   Me.fcancelaValores.Visible = True
End If

End Sub

Private Sub PushButton5_Click()
    MsgBox "El detalle de movimiento por cuenta se debe ejecutar desdesde el módulo contabilida > movimientos por cuentas." + Chr(13) + _
    "Ahora puede continuar. Tenga en cuenta que el listado puede tardar mucho tiempo dependiendo del rango de fecha seleccionado"
   
   Call fbuscarGrilla("(select * from cuentas where Imputable ='S') as t", "Cuenta", "CodigoCuenta", Me.vCtasDescrip.Name, Me)    ' ema:
End Sub

Private Sub PushButton6_Click()
    frmBancoCajaDetalle.vtipolistado = "Movimientos Caja - Banco con movimientos Contables"
   ' frmBancoCajaDetalle.dtpFecha(0).Value = Me.vfecha
   ' frmBancoCajaDetalle.dtpFecha(1).Value = Me.vfecha
    
    vvviene = "movidiario"
    
    'Call frmBancoCajaDetalle.cmdFiltrar_Click
    
   ' Call Reporte_MoviCBA(dtpFecha(0).Value, dtpFecha(1).Value, Val(vidAdesde), Val(vidAhasta))
    
     Call Reporte_MoviCBA(dtpFecha(0).Value, dtpFecha(1).Value, Val(vidd), Val(vidH), chkComprobantesNo.Value)
    
    
    vvviene = ""
   
End Sub




Private Sub PushButton7_Click()
On Error Resume Next
Dim r As Integer
Dim vsql, valor, vtipo As String
Dim vid As Long


r = Me.KlexDetalle.Row

valor = Me.KlexDetalle.TextMatrix(r, 10)

vid = Me.KlexDetalle.TextMatrix(r, 1)

vtipo = Me.KlexDetalle.TextMatrix(r, 0)


If Me.KlexDetalle.TextMatrix(r, 10) = "-" Then
        
        If MsgBox("Está segura de marcar a este vale como NO cancelado ?", vbYesNo) = vbYes Then
             Me.KlexDetalle.TextMatrix(r, 10) = ""
             vsql = "update bancosmovimientos set Conciliado='' where idBancosMovimientos=" + Str(vid)
             Call EjecutarScript(vsql, pathDBMySQL)
        End If
End If

End Sub

Private Sub PushButton8_Click()
Call mostrarTiposCierre("mensual")
End Sub

Private Sub PushButton9_Click()
Call mostrarTiposCierre("diario")
End Sub

Private Sub PusImprimir_Click()
Dim vd, vc As Double

vd = fingresos(vidd, vidH)

vc = fegresos(vidd, vidH)

If grilla2.TextMatrix(1, 1) = "" Then
    MsgBox "No hay datos para mostrar", vbInformation
    Exit Sub
End If


With Mantenimiento.rsCajaCierreSaldos


If .State = 1 Then .Close

.Source = " select ingresos, egresos, " + _
"  fecha, " + _
"  t_cajacierresaldos.idbancos, " + _
"  bancos.descripcion, " + _
"  t_cajacierresaldos.santerior, " + _
"  t_cajacierresaldos.speriodo, " + _
"  t_cajacierresaldos.saldo " + _
"  from t_cajacierre " + _
"  inner join t_cajacierresaldos on " + _
"    (t_cajacierre.idcajacierre = t_cajacierresaldos.idcajacierre) " + _
"  inner join bancos on " + _
"    bancos.idBancos = t_cajacierresaldos.idBancos " + _
" Where " + _
"  t_cajacierresaldos.idcajacierre = " + Str(vidcierecaja) + " " + _
" order by t_cajacierresaldos.idbancos asc "
        
        
If .State = 0 Then .Open
        .Close
        .Open
End With
  

drCajaCierreSaldos.Sections("totales").Controls("eSIngresos").Caption = Format((tsd), "###,###,##0.00")

drCajaCierreSaldos.Sections("totales").Controls("eSEgresos").Caption = Format((tsc), "###,###,##0.00")

Dim vanomes_mostrar As Long

If vidanomes_label.Text = "" Then
    vanomes_mostrar = Val(vidanomes_label.Text)
Else
    vanomes_mostrar = vidanomes
End If


If vidanomes > 0 Then
    drCajaCierreSaldos.Sections("TituloEmpresa").Controls("efiltro").Caption = "Balance de composición de saldos correspontiende al mes: " + Format(vanomes_mostrar, "0000/00")
Else
    drCajaCierreSaldos.Sections("TituloEmpresa").Controls("efiltro").Caption = "Composición desaldos correspondiente al día: " + grilla2.TextMatrix(1, 1) + " - Se observa lo siguiente: " + grilla.TextMatrix(grilla.RowSel, 2)
End If

drCajaCierreSaldos.Show


' ---- init ---
vsd = 0
vsc = 0
'--------------

'Call imprimirGrilla(Me.grilla2, 5)

End Sub


Function fegresos(ByVal viddesde As Long, ByVal vidhasta As Long) As Double
On Error Resume Next

Dim vsql As String

vsql = " select " + _
" sum(debito) as i, " + _
" sum(credito) as e " + _
" from bancosmovimientos t  where not t.idbancos = '098' and  t.idBancosMovimientos  >= " + Str(viddesde) + " and   t.idBancosMovimientos  <= " + Str(vidhasta)

fegresos = Val(traerDatos2(vsql, "e", pathDBMySQL))

If Err Then
    fegresos = 0
    Exit Function
End If

End Function




Function fingresos(ByVal viddesde As Long, ByVal vidhasta As Long) As Double
On Error Resume Next

Dim vsql As String

vsql = " select " + _
" sum(debito) as i, " + _
" sum(credito) as e " + _
" from bancosmovimientos t  where not t.idbancos = '098' and t.idBancosMovimientos  >= " + Str(viddesde) + " and   t.idBancosMovimientos  <= " + Str(vidhasta)

fingresos = Val(traerDatos2(vsql, "i", pathDBMySQL))

If Err Then
    fingresos = 0
    Exit Function
End If

End Function


Private Sub PusListo_Click()


If cierremensual Then
        ' me paro en la primera
        Me.grilla.SetFocus
        Me.grilla.Row = 1
        cierremensual = False
End If


'frmBalance.dtpFecha(0) = "01/" + Str(vvmes) + "/" + Str(vvano)
'Call frmBalance.dtpFecha_KeyPress(0, 13)

viddesde = vidAdesde
vidhasta = vidAhasta


frmBalance.vmesDelBalance = informeDelCierre(vidAdesde, vidhasta)


If vidAdesde = 0 Then
    MsgBox "Debe seleccionar un período cerrado"
    Exit Sub
End If


Call frmBalance.PbAcciones_Click(0)

viddesde = 0
vidhasta = 0

End Sub



Private Sub PusMensual_Click()
dtpFecha(0).Value = Date - Day(Date)
End Sub

Private Sub PusMostrarTodo_Click()
Call Buscar("", " where 1=2")
End Sub

Private Sub PusNroComp_Click()
On Error Resume Next
Dim vfd, vfh As Date


vfd = InputBox("Ingrese fecha desde:")

vfh = InputBox("Ingeres fecha hasta:")

'If Str(vfh) = "" Or Str(vfd) = "" Then Exit Sub
Call listar_nrocomprabantes_noimputados(0, 0, vfd, vfh)

If Err Then Exit Sub
End Sub

Private Sub PusPaso3_Click()
On Error Resume Next
    frmBancoCajaDetalle.vtipolistado = "Todos"
    frmBancoCajaDetalle.dtpFecha(0).Value = Me.vfecha
    frmBancoCajaDetalle.dtpFecha(1).Value = Me.vfecha
    
   ' Me.vGidasientos = grilla.TextMatrix(vr, 4)
    Me.vGiddesdebm = grilla.TextMatrix(vr, 5)
    Me.vGidhastabm = grilla.TextMatrix(vr, 6)
    
    chkFecha.Value = xtpChecked
    
    Call frmBancoCajaDetalle.cmdFiltrar_Click
    Call frmBancoCajaDetalle.cmdImprimir_Click(1)
If Err Then

    MsgBox "Debe seleccionar un cierre de caja", vbInformation
    Exit Sub
End If
End Sub

Private Sub PusRealizarArqueos_Click()


MsgBox "Esta función no está habilitada por el momento", vbInformation

Exit Sub


Call imprimeSaldoCajas
Call imprimeMovimientosCaja
Me.tabbc.SelectedItem = 2

MsgBox "Recuerde cerrar la caja", vbInformation
Me.vtipolistado = "Con detalles de Movimientos contables"

impDirecto = True

' todo: poner la fecha correspondiente al último cierre de caja.
' no permitir modificar datos

'Me.dtpFecha(0).Value = Date
'Me.dtpFecha(1).Value = Date

Call cmdImprimir_Click(1)


End Sub

Private Sub imprimeMovimientosCaja()

End Sub


Private Sub imprimeSaldoCajas()
Me.vtipolistado = "Agrupado por Cajas-Bancos"

impDirecto = True

' todo: poner la fecha correspondiente al último cierre de caja.
' no permitir modificar datos

'Me.dtpFecha(0).Value = Date
'Me.dtpFecha(1).Value = Date

Call cmdImprimir_Click(1)

End Sub

Private Sub txtAlta_Change(Index As Integer)

End Sub

Private Sub PusReimprimirComprobante_Click()
On Error Resume Next
Dim vnrocomprobante As Long

vnrocomprobante = Me.KlexDetalle.TextMatrix(vr, 4)

ActualizarRecibo (vnrocomprobante)


If Err Then Exit Sub
End Sub

Private Sub PusSalos_Click()
    frmBancoCajaDetalle.Hide
    frmBancoCajaDetalle.vtipolistado = "Saldos"
    Call frmBancoCajaDetalle.cmdFiltrar_Click
    Unload frmBancoCajaDetalle
End Sub

Private Sub PusSemanal_Click()
dtpFecha(0).Value = Date - 7
dtpFecha(1).Value = Date
End Sub

Private Sub tabbc_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Me.paraBalance.Visible = False
End Sub

Private Sub vCtasDescrip_Change()
On Error Resume Next
Dim vsql As String

vctasCodigo.Text = vCtasDescrip.Tag

If Err Then Exit Sub
End Sub

Private Sub vdbancocaja_Change()
Me.vcbancocaja.Text = Me.vdbancocaja.Tag
End Sub

Private Sub vdcaja_Change()
vccaja.Text = vdcaja.Tag
End Sub

Private Sub vdtm_Change()
vctm.Text = vdtm.Tag
End Sub

Private Sub vfecha_Change()
On Error Resume Next
Call Buscar("where fecha = '" + strfechaMySQL(vfecha) + "'", "")
If Err Then Exit Sub
End Sub

Private Sub formatearGrilla()
Dim i, j, k As Integer
k = grilla.Cols - 1
For i = 1 To grilla.Rows - 1

If Val(grilla.TextMatrix(i, 8)) > 0 Then

    For j = 1 To k
        grilla.Row = i
        grilla.Col = j
        grilla.CellBackColor = &HFF8080
    Next
    
End If


Next
End Sub

Private Sub mostrarTiposCierre(vtipo As String)
Dim rs, rs2, rs3, rs4 As New ADODB.Recordset


Dim rsP As New ADODB.Recordset

Dim vsql As String

If vtipo = "mensual" Then
    vsql = "select fecha, estado, idcajacierre, idasientosdesde, iddesde, idhasta, idasientoshasta, idanomes  from t_cajacierre  where idanomes > 0  order by idcajacierre desc"
End If

If vtipo = "diario" Then
    vsql = "select fecha, estado, idcajacierre, idasientosdesde, iddesde, idhasta, idasientoshasta, idanomes  from t_cajacierre  where idanomes = 0  order by idcajacierre desc"
End If

If vtipo = "todos" Then
    vsql = "select fecha, estado, idcajacierre, idasientosdesde, iddesde, idhasta, idasientoshasta, idanomes  from t_cajacierre  order by idcajacierre desc"
End If


'vsql = "select fecha, estado, idcajacierre, idasientosdesde, iddesde, idhasta, idasientoshasta, idanomes  from t_cajacierre " + vw1 + " order by idcajacierre desc"
Call rsP.Open(vsql, ConnDDBB, adOpenStatic, adLockPessimistic)

If Not rsP.RecordCount > 0 Then Exit Sub

Set grilla.DataSource = rsP.DataSource
grilla.Refresh


formatearGrilla


End Sub

Private Sub Buscar(vw1, vwhere As String)
On Error Resume Next
Dim rs, rs2, rs3, rs4 As New ADODB.Recordset


Dim rsP As New ADODB.Recordset

Dim vsql, vcampos As String

Set grilla2.DataSource = Nothing
Set grilla.DataSource = Nothing

grilla.Refresh
grilla2.Refresh

grilla.Clear
grilla2.Clear

vsql = "select fecha, estado, idcajacierre, idasientosdesde, iddesde, idhasta, idasientoshasta, idanomes  from t_cajacierre " + vw1 + " order by idcajacierre desc"

'vsql = "select fecha, estado, idcajacierre, idasientosdesde, iddesde, idhasta, idasientoshasta, idanomes  from t_cajacierre " + vw1 + " order by idcajacierre desc"
Call rsP.Open(vsql, ConnDDBB, adOpenStatic, adLockPessimistic)

If Not rsP.RecordCount > 0 Then Exit Sub

Set grilla.DataSource = rsP.DataSource
grilla.Refresh


formatearGrilla


' --------------------------------------------------------



vcampos = "fecha, t_cajacierresaldos.idbancos, bancos.Descripcion, t_cajacierresaldos.santerior, t_cajacierresaldos.speriodo, t_cajacierresaldos.saldo"

vsql = " SELECT " + vcampos + " From  `t_cajacierre` " + _
" INNER JOIN `t_cajacierresaldos` ON (`t_cajacierre`.`idcajacierre` = `t_cajacierresaldos`.`idcajacierre`) " + _
" inner join bancos on bancos.idbancos=t_cajacierresaldos.idbancos " + vwhere + " order by t_cajacierresaldos.idbcajacierresaldos desc"

Call rs4.Open(vsql, ConnDDBB, adOpenStatic, adLockPessimistic)

If Not rs4.RecordCount > 0 Then Exit Sub

Set grilla2.DataSource = rs4.DataSource
grilla2.Refresh

formatearGrilla2

Set Mantenimiento.rsCajaCierreSaldos.DataSource = rs4.DataSource

rs4.Close

If Err Then Exit Sub
End Sub


Private Sub formatearGrilla2()
Dim i As Integer

For i = 1 To grilla2.Rows - 1
    grilla2.TextMatrix(i, 3) = Format(grilla2.TextMatrix(i, 3), "###,###,##0.00")
    grilla2.TextMatrix(i, 4) = Format(grilla2.TextMatrix(i, 4), "###,###,##0.00")
    grilla2.TextMatrix(i, 5) = Format(grilla2.TextMatrix(i, 5), "###,###,##0.00")
    grilla2.TextMatrix(i, 6) = Format(grilla2.TextMatrix(i, 6), "###,###,##0.00")
Next


End Sub



Private Sub configurarGrid()


'vcampos = "idproveedores,fdesde,fhasta,horas,horasextras,valorhora,valorhoraextra,estado,comentarios,idTipoMovimientos"


'1.`idtrabeventuales`,
'2.`codPersona`,
'3. `nombre`,
'4.  `fdesde`,
'5. `fhasta`,
'6. `horas`,
'7. `horasextras`,
'8. `valorhora`,
'9. `valorhoraextra`,
'10.`codTrabajo`,
'11.`tipomovimiento`,
'12. `comentarios`,
'13.`estado`,
'14.`importe`,
'16. `idproveedores`,
'17. `idtipomovimientos`

'Me.grilla.Cols = 3
'Me.grilla.ColWidth(0) = 1000 'idproveedores
'Me.grilla.ColWidth(1) = 0
'Me.grilla.ColWidth(2) = 2000
'Me.grilla.ColWidth(3) = 7000


End Sub

Private Sub vmes_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub vtipovalor_Change()
    vctipovalor.Text = vtipovalor.Tag
End Sub

Private Sub llenarDrRecibo(ByVal vtotal As Double, rs As Recordset)
On Error Resume Next

Dim vsql, vdtm, vnombre, vcomentario1, vcomentario2 As String

'vvsaldo = Format(CalSaldoPersona(Me.txtCliente(0).Text, CP.TablaCtaCte), "$###,###,##0.00")
 
'vnro = Str(getNroRecibo)
 

Unload Mantenimiento
Load Mantenimiento
 
 rs.MoveLast
 
 If rs.Fields("TipoMovimiento") = "VL" Then
 
    vtotal = vtotal / 2
    
 End If
 rs.MoveFirst
 vsql = "Select * from tipomovimientos where Codigo = '" + rs.Fields("TipoMovimiento") + "'"
vdtm = traerDatos2(vsql, "TipoMovimiento", pathDBMySQL)
                
vnombre = ""

 
vsql = "select nombre from proveedores where codigo = '" + rs.Fields("codpersona") + "'"
        
vnombre = "Por cuenta de: " + traerDatos2(vsql, "nombre", pathDBMySQL) ' todo
        
 
 With drRecibo
 
        vsql = "select nombre from proveedores where codigo = '" + rs.Fields("codpersona") + "'"
        
        .Sections(2).Controls("lblCliente").Caption = vnombre ' todo
        
        .Sections(5).Controls("lblconcepto").Caption = Trim(rs.Fields("comentario2"))
        .Sections(5).Controls("lblconcepto2").Caption = Trim(rs.Fields("comentario3"))
 
 
            .Sections(2).Controls("enrocomprobante").Caption = rs.Fields("nrocomprobante")
    
            .Sections(5).Controls("eletras").Caption = EnLetras(Str(vtotal))
            .Sections(5).Controls("lbltotal").Caption = Format(vtotal, "###,###,##0.00")
          
        
           .Sections("TituloEmpresa").Controls("etiqueta1").Caption = "REIMPRESIÓN de comprobante operaciones de Caja"
        .Sections(2).Controls("econcepto").Caption = Trim(vdtm)
        
        If rs.Fields("TipoMovimiento") = "TR" Then
        
            .Sections(5).Controls("etitulototal").Caption = ""
            .Sections(5).Controls("lbltotal").Caption = ""
            .Sections(5).Controls("eletras").Caption = ""
        
        Else
            
            .Sections(5).Controls("eletras").Caption = EnLetras(Str(vtotal))
            .Sections(5).Controls("lbltotal").Caption = Format(vtotal, "###,###,##0.00")
          
        End If
        
        
      
        .Sections(5).Controls("esaldoTitulo").Caption = ""
        .Sections(5).Controls("esaldo").Caption = ""
 
 
         
        
  
        '.Sections(2).Controls("etipo").Caption = Trim(Me.txtAlta(4))

        
        
        .Sections(2).Controls("etiqueta9").Caption = rs.Fields("fecha")
        '.Sections(2).Controls("lbllugar").Caption = vDatosEmpresa.Localidad & ", "
        .Sections(2).Controls("lblfecha").Caption = rs.Fields("fecha")
        
        
        
        
        
'        .Sections(2).Controls("lblCliente").Caption = "Por cuenta de: " + traerDatos2(vsql, "nombre", pathDBMySQL) ' todo
        
'        .Sections(5).Controls("lblconcepto").Caption = rs.Fields("comentario2")
'        .Sections(5).Controls("lblconcepto2").Caption = rs.Fields("comentario3")
        
        
    
          
        
        
        If rs.Fields("TipoMovimiento") = "TR" Then
        
            .Sections(5).Controls("etitulototal").Caption = ""
            .Sections(5).Controls("lbltotal").Caption = ""
            .Sections(5).Controls("eletras").Caption = ""
        
        Else
            
           ' .Sections(5).Controls("eletras").Caption = EnLetras(Str(vtotal))
           ' .Sections(5).Controls("lbltotal").Caption = Str(Format(vtotal, "###,###,##0.00"))
          
        End If
        
        
      
        .Sections(5).Controls("esaldoTitulo").Caption = ""
        .Sections(5).Controls("esaldo").Caption = ""
        
        
        Debug.Print "- Saldo en letra : " + .Sections(5).Controls("eletras").Caption
        Debug.Print "- Saldo:" + .Sections(5).Controls("lbltotal").Caption
        
        
        
       '.Sections(5).Controls("esaldo").Caption = Format(vvsaldo, "$ ###,###,##0.00")
        
       ' .Hide
        
        If reimprimirTodo = 1 Then
            .Hide
            .PrintReport False
        End If
    
    
    End With




If Err Then
    'MsgBox "Error al intentar hacer el recibo" + Str$(Err)
    Exit Sub
End If

End Sub


Function ActualizarRecibo(vnrocomprobante As Long) As Double
On Error Resume Next
Dim vsql, vlinea As String
Dim vsql1, vc, vd As String
Dim rss As New ADODB.Recordset
Dim vtotal, vimporte As Double

If vnrocomprobante = 0 Then Exit Function

vtotal = 0
vimporte = 0

vsql = "select * from bancosmovimientos where nrocomprobante=" + Str(vnrocomprobante)

Call rss.Open(vsql, ConnDDBB, adOpenStatic, adLockReadOnly)
 
 If Not rss.RecordCount > 0 Then
    Exit Function
 End If
 
Dim i As Integer

vsql = "delete from recibo_temp"
Call EjecutarScript(vsql, pathDBMySQL)

vtotal = 0

rss.MoveFirst

    Do Until rss.EOF

        vlinea = rss.Fields("comentario")
        
        vimporte = Val(EsNulo(rss.Fields("debito"))) + Val(EsNulo(rss.Fields("credito")))
        
        vtotal = vtotal + vimporte
        
        vsql = "insert into recibo_temp (descripcion,monto) values ('" + vlinea + "'," + Str(vimporte) + ") "
        
       If Not vlinea = "" Then Call EjecutarScript(vsql, pathDBMySQL)
       
       rss.MoveNext
    
    
    Loop
    
    
Call llenarDrRecibo(vtotal, rss)


If Err Then
    Exit Function
    'MsgBox Err.Description
    'Exit Sub
End If

End Function

Private Sub vvano_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    PusListo.SetFocus
End If

End Sub

Private Sub vvmes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    vvano.SetFocus
End If

End Sub

Public Function estanCancelandoVale() As Boolean
Dim vsql As String
estanCancelandoVale = False

vsql = "select count(idbancosmovimientos) as c from bancosmovimientos where conciliado = 'Cancelado'"

If Val(traerDatos2(vsql, "c", pathDBMySQL)) > 0 Then
    estanCancelandoVale = True
End If


End Function
