VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.ShortcutBar.v13.0.0.Demo.ocx"
Begin VB.Form frmDeudasServicios2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estado de  Ctas de Servcios "
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   13425
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   8880
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   13305
      _Version        =   851968
      _ExtentX        =   23469
      _ExtentY        =   15663
      _StockProps     =   68
      ItemCount       =   5
      Item(0).Caption =   "Libre Deuda"
      Item(0).ControlCount=   33
      Item(0).Control(0)=   "vcomentario"
      Item(0).Control(1)=   "PusImprimir"
      Item(0).Control(2)=   "vcontribuyente"
      Item(0).Control(3)=   "GroupBox1"
      Item(0).Control(4)=   "PusGenerarListado"
      Item(0).Control(5)=   "barra"
      Item(0).Control(6)=   "PusPersonas"
      Item(0).Control(7)=   "GroupBox2"
      Item(0).Control(8)=   "PusLimpiar"
      Item(0).Control(9)=   "GroupBox3"
      Item(0).Control(10)=   "PusExportar"
      Item(0).Control(11)=   "lblAgregarComentario"
      Item(0).Control(12)=   "PusCargarInformación"
      Item(0).Control(13)=   "lblDeudaAcumulada"
      Item(0).Control(14)=   "vsaldo"
      Item(0).Control(15)=   "grilla2"
      Item(0).Control(16)=   "v2"
      Item(0).Control(17)=   "vid"
      Item(0).Control(18)=   "vperiodo"
      Item(0).Control(19)=   "PusBorrar"
      Item(0).Control(20)=   "vimporte"
      Item(0).Control(21)=   "PusAgregar"
      Item(0).Control(22)=   "lblImporte"
      Item(0).Control(23)=   "lblPeríodo"
      Item(0).Control(24)=   "Label1"
      Item(0).Control(25)=   "vnrorecibo"
      Item(0).Control(26)=   "PusExcel"
      Item(0).Control(27)=   "vlibredeuda"
      Item(0).Control(28)=   "GroupBox4"
      Item(0).Control(29)=   "GroPeríodo"
      Item(0).Control(30)=   "GroOrdenadoPor"
      Item(0).Control(31)=   "grilla22"
      Item(0).Control(32)=   "PusExcel2"
      Item(1).Caption =   "Datos Generales"
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "grilla"
      Item(1).Control(1)=   "PushButton1"
      Item(1).Control(2)=   "listamenu"
      Item(1).Control(3)=   "MSHFlexGrid1"
      Item(2).Caption =   "Indicadores Generales"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "indicadores"
      Item(3).Caption =   "Proyecciones"
      Item(3).ControlCount=   0
      Item(4).Caption =   "Item"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "g2e"
      Begin XtremeSuiteControls.GroupBox GroOrdenadoPor 
         Height          =   510
         Left            =   4770
         TabIndex        =   58
         Top             =   1125
         Width           =   3030
         _Version        =   851968
         _ExtentX        =   5345
         _ExtentY        =   900
         _StockProps     =   79
         Caption         =   "Ordenado por: "
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton rdcodigo 
            Height          =   240
            Left            =   1080
            TabIndex        =   59
            Top             =   225
            Width           =   870
            _Version        =   851968
            _ExtentX        =   1535
            _ExtentY        =   423
            _StockProps     =   79
            Caption         =   "Código"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rdnombre 
            Height          =   240
            Left            =   1980
            TabIndex        =   60
            Top             =   225
            Width           =   870
            _Version        =   851968
            _ExtentX        =   1535
            _ExtentY        =   423
            _StockProps     =   79
            Caption         =   "Nombre"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GroPeríodo 
         Height          =   735
         Left            =   1890
         TabIndex        =   48
         Top             =   2520
         Width           =   11220
         _Version        =   851968
         _ExtentX        =   19791
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Período:"
         UseVisualStyle  =   -1  'True
         Begin VB.CheckBox chkTitulo 
            Caption         =   "Mostrar texto en el listado"
            Height          =   285
            Left            =   8460
            TabIndex        =   57
            Top             =   270
            Width           =   2445
         End
         Begin MSComCtl2.DTPicker vPeriodoDesde 
            Height          =   375
            Left            =   4050
            TabIndex        =   54
            Top             =   225
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "MM/yyyy"
            Format          =   80871427
            CurrentDate     =   43112
         End
         Begin XtremeSuiteControls.RadioButton RadioButton2 
            Height          =   330
            Left            =   135
            TabIndex        =   49
            Top             =   270
            Width           =   735
            _Version        =   851968
            _ExtentX        =   1296
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Todos"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rdPagos 
            Height          =   330
            Left            =   945
            TabIndex        =   50
            Top             =   270
            Width           =   735
            _Version        =   851968
            _ExtentX        =   1296
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Pagos"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rdImpagos 
            Height          =   330
            Left            =   1800
            TabIndex        =   51
            Top             =   270
            Width           =   960
            _Version        =   851968
            _ExtentX        =   1693
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Impagos"
            UseVisualStyle  =   -1  'True
         End
         Begin MSComCtl2.DTPicker vPeriodoHasta 
            Height          =   375
            Left            =   6525
            TabIndex        =   55
            Top             =   225
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "MM/yyyy"
            Format          =   80871427
            CurrentDate     =   43112
         End
         Begin XtremeSuiteControls.Label lblHasta 
            Height          =   330
            Left            =   5895
            TabIndex        =   53
            Top             =   225
            Width           =   690
            _Version        =   851968
            _ExtentX        =   1217
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Hasta:"
         End
         Begin XtremeSuiteControls.Label lblDesde 
            Height          =   330
            Left            =   3195
            TabIndex        =   52
            Top             =   225
            Width           =   690
            _Version        =   851968
            _ExtentX        =   1217
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Desde:"
         End
      End
      Begin XtremeSuiteControls.PushButton PusExcel 
         Height          =   405
         Left            =   3645
         TabIndex        =   43
         Top             =   4185
         Width           =   1065
         _Version        =   851968
         _ExtentX        =   1879
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Excel"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit vnrorecibo 
         Height          =   345
         Left            =   6435
         TabIndex        =   42
         Top             =   8430
         Width           =   1485
         _Version        =   851968
         _ExtentX        =   2619
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin VB.TextBox vid 
         Height          =   375
         Left            =   9165
         TabIndex        =   32
         Top             =   8430
         Width           =   1035
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid listamenu 
         Height          =   6615
         Left            =   -59230
         TabIndex        =   29
         Top             =   480
         Visible         =   0   'False
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   11668
         _Version        =   393216
         ForeColor       =   4210752
         Cols            =   1
         FixedCols       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Baskerville Old Face"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   1
      End
      Begin XtremeSuiteControls.PushButton PusCargarInformación 
         Height          =   315
         Left            =   4800
         TabIndex        =   26
         Top             =   780
         Width           =   3075
         _Version        =   851968
         _ExtentX        =   5424
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Cargar información del Contribuyente"
         UseVisualStyle  =   -1  'True
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid indicadores 
         Height          =   3555
         Left            =   -69880
         TabIndex        =   25
         Top             =   420
         Visible         =   0   'False
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   6271
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid g2e 
         Height          =   3345
         Left            =   -69820
         TabIndex        =   24
         Top             =   510
         Visible         =   0   'False
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   5900
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla 
         Height          =   6615
         Left            =   -69880
         TabIndex        =   22
         Top             =   450
         Visible         =   0   'False
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   11668
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.TextBox vcontribuyente 
         Height          =   345
         Left            =   1680
         TabIndex        =   3
         Top             =   420
         Width           =   6195
      End
      Begin VB.TextBox vcomentario 
         Height          =   345
         Left            =   90
         TabIndex        =   1
         Top             =   1080
         Width           =   4470
      End
      Begin XtremeSuiteControls.PushButton PusImprimir 
         Height          =   405
         Left            =   120
         TabIndex        =   2
         Top             =   4200
         Width           =   1515
         _Version        =   851968
         _ExtentX        =   2672
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Imprimir"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   765
         Left            =   180
         TabIndex        =   4
         Top             =   1485
         Width           =   1545
         _Version        =   851968
         _ExtentX        =   2725
         _ExtentY        =   1349
         _StockProps     =   79
         Caption         =   "Tipo:"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton rdResumen 
            Height          =   180
            Left            =   360
            TabIndex        =   5
            Top             =   210
            Width           =   1035
            _Version        =   851968
            _ExtentX        =   1826
            _ExtentY        =   317
            _StockProps     =   79
            Caption         =   "Resumen"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rdDetalle 
            Height          =   315
            Left            =   360
            TabIndex        =   6
            Top             =   405
            Width           =   1125
            _Version        =   851968
            _ExtentX        =   1984
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Detalle"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.PushButton PusGenerarListado 
         Height          =   405
         Left            =   180
         TabIndex        =   7
         Top             =   3375
         Width           =   12960
         _Version        =   851968
         _ExtentX        =   22860
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Generar Listado"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ProgressBar barra 
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   3870
         Width           =   12885
         _Version        =   851968
         _ExtentX        =   22728
         _ExtentY        =   344
         _StockProps     =   93
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton PusPersonas 
         Height          =   390
         Left            =   120
         TabIndex        =   9
         Top             =   390
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2566
         _ExtentY        =   688
         _StockProps     =   79
         Caption         =   "Contribuyente"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmDeudasServicios2.frx":0000
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   825
         Left            =   4650
         TabIndex        =   10
         Top             =   1635
         Width           =   3195
         _Version        =   851968
         _ExtentX        =   5636
         _ExtentY        =   1455
         _StockProps     =   79
         Caption         =   "Agrupado por:"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton rdContribuyente 
            Height          =   315
            Left            =   120
            TabIndex        =   11
            Top             =   210
            Width           =   1665
            _Version        =   851968
            _ExtentX        =   2937
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Contribuyentes"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rdPeriodo 
            Height          =   270
            Left            =   120
            TabIndex        =   12
            Top             =   510
            Width           =   1335
            _Version        =   851968
            _ExtentX        =   2355
            _ExtentY        =   476
            _StockProps     =   79
            Caption         =   "Períodos"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RadioButton5 
            Height          =   315
            Left            =   1950
            TabIndex        =   13
            Top             =   240
            Width           =   915
            _Version        =   851968
            _ExtentX        =   1614
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Importes"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RadZona 
            Height          =   270
            Left            =   1950
            TabIndex        =   14
            Top             =   510
            Width           =   975
            _Version        =   851968
            _ExtentX        =   1720
            _ExtentY        =   476
            _StockProps     =   79
            Caption         =   "Zona"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.PushButton PusLimpiar 
         Height          =   465
         Left            =   12090
         TabIndex        =   15
         Top             =   4170
         Width           =   885
         _Version        =   851968
         _ExtentX        =   1561
         _ExtentY        =   820
         _StockProps     =   79
         Caption         =   "Limpiar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   885
         Left            =   1920
         TabIndex        =   16
         Top             =   1590
         Width           =   2625
         _Version        =   851968
         _ExtentX        =   4630
         _ExtentY        =   1561
         _StockProps     =   79
         Caption         =   "Zona: "
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton rdUrbano 
            Height          =   315
            Left            =   90
            TabIndex        =   17
            Top             =   210
            Width           =   1125
            _Version        =   851968
            _ExtentX        =   1984
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Urbano"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rdRural 
            Height          =   315
            Left            =   90
            TabIndex        =   18
            Top             =   510
            Visible         =   0   'False
            Width           =   1125
            _Version        =   851968
            _ExtentX        =   1984
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Rural"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RdComercio 
            Height          =   315
            Left            =   1320
            TabIndex        =   19
            Top             =   210
            Width           =   1125
            _Version        =   851968
            _ExtentX        =   1984
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Comercio"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RadioButton1 
            Height          =   315
            Left            =   1320
            TabIndex        =   38
            Top             =   510
            Visible         =   0   'False
            Width           =   1125
            _Version        =   851968
            _ExtentX        =   1984
            _ExtentY        =   556
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.PushButton PusExportar 
         Height          =   405
         Left            =   1650
         TabIndex        =   20
         Top             =   4200
         Width           =   1515
         _Version        =   851968
         _ExtentX        =   2672
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Exportar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   405
         Left            =   -58240
         TabIndex        =   23
         Top             =   7410
         Visible         =   0   'False
         Width           =   705
         _Version        =   851968
         _ExtentX        =   1244
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Exportar"
         UseVisualStyle  =   -1  'True
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla2 
         Height          =   3420
         Left            =   90
         TabIndex        =   30
         Top             =   4770
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   6033
         _Version        =   393216
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin XtremeSuiteControls.FlatEdit v2 
         Height          =   375
         Left            =   11385
         TabIndex        =   31
         Top             =   8400
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3096
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin MSComCtl2.DTPicker vperiodo 
         Height          =   345
         Left            =   825
         TabIndex        =   33
         Top             =   8430
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "MM/yyyy"
         Format          =   80871427
         CurrentDate     =   42479
      End
      Begin XtremeSuiteControls.PushButton PusBorrar 
         Height          =   375
         Left            =   10335
         TabIndex        =   34
         Top             =   8430
         Width           =   1005
         _Version        =   851968
         _ExtentX        =   1773
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Borrar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PusAgregar 
         Height          =   375
         Left            =   8055
         TabIndex        =   35
         Top             =   8460
         Width           =   1005
         _Version        =   851968
         _ExtentX        =   1773
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Agregar"
         UseVisualStyle  =   -1  'True
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   945
         Left            =   -69910
         TabIndex        =   39
         Top             =   7170
         Visible         =   0   'False
         Width           =   10635
         _ExtentX        =   18759
         _ExtentY        =   1667
         _Version        =   393216
         ForeColor       =   255
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin XtremeSuiteControls.FlatEdit vimporte 
         Height          =   345
         Left            =   3915
         TabIndex        =   40
         Top             =   8430
         Width           =   1485
         _Version        =   851968
         _ExtentX        =   2619
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin RichTextLib.RichTextBox vlibredeuda 
         Height          =   2145
         Left            =   7920
         TabIndex        =   44
         Top             =   360
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   3784
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         DisableNoScroll =   -1  'True
         TextRTF         =   $"frmDeudasServicios2.frx":059A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.GroupBox GroupBox4 
         Height          =   1080
         Left            =   180
         TabIndex        =   45
         Top             =   2250
         Width           =   1545
         _Version        =   851968
         _ExtentX        =   2725
         _ExtentY        =   1905
         _StockProps     =   79
         Caption         =   "Condición: "
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
         Begin XtremeSuiteControls.RadioButton rdBaja 
            Height          =   315
            Left            =   90
            TabIndex        =   46
            Top             =   210
            Width           =   1350
            _Version        =   851968
            _ExtentX        =   2381
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Dado de Baja"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rdActivo 
            Height          =   315
            Left            =   90
            TabIndex        =   47
            Top             =   450
            Width           =   1125
            _Version        =   851968
            _ExtentX        =   1984
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Activo"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rdTodosAB 
            Height          =   315
            Left            =   90
            TabIndex        =   56
            Top             =   720
            Width           =   1125
            _Version        =   851968
            _ExtentX        =   1984
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Todos"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla22 
         Height          =   2970
         Left            =   6390
         TabIndex        =   61
         Top             =   4815
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   5239
         _Version        =   393216
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin XtremeSuiteControls.PushButton PusExcel2 
         Height          =   405
         Left            =   11925
         TabIndex        =   62
         Top             =   7875
         Width           =   1065
         _Version        =   851968
         _ExtentX        =   1879
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Excel2"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   315
         Left            =   5565
         TabIndex        =   41
         Top             =   8430
         Width           =   825
         _Version        =   851968
         _ExtentX        =   1455
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Nro.Doc."
      End
      Begin XtremeSuiteControls.Label lblPeríodo 
         Height          =   315
         Left            =   135
         TabIndex        =   37
         Top             =   8430
         Width           =   795
         _Version        =   851968
         _ExtentX        =   1402
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Período :"
      End
      Begin XtremeSuiteControls.Label lblImporte 
         Height          =   315
         Left            =   2115
         TabIndex        =   36
         Top             =   8460
         Width           =   1725
         _Version        =   851968
         _ExtentX        =   3043
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Importe Deuda Original :"
      End
      Begin XtremeShortcutBar.ShortcutCaption vsaldo 
         Height          =   525
         Left            =   7470
         TabIndex        =   28
         Top             =   4110
         Width           =   3375
         _Version        =   851968
         _ExtentX        =   5953
         _ExtentY        =   926
         _StockProps     =   14
         Caption         =   "0.00"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         GradientColorLight=   128
         GradientColorDark=   255
         ForeColor       =   16777215
      End
      Begin XtremeSuiteControls.Label lblDeudaAcumulada 
         Height          =   255
         Left            =   5430
         TabIndex        =   27
         Top             =   4260
         Width           =   1725
         _Version        =   851968
         _ExtentX        =   3043
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Deuda Acumulada :"
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
      Begin XtremeSuiteControls.Label lblAgregarComentario 
         Height          =   255
         Left            =   90
         TabIndex        =   21
         Top             =   855
         Width           =   1515
         _Version        =   851968
         _ExtentX        =   2672
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Agregar comentario: "
      End
   End
End
Attribute VB_Name = "frmDeudasServicios2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vViene As String

Dim vvtitulo As String
Dim vetitulo As String


Private Sub Form_Load()

Me.Width = 13515
Me.Height = 9390

vperiodo.Value = Date - Day(Date) + 1

init

End Sub


Private Sub init()

Me.listamenu.ColWidth(0) = 2000
Me.listamenu.ColWidth(1) = 0

Me.vPeriodoDesde = Date
Me.vPeriodoHasta = Date

'--------------------------------------------------------------
Dim vsql, vsql2, v As String
On Error Resume Next
Dim rs4 As New ADODB.Recordset

Me.listamenu.Clear


vsql = "select nombre, consulta  from t_consultasmysql order by nombre"

Call rs4.Open(vsql, ConnDDBB, adOpenStatic, adLockPessimistic)

Set Me.listamenu.DataSource = rs4.DataSource

rs4.Close

If Err Then Exit Sub
End Sub

Function sqlTotal(vzona As String, pago As Boolean) As String
On Error Resume Next

sqlTotal = "select  Apellido, '' as Periodo, sum(ImporteTotal1) as ImporteTotal1, sum(Intereses) as Intereses,  sum(Deuda) as Deuda" + _
" from deudarural t Where t.zona = '" + vzona + "' Group By t.ID_Persona order by Apellido"

If Err Then Exit Function
End Function



Function sqlruralresumen(vzona As String) As String
On Error Resume Next

If vzona = "Rural" Then
        sqlruralresumen = "select  Apellido, '' as Periodo, sum(ImporteTotal1) as ImporteTotal1, sum(Intereses) as Intereses,  sum(Deuda) as Deuda" + _
        " from deudarural t Where t.zona = '" + vzona + "' Group By t.ID_Persona order by Apellido"
End If


If vzona = "Urbano" Then
        sqlruralresumen = "select  Apellido, '' as Periodo, sum(ImporteTotal1) as ImporteTotal1, sum(Intereses) as Intereses,  sum(Deuda) as Deuda" + _
        " from deudaubana t Where t.zona = '" + vzona + "' Group By t.ID_Persona order by Apellido"
End If


If Err Then Exit Function
End Function


Private Sub grilla2_Click()

Dim i, j  As Long

i = grilla2.Row

Me.vid.Text = grilla2.TextMatrix(i, 1)


j = grilla.Col

vv = grilla2.TextMatrix(i, j)


'If i = 0 Then ordenar_por (vv)


End Sub

Private Sub ordenar_por(vv As String)
'----------------




'----------------
End Sub


Private Sub listamenu_Click()
Dim vsql, vv, vw As String

Dim i, j As Integer

i = Me.listamenu.Row


vv = Me.listamenu.TextMatrix(i, 0)

vsql = "select width as c from t_consultasmysql where nombre = '" + Trim(vv) + "'"

vw = TraerDato2(vsql, "c", pathDBMySQL)


vsql = "select  consulta as c from t_consultasmysql where nombre = '" + Trim(vv) + "'"

vv2 = TraerDato2(vsql, "c", pathDBMySQL)


Call mostrargrilla(Me.grilla, TraerDato2(vsql, "c", pathDBMySQL), vw)


End Sub



Private Sub add_pago()
On Error Resume Next

Dim vcampos, vvalores As String
Dim vid_recibo_comercio_resumen As Long



' agrega pago en recibo_comercio_resumen

vcampos = "fecha_pago,id_comercios,nro_recibo_comercio"



vvalores = vvalores + "'" + strfechaMySQL(Now()) + "',"
vvalores = vvalores + Str(Me.vcontribuyente.Tag) + ","
vvalores = vvalores + "'" + vnrorecibo + "'"


vsql = "insert into recibo_comercio_resumen (" + vcampos + ") values (" + vvalores + ")"

Call EjecutarScript(vsql, ConnComunaDB)

vsql = "select max(id_recibo_comercio_resumen) as c from recibo_comercio_resumen "

vid_recibo_comercio_resumen = traerDatos2(vsql, "c", pathcomunadb)


' agrego pago en recibo_comercio_detalle
vcampos = "id_recibo_comercio_resumen,importe,periodo_anomes"


vvalores = ""
vvalores = vvalores + Str(vid_recibo_comercio_resumen) + ","
vvalores = vvalores + Str(Me.vimporte) + ","
vvalores = vvalores + "'" + Format(Me.vperiodo, "mm/yyyy") + "'"

vsql = "insert into recibo_comercio_detalle (" + vcampos + ") values (" + vvalores + ")"

Call EjecutarScript(vsql, ConnComunaDB)


If Err Then Exit Sub
End Sub


Function pathcomunadb() As String
    pathcomunadb = ConnComunaDB.ConnectionString
End Function
Private Sub PusAgregar_Click()
Dim vsql, vcampo, vValor, vzona As String


If UCase(LeerXml("Cliente")) = UCase("ComunaWw") Then
    
    add_pago
    
    Exit Sub
End If


If MsgBox("Confirma el ingreso ?", vbYesNo) = vbNo Then
    Exit Sub
End If

If rdRural.Value Then
    vzona = "Rural"
End If


If Me.rdUrbano Then
    vzona = "Urbano"
End If



vcampo = "ID_Contribuyente, Periodo, FechaGeneracion,FechaVencimiento1,ImporteTotal1,ImporteTotal2,Zona"

vvalores = ""
vvalores = vvalores + Str(Me.vcontribuyente.Tag) + ","
vvalores = vvalores + "'" + strfechaMySQL(Me.vperiodo.Value) + "',"
vvalores = vvalores + "'" + strfechaMySQL(Me.vperiodo.Value) + "',"
vvalores = vvalores + "'" + strfechaMySQL(Me.vperiodo.Value) + "',"
vvalores = vvalores + Str(Me.vimporte.Text) + ","
vvalores = vvalores + Str(Val(Me.vimporte.Text) * 1.03) + ","
vvalores = vvalores + "'" + vzona + "'"


vsql = " insert into t_recibos (" + vcampo + ") values (" + vvalores + ")"

Dim rec As New Recordset

Call rec.Open(vsql, ConnComunaDB, adOpenDynamic, adLockPessimistic)


'Exit Sub

'Call EjecutarScript(vsql, ConnComunaDB)

Call PusGenerarListado_Click

Me.vimporte.Text = ""

End Sub

Private Sub PusBorrar_Click()

Dim vsql As String


If Me.vid.Text = "" Then
    MsgBox "Debe seleccionar una cuota de la grilla "
    Exit Sub
End If


If MsgBox("Confirma la baja ?", vbYesNo) = vbNo Then
   If Not UCase(InputBox("Ingresar la palabra BORRAR ", "Confirmación")) = "BORRAR" Then Exit Sub
End If


Dim rec2 As New Recordset


If Me.RdComercio.Value Then

    vsql = "delete   from recibo_comercio_resumen " + _
    " Where id_recibo_comercio_resumen = " + Trim(Me.vid)

    Call rec2.Open(vsql, ConnComunaDB, adOpenDynamic, adLockPessimistic)
       
    vsql = "delete  from recibo_comercio_detalle  " + _
    " Where id_recibo_comercio_resumen = " + Trim(Me.vid)


    Call rec2.Open(vsql, ConnComunaDB, adOpenDynamic, adLockPessimistic)


End If

    
'Call EjecutarScript(vsql, ConnComunaDB)

Call PusGenerarListado_Click


Me.vimporte.Text = ""
Me.vid.Text = ""


End Sub

Private Sub PusCargarInformación_Click()
On Error Resume Next
Dim vsql As String
Dim vdatos As String
Dim vzona As Integer

Dim vmfrente, vmterreno As Double



vsql = "select t.ID_Zona  as c from  t_contribuyente t where t.ID_contribuyente = " + Str(Me.vcontribuyente.Tag)
vzona = Val(EsNulo(traerDatos2(vsql, "c", pathDBMySQLComuna)))

vsql = "select t.MetrosFrente as c from  t_contribuyente t where t.ID_contribuyente = " + Str(Me.vcontribuyente.Tag)
vmfrente = Val(EsNulo(traerDatos2(vsql, "c", pathDBMySQLComuna)))

vsql = "select t.Terreno  as c from  t_contribuyente t where t.ID_contribuyente = " + Str(Me.vcontribuyente.Tag)
vmterreno = Val(EsNulo(traerDatos2(vsql, "c", pathDBMySQLComuna)))

vdatos = ""


If vzona = 1 Then
    Me.rdUrbano.Value = True
    vdatos = vdatos + "Metros de Frente : [" + Str(vmfrente) + "]"
    vdatos = vdatos + " -- Metros Lateral : [" + Str(vmterreno) + "]"
Else
    Me.rdRural.Value = True
    vdatos = vdatos + "Hectareas : [" + Str(vmterreno) + "]"
End If

Me.vcomentario.Text = vdatos

Me.RdComercio.Value = True

If Err Then Exit Sub
End Sub

Private Sub PusExcel_Click()
Call grillaToExcel(Me.grilla2)
End Sub

Private Sub PusExcel2_Click()
Call grillaToExcel(Me.grilla22)
End Sub

Private Sub PusExportar_Click()
On Error Resume Next
Call grillaToExcel2(Me.g2e)

If Err Then Exit Sub
End Sub

Private Sub PusGenerarListado_Click()
On Error Resume Next

Dim vordenar As String
Dim vwhere As String

vetitulo = ""

vwhere = " 1=1 "

Dim vsql, vsql2, vvalue, vcampos  As String
Dim rec As New Recordset

vcampos = "c1,c2,c3,c4,c5"

Me.barra.Value = 0
vsql = ""

        vsql = "delete from Rural"
        Call EjecutarScript(vsql, PathDBListados)

vsql = ""


If Me.rdActivo Then
    vwhere = vwhere + " and (activo = 'S') "
    vetitulo = vetitulo + "Comercios Activos - "
End If

If Me.rdBaja Then
    vwhere = vwhere + " and (activo = 'N') "
    vetitulo = vetitulo + "Comercios Inactivo - "
End If


Dim vpdesde, vphasta As Long
 
If Me.rdPagos Then
   
    vpdesde = Val(Str(Year(Me.vPeriodoDesde)) + Format(Month(Me.vPeriodoDesde), "00"))
    vphasta = Val(Str(Year(Me.vPeriodoHasta)) + Format(Month(Me.vPeriodoHasta), "00"))
    vwhere = vwhere + "and  ( periodo2 >= " + Str(vpdesde) + " and periodo2 <= " + Str(vphasta) + " ) "

    vetitulo = vetitulo + "Períodos Pagos " + Str(vpdesde) + " ... " + Str(vphasta)
End If


If Me.rdImpagos Then
    vpdesde = Val(Str(Year(Me.vPeriodoDesde)) + Format(Month(Me.vPeriodoDesde), "00"))
    vphasta = Val(Str(Year(Me.vPeriodoHasta)) + Format(Month(Me.vPeriodoHasta), "00"))
    vwhere = vwhere + "and  not ( periodo2 >= " + Str(vpdesde) + " and periodo2 <= " + Str(vphasta) + " ) "

    vetitulo = vetitulo + "Períodos Impagos " + Str(vpdesde) + " ... " + Str(vphasta)
End If


If Me.rdResumen Then

    If Me.RdComercio Then
        ' vsql = sqlruralresumen("Rural")
        'vsql = sqlComercioDeuda
        
       
       
       If Me.rdcodigo.Value = True Then vordenar = " order by 1 "
       
       If Me.rdnombre.Value = True Then vordenar = " order by 2 "
        
        vsql = "select id, nombre_comercio as comercio, concat(mm,' / ',aaaa) as periodo_anomes, Ultimo_Importe as UltPago from deudacomerciototal where " + vwhere + vordenar
        
        If Me.rdPagos Or Me.rdImpagos Then
                Dim vvcampos As String
                vvcampos = " dc.id_comercios as id,  dc.nombre_comercio,  substring(max(dc.periodo_anomes),1,2) as mm,  substring(max(dc.periodo_anomes),4,4) as aaaa,sum(dc.importe) as Ultimo_Importe,0 as MesesImpagos,0 as Deuda"
                vsql = "select " + vvcampos + " From deudacomercio dc Where " + vwhere + " group by id_comercios order by nombre_comercio"
        End If
        
    End If

End If


If Me.rdDetalle Then
    vsql = " select * from deudaRural where " + vwhere + " ID_Persona = " + Me.vcontribuyente.Tag
    
    vetitulo = vetitulo + " Detallada - "
    
     If Me.rdUrbano Then
        vsql = " select * from deudaUrbana where ID_Persona = " + Me.vcontribuyente.Tag
    End If


    If Me.rdRural Then
        vsql = " select * from deudaRural where ID_Persona = " + Me.vcontribuyente.Tag
    End If

    If Me.RdComercio Then
       ' vsql = vsqldeudacomerciO
       'Dim vordenar As String
       
       If Me.rdcodigo.Value = True Then vordenar = " order by 1 "
       
       If Me.rdnombre.Value = True Then vordenar = " order by 2 "
       
        vsql = " select * from deudacomercio where " + vwhere + " and id_comercios = " + Me.vcontribuyente.Tag + vordenar
        vvtitulo = "Nro. recibo"
    End If
End If

Dim vsql3 As String

With rec

    Call .Open(vsql, ConnComunaDB, adOpenDynamic, adLockPessimistic)
    
    barra.Max = .RecordCount
    
    Dim vtotal As String
    
    vtotal = 0
    
    vsql3 = vsql
    
    Do Until .EOF
    
    
    If rdDetalle Then
        
        vvalue = vvalue + "'" + Left(.Fields("nombre_comercio"), 100) + "',"
        vvalue = vvalue + "'" + .Fields("periodo_anomes") + "',"
        vvalue = vvalue + "0,"
        vvalue = vvalue + "" + Str(.Fields("nro_recibo_comercio")) + ","
        vvalue = vvalue + Str(.Fields("importe"))
        
        
        If vvalue = "" Then Exit Do
        
        vtotal = vtotal + .Fields("Deuda")
    
    End If
    
    
    If rdResumen.Value Then
        
        vvalue = vvalue + "'"
        vvalue = vvalue + Left(.Fields("id"), 20) + "','"
        vvalue = vvalue + Left(.Fields("comercio"), 100) + " - "
        vvalue = vvalue + (.Fields(2)) + "',0,"
        vvalue = vvalue + Str(.Fields("UltPago")) + ",0"
       
        
        If vvalue = "" Then Exit Do
        
        ' vtotal = vtotal + .Fields("Deuda")
    
    End If
    
        barra.Value = barra.Value + 1
          
          vsql = "insert into Rural (" + vcampos + ") values (" + vvalue + ")"
        Call EjecutarScript(vsql, PathDBListados)
    
        vvalue = ""
        
        .MoveNext
        
    Loop
    
    'vsql = "insert into Rural (" + vcampos + ") values (" + vvalores + ")"
    'Call EjecutarScript(vsql, PathDBListados)
     
   ' totales totales totales
   If Me.RdComercio Then
        vsql2 = "select deuda as c from deudacomerciototal where id = " + Me.vcontribuyente.Tag
        vtotal = traerDatos2(vsql2, "c", ConnComunaDB.ConnectionString)
   End If
   
        Me.vsaldo.Caption = Format(vtotal, "###,###,##0.00")
    
        Dim vvw As String
            
        vvw = "0,3000,1000,1000,1000,1000,1000"
        
        Call mostrargrilla(Me.grilla2, vsql3, vvw)
        
        Dim vsql33 As String
        vsql33 = "select id,nombre_comercio,MesesImpagos,Ultimo_Importe from deudacomerciototal where  1=1 and  (activo = 'S')  order by 1"
         
         Call mostrargrilla(Me.grilla22, vsql33, vvw)
   
   
   'Set Me.g2e.Recordset = rec
   
  ' Set grilla2.DataSource = rec.DataSource
    
End With


Call reemplazarCamposEnTitulo


barra.Value = 0
Me.PusImprimir.Enabled = True


If Err Then Exit Sub
End Sub


Private Sub reemplazarCamposEnTitulo()
Dim vsql As String

vsql = "select * from comercios where id_comercios = " + Me.vcontribuyente.Tag

Dim v1, v2, v3, v4, v5, v As String

v1 = traerDatos2(vsql, "direccion", ConnComunaDB.ConnectionString)
v2 = traerDatos2(vsql, "cuit", ConnComunaDB.ConnectionString)
v3 = traerDatos2(vsql, "ing_brutos", ConnComunaDB.ConnectionString)
v4 = traerDatos2(vsql, "fecha_alta", ConnComunaDB.ConnectionString)

v5 = "COMERCIO ubicado en " + Trim$(v1) + ", CUIT: " + Trim$(v2) + ", I.B.:" + Trim$(v3) + ", F.Alta: " + (v4)


v = vlibredeuda.Text

v = Replace$(v, "COMERCIO", v5)

v = Replace$(v, "[campo5]", Str(Date))

v = Replace$(v, "[campo6]", vsaldo.Caption)



Me.vlibredeuda.Text = v


End Sub

Private Sub PushButton1_Click()
Call grillaToExcel(Me.grilla)
End Sub

Private Sub PusImprimir_Click()

Dim vmensaje As String

            Unload Mantenimiento
            Load Mantenimiento
             
            MsgBox "Prepare la impresora ...", vbInformation, "Mensaje ..."
            
            
            Unload Mantenimiento
            Load Mantenimiento
                 
            vmensaje = "Contribuyente: " + Trim(UCase(Me.vcontribuyente)) + "  -  "
            
            drWsurIntereses.Sections("TituloEmpresa").Controls("ecomentario").Caption = UCase(vcomentario)
            
            If vvtitulo = "" Then vvtitulo = "Ult.Imp - Meses Impagos "
            
            drWsurIntereses.Sections("sección2").Controls("lblCampo04").Caption = vvtitulo
            
            drWsurIntereses.Sections("sección2").Controls("etitulo").Caption = vetitulo
            
            If Me.chkTitulo.Value = 1 Then
            
                drWsurIntereses.Sections("TituloEmpresa").Controls("elibredeuda").Caption = Me.vlibredeuda.Text
            Else
            
                drWsurIntereses.Sections("TituloEmpresa").Controls("elibredeuda").Caption = ""
            
            End If
            
            
            
            drWsurIntereses.Show
            
 
End Sub

Private Sub PusLimpiar_Click()
Me.vcontribuyente.Tag = ""
Me.vcontribuyente.Text = ""
Me.vcomentario.Text = ""
Me.PusImprimir.Enabled = False
barra.Value = 0
End Sub

Private Sub PusPersonas_Click()

Me.vcontribuyente.Tag = ""
Me.vcontribuyente.Text = ""

Call fbuscarGrilla("comercios", "nombre_comercio", "id_comercios", Me.vcontribuyente.Name, Me, "nombre_comercio", True)
' ema:

vcp = "p"

End Sub

Private Sub vcomentario_Change()
Me.PusImprimir.Enabled = False
End Sub

Private Sub vcontribuyente_Change()
Call PusCargarInformación_Click

Me.PusImprimir.Enabled = False

'Call PusPersonas_Click
'Call PusImprimir_Click

End Sub

Private Sub vcontribuyente_Click()
Call PusPersonas_Click
End Sub



Function sqlComercioDeuda() As String
Dim v, v1, v2, v3, v4 As String


Dim vordenar As String


If Me.rdcodigo.Value Then vordenar = " order by fecha desc"

If Me.rdnombre.Value Then vordenar = " order by comercio asc"


' v1 = "select * from deudacomerciototal d "


v1 = " select " + _
" c.nombre_comercio as comercio, " + _
" max(r.fecha_pago) as Fecha, " + _
" (rd.importe) as UltPago, " + _
" r.id_comercios as id, " + _
" rd.periodo_anomes " + _
" from recibo_comercio_resumen r " + _
" Inner Join "

v2 = " (select * from recibo_comercio_detalle t order by t.id_recibo_comercio_resumen desc) rd " + _
" on r.id_recibo_comercio_resumen = rd.id_recibo_comercio_resumen " + _
" inner join comercios c on c.id_comercios = r.id_comercios Where c.activo = 'S' group by comercio " + vordenar

v3 = v1 + v2

sqlComercioDeuda = v3

End Function
