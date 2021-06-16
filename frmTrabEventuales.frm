VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmTrabEventuales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión de relacioens entre CONCEPTOS <-> CAJA <-> CTAS. CONTABLES"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   15870
   Begin XtremeSuiteControls.GroupBox fcomprobante 
      Height          =   2325
      Left            =   5670
      TabIndex        =   52
      Top             =   420
      Visible         =   0   'False
      Width           =   9675
      _Version        =   851968
      _ExtentX        =   17066
      _ExtentY        =   4101
      _StockProps     =   79
      Caption         =   "Comprobante:"
      ForeColor       =   0
      BackColor       =   -2147483644
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
      Begin XtremeSuiteControls.ProgressBar vpb 
         Height          =   375
         Left            =   240
         TabIndex        =   60
         Top             =   1680
         Width           =   4455
         _Version        =   851968
         _ExtentX        =   7858
         _ExtentY        =   661
         _StockProps     =   93
      End
      Begin XtremeSuiteControls.FlatEdit vobservaciones 
         Height          =   345
         Left            =   2010
         TabIndex        =   56
         Top             =   870
         Width           =   7365
         _Version        =   851968
         _ExtentX        =   12991
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin MSComCtl2.DTPicker vfecha 
         Height          =   375
         Left            =   2010
         TabIndex        =   55
         Top             =   420
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   661
         _Version        =   393216
         Format          =   185794561
         CurrentDate     =   41882
      End
      Begin XtremeSuiteControls.PushButton PusComenzarA 
         Height          =   375
         Left            =   6840
         TabIndex        =   57
         Top             =   1350
         Width           =   2535
         _Version        =   851968
         _ExtentX        =   4471
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Comenzar a imprimir"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmTrabEventuales.frx":0000
      End
      Begin XtremeSuiteControls.Label lblObservacionesGrales 
         Height          =   345
         Left            =   360
         TabIndex        =   54
         Top             =   810
         Width           =   1665
         _Version        =   851968
         _ExtentX        =   2937
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Observaciones grales: "
         BackColor       =   -2147483644
      End
      Begin XtremeSuiteControls.Label lblFechaDe 
         Height          =   345
         Left            =   420
         TabIndex        =   53
         Top             =   420
         Width           =   1575
         _Version        =   851968
         _ExtentX        =   2778
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Fecha de los recibos: "
         BackColor       =   -2147483644
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox4 
      Height          =   765
      Left            =   30
      TabIndex        =   45
      Top             =   7560
      Width           =   15735
      _Version        =   851968
      _ExtentX        =   27755
      _ExtentY        =   1349
      _StockProps     =   79
      Caption         =   "Modificaciones de todos los datos seleccionados en la grilla: "
      UseVisualStyle  =   -1  'True
      Begin MSComCtl2.DTPicker vmfd 
         Height          =   375
         Left            =   6480
         TabIndex        =   46
         Top             =   270
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   185794561
         CurrentDate     =   41887
      End
      Begin MSComCtl2.DTPicker vmfh 
         Height          =   375
         Left            =   10530
         TabIndex        =   47
         Top             =   270
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   185794561
         CurrentDate     =   41887
      End
      Begin XtremeSuiteControls.PushButton PusEjecutarModificación 
         Height          =   345
         Left            =   13320
         TabIndex        =   50
         Top             =   270
         Width           =   2235
         _Version        =   851968
         _ExtentX        =   3942
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Ejecutar modificación"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmTrabEventuales.frx":059A
      End
      Begin XtremeSuiteControls.Label Label7 
         Height          =   255
         Left            =   8670
         TabIndex        =   49
         Top             =   330
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Último día de trabajo:"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   255
         Left            =   4560
         TabIndex        =   48
         Top             =   330
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Día el inicio del trabajo:"
         Alignment       =   1
      End
   End
   Begin MSComctlLib.StatusBar sbTotales 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   37
      Top             =   8340
      Width           =   15870
      _ExtentX        =   27993
      _ExtentY        =   820
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.TabControl tab 
      Height          =   7095
      Left            =   90
      TabIndex        =   12
      Top             =   420
      Width           =   15765
      _Version        =   851968
      _ExtentX        =   27808
      _ExtentY        =   12515
      _StockProps     =   68
      ItemCount       =   2
      Item(0).Caption =   "Altas"
      Item(0).ControlCount=   14
      Item(0).Control(0)=   "GroupBox3"
      Item(0).Control(1)=   "GroupBox2"
      Item(0).Control(2)=   "lblCuentaContable"
      Item(0).Control(3)=   "Label3"
      Item(0).Control(4)=   "PusBuscarCliente"
      Item(0).Control(5)=   "PushButton4"
      Item(0).Control(6)=   "vproveedor"
      Item(0).Control(7)=   "vttrabajo"
      Item(0).Control(8)=   "GroRemuneración"
      Item(0).Control(9)=   "GroPeríodos"
      Item(0).Control(10)=   "GroDatosAdicciondos"
      Item(0).Control(11)=   "vcodproveedor"
      Item(0).Control(12)=   "vestado"
      Item(0).Control(13)=   "lblEstado"
      Item(1).Caption =   "Datos"
      Item(1).ControlCount=   14
      Item(1).Control(0)=   "grilla"
      Item(1).Control(1)=   "lblBuscar"
      Item(1).Control(2)=   "vbucar"
      Item(1).Control(3)=   "GroupBox1"
      Item(1).Control(4)=   "PbAcciones3"
      Item(1).Control(5)=   "PbAcciones2"
      Item(1).Control(6)=   "PushButton3"
      Item(1).Control(7)=   "vfbdesde"
      Item(1).Control(8)=   "vfbhasta"
      Item(1).Control(9)=   "Label4"
      Item(1).Control(10)=   "Label5"
      Item(1).Control(11)=   "PusBuscar"
      Item(1).Control(12)=   "PusExportar"
      Item(1).Control(13)=   "PusVolver"
      Begin XtremeSuiteControls.ComboBox vestado 
         Height          =   315
         Left            =   3780
         TabIndex        =   9
         Top             =   4590
         Width           =   9855
         _Version        =   851968
         _ExtentX        =   17383
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.GroupBox GroDatosAdicciondos 
         Height          =   735
         Left            =   1500
         TabIndex        =   34
         Top             =   5130
         Width           =   12165
         _Version        =   851968
         _ExtentX        =   21458
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Datos adicciondos: "
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit vcomentarios 
            Height          =   345
            Left            =   2880
            TabIndex        =   10
            Top             =   300
            Width           =   9075
            _Version        =   851968
            _ExtentX        =   16007
            _ExtentY        =   609
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.Label lblComentariosY 
            Height          =   255
            Left            =   0
            TabIndex        =   35
            Top             =   360
            Width           =   2445
            _Version        =   851968
            _ExtentX        =   4313
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Comentarios y observaciones: "
            Alignment       =   1
         End
      End
      Begin XtremeSuiteControls.GroupBox GroPeríodos 
         Height          =   645
         Left            =   2310
         TabIndex        =   31
         Top             =   2670
         Width           =   11325
         _Version        =   851968
         _ExtentX        =   19976
         _ExtentY        =   1138
         _StockProps     =   79
         Caption         =   "Períodos: "
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
         Begin MSComCtl2.DTPicker vfdesde 
            Height          =   375
            Left            =   2790
            TabIndex        =   3
            Top             =   210
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   661
            _Version        =   393216
            Format          =   185794561
            CurrentDate     =   41882
         End
         Begin MSComCtl2.DTPicker vfhasta 
            Height          =   375
            Left            =   7440
            TabIndex        =   4
            Top             =   210
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   661
            _Version        =   393216
            Format          =   185794561
            CurrentDate     =   41882
         End
         Begin XtremeSuiteControls.Label lblDíaDe 
            Height          =   255
            Left            =   5490
            TabIndex        =   33
            Top             =   270
            Width           =   1815
            _Version        =   851968
            _ExtentX        =   3201
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Último día de trabajo:"
            Alignment       =   1
         End
         Begin XtremeSuiteControls.Label lblDíaEl 
            Height          =   255
            Left            =   870
            TabIndex        =   32
            Top             =   270
            Width           =   1815
            _Version        =   851968
            _ExtentX        =   3201
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Día el inicio del trabajo:"
            Alignment       =   1
         End
      End
      Begin XtremeSuiteControls.GroupBox GroRemuneración 
         Height          =   1005
         Left            =   1140
         TabIndex        =   26
         Top             =   3450
         Width           =   12915
         _Version        =   851968
         _ExtentX        =   22781
         _ExtentY        =   1773
         _StockProps     =   79
         Caption         =   "Remuneración:"
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit vHoras 
            Height          =   315
            Left            =   2850
            TabIndex        =   5
            Top             =   210
            Width           =   2085
            _Version        =   851968
            _ExtentX        =   3678
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vhorasExtras 
            Height          =   315
            Left            =   2850
            TabIndex        =   7
            Top             =   570
            Width           =   2085
            _Version        =   851968
            _ExtentX        =   3678
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vHorasImporte 
            Height          =   315
            Left            =   6930
            TabIndex        =   6
            Top             =   240
            Width           =   2085
            _Version        =   851968
            _ExtentX        =   3678
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vhorasExtrasImporte 
            Height          =   315
            Left            =   6960
            TabIndex        =   8
            Top             =   600
            Width           =   2085
            _Version        =   851968
            _ExtentX        =   3678
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vtotalSemanal 
            Height          =   315
            Left            =   10770
            TabIndex        =   58
            Top             =   240
            Width           =   2085
            _Version        =   851968
            _ExtentX        =   3678
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.Label Label8 
            Height          =   255
            Left            =   9360
            TabIndex        =   59
            Top             =   240
            Width           =   1275
            _Version        =   851968
            _ExtentX        =   2249
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Total Semanal:"
            Alignment       =   1
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Left            =   5070
            TabIndex        =   30
            Top             =   600
            Width           =   1785
            _Version        =   851968
            _ExtentX        =   3149
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Importe Horas Extras:"
            ForeColor       =   255
            Alignment       =   1
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Left            =   5160
            TabIndex        =   29
            Top             =   240
            Width           =   1725
            _Version        =   851968
            _ExtentX        =   3043
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Importe por Horas: "
            Alignment       =   1
         End
         Begin XtremeSuiteControls.Label lblHorasExtras 
            Height          =   255
            Left            =   1170
            TabIndex        =   28
            Top             =   570
            Width           =   1545
            _Version        =   851968
            _ExtentX        =   2725
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Horas extras Total:"
            ForeColor       =   255
            Alignment       =   1
         End
         Begin XtremeSuiteControls.Label lblComentarios 
            Height          =   255
            Left            =   1650
            TabIndex        =   27
            Top             =   240
            Width           =   1125
            _Version        =   851968
            _ExtentX        =   1984
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Horas por días: "
            Alignment       =   1
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   195
         Left            =   -69940
         TabIndex        =   24
         Top             =   780
         Visible         =   0   'False
         Width           =   15675
         _Version        =   851968
         _ExtentX        =   27649
         _ExtentY        =   344
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
      End
      Begin XtremeSuiteControls.FlatEdit vbucar 
         Height          =   315
         Left            =   -69190
         TabIndex        =   18
         Top             =   1140
         Visible         =   0   'False
         Width           =   6465
         _Version        =   851968
         _ExtentX        =   11404
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla 
         Height          =   5295
         Left            =   -69910
         TabIndex        =   16
         Top             =   1590
         Visible         =   0   'False
         Width           =   15585
         _ExtentX        =   27490
         _ExtentY        =   9340
         _Version        =   393216
         FixedCols       =   0
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).BandIndent=   9
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   135
         Left            =   30
         TabIndex        =   13
         Top             =   780
         Width           =   15705
         _Version        =   851968
         _ExtentX        =   27702
         _ExtentY        =   238
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   405
         Left            =   60
         TabIndex        =   14
         Top             =   420
         Width           =   1725
         _Version        =   851968
         _ExtentX        =   3043
         _ExtentY        =   714
         _StockProps     =   79
         Appearance      =   2
         BorderStyle     =   2
         Begin XtremeSuiteControls.PushButton PbAcciones 
            Height          =   345
            Index           =   6
            Left            =   30
            TabIndex        =   11
            Top             =   0
            Width           =   1485
            _Version        =   851968
            _ExtentX        =   2619
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "Guardar <F2>"
            UseVisualStyle  =   -1  'True
            Picture         =   "frmTrabEventuales.frx":0B34
         End
      End
      Begin XtremeSuiteControls.PushButton PusBuscarCliente 
         Height          =   285
         Left            =   4200
         TabIndex        =   0
         Top             =   1800
         Width           =   825
         _Version        =   851968
         _ExtentX        =   1455
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "<F2>"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmTrabEventuales.frx":0F01
      End
      Begin XtremeSuiteControls.FlatEdit vproveedor 
         Height          =   315
         Left            =   5100
         TabIndex        =   2
         Top             =   1800
         Width           =   5865
         _Version        =   851968
         _ExtentX        =   10345
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton PbAcciones3 
         Height          =   345
         Left            =   -67510
         TabIndex        =   19
         Top             =   420
         Visible         =   0   'False
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Borrar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmTrabEventuales.frx":149B
      End
      Begin XtremeSuiteControls.PushButton PbAcciones2 
         Height          =   345
         Left            =   -66400
         TabIndex        =   20
         Top             =   420
         Visible         =   0   'False
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Modificar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmTrabEventuales.frx":1A35
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   285
         Left            =   4200
         TabIndex        =   1
         Top             =   2280
         Width           =   825
         _Version        =   851968
         _ExtentX        =   1455
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "<F3>"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmTrabEventuales.frx":1FCF
      End
      Begin XtremeSuiteControls.FlatEdit vttrabajo 
         Height          =   315
         Left            =   5100
         TabIndex        =   21
         Top             =   2250
         Width           =   5865
         _Version        =   851968
         _ExtentX        =   10345
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vcodproveedor 
         Height          =   315
         Left            =   11100
         TabIndex        =   23
         Top             =   1800
         Width           =   2475
         _Version        =   851968
         _ExtentX        =   4366
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   345
         Left            =   -65290
         TabIndex        =   25
         Top             =   420
         Visible         =   0   'False
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmTrabEventuales.frx":2569
      End
      Begin MSComCtl2.DTPicker vfbdesde 
         Height          =   375
         Left            =   -60730
         TabIndex        =   38
         Top             =   1110
         Visible         =   0   'False
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   185794561
         CurrentDate     =   41887
      End
      Begin MSComCtl2.DTPicker vfbhasta 
         Height          =   375
         Left            =   -56680
         TabIndex        =   39
         Top             =   1110
         Visible         =   0   'False
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   185794561
         CurrentDate     =   41887
      End
      Begin XtremeSuiteControls.PushButton PusBuscar 
         Height          =   345
         Left            =   -55480
         TabIndex        =   42
         Top             =   420
         Visible         =   0   'False
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Buscar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmTrabEventuales.frx":2B03
      End
      Begin XtremeSuiteControls.PushButton PusExportar 
         Height          =   345
         Left            =   -63490
         TabIndex        =   43
         Top             =   420
         Visible         =   0   'False
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Exportar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmTrabEventuales.frx":309D
      End
      Begin XtremeSuiteControls.PushButton PusVolver 
         Height          =   345
         Left            =   -69850
         TabIndex        =   44
         Top             =   420
         Visible         =   0   'False
         Width           =   1185
         _Version        =   851968
         _ExtentX        =   2090
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Volver"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmTrabEventuales.frx":3637
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   255
         Left            =   -62650
         TabIndex        =   41
         Top             =   1170
         Visible         =   0   'False
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Día el inicio del trabajo:"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Left            =   -58540
         TabIndex        =   40
         Top             =   1170
         Visible         =   0   'False
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Último día de trabajo:"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblEstado 
         Height          =   255
         Left            =   2460
         TabIndex        =   36
         Top             =   4590
         Width           =   1125
         _Version        =   851968
         _ExtentX        =   1984
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Estado:"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Left            =   300
         TabIndex        =   22
         Top             =   2250
         Width           =   3585
         _Version        =   851968
         _ExtentX        =   6324
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ingrese el tipo de trabajo que debe realizar:"
         Alignment       =   1
      End
      Begin VB.Label lblBuscar 
         Caption         =   "Buscar:"
         Height          =   285
         Left            =   -69850
         TabIndex        =   17
         Top             =   1200
         Visible         =   0   'False
         Width           =   705
      End
      Begin XtremeSuiteControls.Label lblCuentaContable 
         Height          =   255
         Left            =   180
         TabIndex        =   15
         Top             =   1770
         Width           =   3675
         _Version        =   851968
         _ExtentX        =   6482
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ingrese nombre del Evenual: "
         Alignment       =   1
      End
   End
   Begin XtremeSuiteControls.PushButton PusImprimirComprobantes 
      Height          =   375
      Left            =   13260
      TabIndex        =   51
      Top             =   0
      Width           =   2535
      _Version        =   851968
      _ExtentX        =   4471
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Comprobantes Individuales"
      UseVisualStyle  =   -1  'True
      Picture         =   "frmTrabEventuales.frx":3BD1
   End
End
Attribute VB_Name = "frmTrabEventuales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vid As Long
Dim vidArticulos As Long
Dim vulinea, vrow As Integer
Dim vidcampo As String
Dim vtabla As String
Dim rsPresupuesto, rsPresupuesto2 As New ADODB.Recordset
Dim mvc(11, 4) As String
Dim vModo, vgsql, vcondi  As String
Public vViene As String


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


Private Sub Form_Load()
    init
End Sub

Private Sub init()

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2 - 1000

vModo = "nuevo"

Call configurarGrid

Me.tab.SelectedItem = 0
If Not vViene = "" Then Me.tab.SelectedItem = 0


'Me.vcontexto.AddItem "Cuenta auxiliar: " + traerDatos2("select * from ctaauxiliar", "codigocta", pathDBMySQL)

vtabla = "t_trabeventuales"
vidcampo = "idtrabeventuales"

'Me.vnombrecta.Tag = ""
'Me.vnombrecta.Text = ""

vestado.AddItem "Suspendido"
vestado.AddItem "Activo"


vestado.Text = "Activo"

vfdesde.Value = Date
vfhasta.Value = Date + 5

vfbdesde.Value = Date
vfbhasta.Value = Date


Call initmvc

Call vbucar_Change

Me.vmfd.Value = Date

'CentrarFormulario (Me)
End Sub


Private Sub initmvc()

mvc(1, 1) = "vproveedor"
mvc(1, 3) = "tag" ' tiene tag
mvc(1, 4) = "idproveedores" ' tiene tag

mvc(2, 1) = "vfdesde"
mvc(3, 1) = "vfhasta"
mvc(4, 1) = "vHoras"
mvc(5, 1) = "vhorasExtras"
mvc(6, 1) = "vHorasImporte"
mvc(7, 1) = "vhorasExtrasImporte"
mvc(8, 1) = "vestado"
mvc(9, 1) = "vcomentarios"
mvc(10, 1) = "vttrabajo"
mvc(10, 2) = "tipomovimiento"
mvc(10, 4) = "idtipomovimientos"

mvc(1, 2) = "idproveedores"
mvc(2, 2) = "fdesde"
mvc(3, 2) = "fhasta"
mvc(4, 2) = "horas"
mvc(5, 2) = "horasextras"
mvc(6, 2) = "valorhora"
mvc(7, 2) = "valorhoraextra"
mvc(8, 2) = "estado"
mvc(9, 2) = "comentarios"

mvc(11, 1) = "vcodProveedor"
mvc(11, 2) = "codPersona"






End Sub

Private Sub lblImporte_Click()

End Sub
Private Sub LimpiarCampos()

Me.vproveedor.Tag = ""
Me.vproveedor.Text = ""

vfdesde.Value = Date
vfhasta.Value = Date
vHoras = ""
vhorasExtras = ""
vHorasImporte = ""
vhorasExtrasImporte = ""
Me.vestado = ""
Me.vcomentarios = ""

Me.vttrabajo.Tag = ""
Me.vttrabajo.Text = ""

Call init

End Sub


Private Sub pintar(ByVal i As Integer, g As MSHFlexGrid)
On Error Resume Next
Dim j, k, kk As Integer

k = g.Row
kk = g.Col

g.Row = i

For j = 1 To g.Cols - 1
    g.Col = j
    g.CellBackColor = vbGreen
Next

g.Row = k
g.Col = kk
If Err Then Exit Sub
End Sub


Private Sub grilla_Click()


vid = grilla.TextMatrix(grilla.Row, 0)
vidArticulos = grilla.TextMatrix(grilla.Row, 1)
vrow = grilla.Row
Call pintar(grilla.Row, Me.grilla)
Call despintar(vulinea, Me.grilla)

grilla.CellBackColor = vbRed

vulinea = grilla.Row


End Sub



Private Sub despintar(ByVal i As Integer, g As MSHFlexGrid)
On Error Resume Next

Dim j, k, kk As Integer
k = g.Row
kk = g.Col
If i = 0 Then Exit Sub
g.Row = i

For j = 1 To g.Cols - 1
    g.Col = j
    g.CellBackColor = vbWhite
Next

g.Row = k
g.Col = kk

If Err Then Exit Sub
End Sub


Private Sub grilla_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
'MsgBox "1"
End Sub

Private Sub grilla_DblClick()
vrow = grilla.Row
End Sub

'Private Sub g_Click()
'vidArticulos = g.TextMatrix(g.Row, 13)
'
'Call pintar(g.Row, grilla)
'
'Call despintar(vulinea, grilla)
'
'grilla.CellBackColor = vbRed
'
'vulinea = g.Row
'
'End Sub

Private Sub PbAcciones_Click(Index As Integer)
On Error Resume Next
Dim vsql, vcampos, vvalores As String

If Not ValidarCampos Then Exit Sub

If vModo = "modificando" Then
    Call PbAcciones3_Click
End If

vcampos = "idproveedores,fdesde,fhasta,horas,horasextras,valorhora,valorhoraextra,estado,comentarios,idTipoMovimientos"


vvalores = Me.vproveedor.Tag + ",'" + strfechaMySQL(vfdesde) + "','" + strfechaMySQL(vfhasta) + "'," + Str(Val(Me.vHoras.Text)) + "," + Str(Val(Me.vhorasExtras.Text)) + "," + _
Str(Val(Me.vHorasImporte.Text)) + "," + Str(Val(Me.vhorasExtrasImporte.Text)) + ",'" + Me.vestado.Text + "','" + Me.vcomentarios.Text + "'," + Str(Val(Me.vttrabajo.Tag))


vsql = "insert into t_trabeventuales (" + vcampos + ") values (" + vvalores + ")"

Call EjecutarScript(vsql, pathDBMySQL)

Call LimpiarCampos

Call vbucar_Change

If Err Then Exit Sub
End Sub
Function ValidarCampos() As Boolean
ValidarCampos = True
Exit Function

End Function

Private Sub PbAcciones2_Click()
cargarDatos (vid)
Me.tab.SelectedItem = 0
vModo = "modificando"
End Sub

Private Sub cargarDatos(vid As Long)
On Error Resume Next

Dim rec As Recordset
Dim vsql As String
Dim i As Integer
Dim vn1, vn2, vn3 As String
i = 1

vsql = "select * from eventuales where idtrabeventuales=" + Str(vid)

  With rsPresupuesto2
        Call .Open(vsql, ConnDDBB, adOpenStatic, adLockPessimistic)
  
        For i = 1 To UBound(mvc, 1)
            
              If Not EsNulo(mvc(i, 4)) = "" Then
                Me.Controls(EsNulo(mvc(i, 1))).Tag = .Fields(EsNulo(mvc(i, 4)))
            End If
                If Left(EsNulo(mvc(i, 1)), 2) = "vf" Or Left(EsNulo(mvc(i, 1)), 1) = "f" Then
                    Me.Controls(EsNulo(mvc(i, 1))).Value = .Fields(mvc(i, 2))
                Else
                    Me.Controls(EsNulo(mvc(i, 1))).Text = .Fields(mvc(i, 2))
                End If
            
        Next
        
        .Close
  End With
    
If Err Then Exit Sub
End Sub

Private Sub PbAcciones3_Click()
On Error Resume Next
Dim vsql As String
Dim vid As Integer

If Not MsgBox("Confirma la opración ? ", vbYesNo, "Borrar / Modificar") = vbYes Then
    Exit Sub
End If


vid = grilla.TextMatrix(grilla.Row, 0)

vsql = "delete from " + vtabla + " where " + vidcampo + " = " + Str(vid) + ""

Call EjecutarScript(vsql, pathDBMySQL)

Call vbucar_Change

If Err < 0 Then
    MsgBox "No se pudo borrar la linea", vbCritical
    Exit Sub
End If

End Sub

Private Sub PushBuscarCliente_Click()
End Sub



Private Sub PusBuscar_Click()
Call buscarFecha
End Sub

Private Sub PusBuscarCliente_Click()
Call fbuscarGrilla("(select * from proveedores where tipoproveedor='Eventuales') as p", "Nombre", "idProveedores", Me.vproveedor.Name, Me, , False)  ' ema:
End Sub

Private Sub PushButton1_Click()
End Sub
Private Sub PushButton2_Click()
End Sub


Private Sub PusComenzarA_Click()
Dim i As Integer
Me.vpb.Min = 0
Me.vpb.Max = grilla.Rows - 1
Me.vpb.Value = 0


For i = 1 To grilla.Rows - 1
    Me.vpb.Value = Me.vpb + 1
    vid = grilla.TextMatrix(i, 1)
    llenarDrRecibo (vid)
  
    Call drRecibo.PrintReport(False, rptRangeAllPages)
    Unload drRecibo.object
Next

Me.fcomprobante.Visible = False
End Sub

Private Sub PusEjecutarModificación_Click()
Dim i, vr As Integer
Dim vid As Long
Dim vsql1, vsql2  As String


For i = 1 To grilla.Rows - 1
    vr = grilla.Row
    vid = grilla.TextMatrix(vr, 1)
    
    vsql1 = " fdesde = '" + strfechaMySQL(vmfd.Value) + "', fhasta='" + strfechaMySQL(vmfh.Value) + "'"
    vsql2 = "update t_trabeventuales set " + vsql1 + " where idtrabeventuales=" + Str(vid)
    
    Call EjecutarScript(vsql2, pathDBMySQL)
    

Next

Call vbucar_Change

End Sub

Private Sub PusExportar_Click()
Call generarExcel("archivo.xls", Me.grilla)
End Sub

Private Sub PushButton3_Click()
  
        Unload Mantenimiento
        Load Mantenimiento
        
        With Mantenimiento.rsTrabEventuales
            
            If .State = 1 Then .Close
             
            'If rsPresupuesto.State = 1 Then .Close
            .Source = vgsql
           'Set .Source = rsPresupuesto.Source         ' Emma: le pasa lo que está en bfactura al conector del datareport
            
            If .State = 0 Then .Open
            .Close
            .Open
        
        End With
        
        With drTrabEventuales
            .Sections("titulos").Controls("efiltro").Caption = vcondi
            .Show
        End With
        
End Sub

Private Sub PushButton4_Click()
    Call fbuscarGrilla("tipoMovimientos", "TipoMovimiento", "idTipoMovimientos", Me.vttrabajo.Name, Me, , False)   ' ema:
End Sub

Private Sub PusImprimirComprobantes_Click()
Me.fcomprobante.Visible = True
vfecha.Value = Date
Me.vobservaciones.Text = ""
End Sub

Private Sub imprimirComprobante(vid As Long)

End Sub


Private Sub PusVolver_Click()
If vViene = "frmIngresosEgresos" Then
 Call frmIngresosEgresos.cargarEventuales
End If
End Sub

Private Sub tab_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Call vbucar_Change
End Sub

Private Sub vbanco_Change()
Dim vsql As String

'vsql = "select idbancos from bancos t where t.idcuentas = " + Me.vcuenta.Tag
'Me.vCodigoCuenta = traerDatos2(vsql, "codigocuenta", pathDBMySQL)
  
 'Me.vcodigobanco.Text = Me.vbanco.Tag
  
End Sub

Private Sub buscarFecha()
On Error Resume Next
Dim vsql, vwhere, vcampo, vsqlFecha   As String

'vCampo = "`conceptos2`.`ref`,  `conceptos2`.`descripcion`,  `conceptos2`.`idbancos`,`bancos`.`Descripcion`,conceptos2.debito as D,conceptos2.credito as C,`cuentas`.`CodigoCuenta`,`cuentas`.`Cuenta`, conceptos2.debe as DD, conceptos2.haber as H,`conceptos2`.`comentarios`,`conceptos2`.`idconceptos`,`conceptos2`.`idcuentas`,c2.codigocuenta,c2.cuenta"

vsqlFecha = ""

If Me.vfbdesde.CheckBox Then
    vsqlFecha = vsqlFecha + " and ( fdesde >= '" + strfechaMySQL(vfbdesde) + "')"
End If

If Me.vfbhasta.CheckBox Then
    vsqlFecha = vsqlFecha + " and ( fhasta <= '" + strfechaMySQL(vfbhasta) + "')"
End If


Dim rsPresupuesto As New ADODB.Recordset

vwhere = " where ( nombre like '%" + vbucar + "%' or codPersona like '%" + vbucar + "%' or codTrabajo like '%" + vbucar + "%') " + vsqlFecha
vsql = "select *  from eventuales" + vwhere
    
vgsql = vsql
vcondi = vwhere


    With rsPresupuesto
        Call .Open(vsql, ConnDDBB, adOpenStatic, adLockPessimistic)
        Set grilla.DataSource = .DataSource
    End With
End Sub


Private Sub vbucar_Change()
On Error Resume Next
Dim vsql, vwhere, vcampo, vsqlFecha   As String


'vCampo = "`conceptos2`.`ref`,  `conceptos2`.`descripcion`,  `conceptos2`.`idbancos`,`bancos`.`Descripcion`,conceptos2.debito as D,conceptos2.credito as C,`cuentas`.`CodigoCuenta`,`cuentas`.`Cuenta`, conceptos2.debe as DD, conceptos2.haber as H,`conceptos2`.`comentarios`,`conceptos2`.`idconceptos`,`conceptos2`.`idcuentas`,c2.codigocuenta,c2.cuenta"

vsqlFecha = ""

If Me.vfbdesde.CheckBox Then
    vsqlFecha = vsqlFecha + " and ( fdesde >= '" + strfechaMySQL(vfbdesde) + "')"
End If

If Me.vfbhasta.CheckBox Then
    vsqlFecha = vsqlFecha + " and ( fhasta <= '" + strfechaMySQL(vfbhasta) + "')"
End If


Dim rsPresupuesto As New ADODB.Recordset

vwhere = " where ( nombre like '%" + vbucar + "%' or codPersona like '%" + vbucar + "%' or codTrabajo like '%" + vbucar + "%') "
vsql = "select *  from eventuales " + vwhere
    
    
vgsql = vsql
vcondi = vwhere
    
    With rsPresupuesto
        Call .Open(vsql, ConnDDBB, adOpenStatic, adLockPessimistic)
        Set grilla.DataSource = .DataSource
    End With
        
End Sub

Private Sub vcodigoCta_Change()
'Me.vimporte.SetFocus
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

Me.grilla.Cols = 17

Me.grilla.ColWidth(0) = 0 'idproveedores
Me.grilla.ColWidth(1) = 1000
Me.grilla.ColWidth(2) = 3000

Me.grilla.ColWidth(3) = 1000
Me.grilla.ColWidth(4) = 1000

Me.grilla.ColWidth(5) = 300
Me.grilla.ColWidth(6) = 300

Me.grilla.ColWidth(7) = 300
Me.grilla.ColWidth(8) = 300


Me.grilla.ColWidth(9) = 600
Me.grilla.ColWidth(10) = 2000
Me.grilla.ColWidth(11) = 3300
Me.grilla.ColWidth(12) = 1000

Me.grilla.ColWidth(13) = 0
Me.grilla.ColWidth(14) = 1000 ' importe


Me.grilla.ColWidth(15) = 0 ' importe
Me.grilla.ColWidth(16) = 0
Me.grilla.ColWidth(17) = 0

End Sub

Private Sub vnombrecta_Change()
Dim vsql As String

'vsql = "select codigocuenta from cuentas t where t.idcuentas = " + Me.vnombrecta.Tag
'Me.vcodigoCta = traerDatos2(vsql, "codigocuenta", pathDBMySQL)
    
    
End Sub

Private Sub vcuenta_Change()
'Dim vsql As String

'vsql = "select codigocuenta from cuentas t where t.idcuentas = " + Me.vcuenta.Tag
'Me.vcodigoCta.Text = traerDatos2(vsql, "codigocuenta", pathDBMySQL)
        

End Sub

Private Sub vcuenta2_Change()
'Dim vsql As String
'vsql = "select codigocuenta from cuentas t where t.idcuentas = " + Me.vcuenta.Tag
'Me.vcodigoCta2.Text = traerDatos2(vsql, "codigocuenta", pathDBMySQL)
End Sub

Private Sub vfbdesde_Change()
Call buscarFecha
End Sub

Private Sub vfbhasta_Change()
Call buscarFecha
End Sub

Private Sub vfdesde_Change()
Me.vfhasta.Value = Me.vfdesde.Value + 5
End Sub

Private Sub vHoras_Change()
Call CalcularTotales
End Sub

Private Sub CalcularTotales()
Dim vthoras, vt, vte, vtt As Double
Dim vdias, vdiase As Integer

vdias = Me.vfhasta - Me.vfdesde

vt = vdias * Val(Me.vHoras) * Val(Me.vHorasImporte)
vte = Val(Me.vhorasExtras) * Val(Me.vhorasExtrasImporte)

vtt = vt + vte

Me.sbTotales.Panels.Item(1).Text = "Remuneración Total:"
Me.sbTotales.Panels.Item(2).Text = Format(vtt, "$ ###,###,###,##0.00")
End Sub

Private Sub vhorasExtras_Change()
Call CalcularTotales
End Sub

Private Sub vhorasExtrasImporte_Change()
Call CalcularTotales
End Sub

Private Sub vHorasImporte_Change()
Call CalcularTotales
End Sub

Private Sub vmfd_Change()
vmfh.Value = vmfd.Value + 5
End Sub

Private Sub vproveedor_Change()
On Error Resume Next
Dim vsql As String

vsql = "select * from proveedores where idproveedores = " + Me.vproveedor.Tag
Me.vcodproveedor = TraerDato2(vsql, "codigo", pathDBMySQL)

If Err Then Exit Sub
End Sub


Private Sub llenarDrRecibo(vid As Long)
On Error Resume Next

Dim vvsaldo, vtotal As Double
Dim vnro, vsql As String

'vvsaldo = Format(CalSaldoPersona(Me.txtCliente(0).Text, CP.TablaCtaCte), "$###,###,##0.00")
 
vnro = Str(getNroRecibo)
 
Call ActualizarRecibo(vid)

Unload Mantenimiento
Load Mantenimiento
 
 With drRecibo
        
  
        .Sections(2).Controls("etipo").Caption = Trim("Recibo")
        .Sections("TituloEmpresa").Controls("etiqueta1").Caption = "Recibo de cobro de Eventual"
        
        
        .Sections(2).Controls("etiqueta9").Caption = Str(Me.vfecha)
        '.Sections(2).Controls("lbllugar").Caption = vDatosEmpresa.Localidad & ", "
        .Sections(2).Controls("lblfecha").Caption = Str(Date)
        
        vsql = "select * from eventuales where idtrabeventuales =" + Str(vid)
        .Sections(2).Controls("lblCliente").Caption = "Por cuenta de: " + Trim(traerDatos2(vsql, "nombre", pathDBMySQL))
        
        Debug.Print .Sections(2).Controls("lblCliente").Caption
        
        
        
        .Sections(5).Controls("lblconcepto").Caption = Trim(Me.vobservaciones.Text)
        
        
        vtotal = traerDatos2(vsql, "importe", pathDBMySQL)
        .Sections(5).Controls("lbltotal").Caption = Format(vtotal, "$###,###,##0.00")
        
        Debug.Print .Sections(5).Controls("lbltotal").Caption
        
        
        .Sections(5).Controls("eletras").Caption = EnLetras(Str(vtotal))
        .Sections(5).Controls("esaldoTitulo").Caption = ""
        .Sections(5).Controls("esaldo").Caption = ""
        
        .Sections(2).Controls("enrorecibo").Caption = vnro
        
    End With



If Err Then
    'MsgBox "Error al intentar hacer el recibo" + Str$(Err)
    Exit Sub
End If

End Sub


Private Sub ActualizarRecibo(vid)
On Error Resume Next
Dim vsql, vlinea, vline As String
Dim i As Integer
Dim vtotal As Double

Dim rs As New ADODB.Recordset

vsql = "delete from recibo_temp"
Call EjecutarScript(vsql, pathDBMySQL)

vtotal = 0
i = vid


vsql = "select *  from eventuales where idtrabeventuales = " + Str(vid)
    
    With rs
        Call .Open(vsql, ConnDDBB, adOpenStatic, adLockPessimistic)
    End With


vline = "Período : desde " + Str(rs.Fields("fdesde")) + " hasta:" + Str(rs.Fields("fhasta"))
vsql = "insert into recibo_temp (descripcion,monto) values ('" + vline + "'," + Str(0) + ") "

Call EjecutarScript(vsql, pathDBMySQL)
'---
vline = "Horas por días :" + Format(rs.Fields("horas"), "###0.00") + " - Valor/hs: " + Format(rs.Fields("valorhora"), "###0.00")
vsql = "insert into recibo_temp (descripcion,monto) values ('" + vline + "'," + Str(rs.Fields("horas") * rs.Fields("valorhora")) + ") "

Call EjecutarScript(vsql, pathDBMySQL)
'---
vline = "Horas extras Total :" + Format(rs.Fields("horasextras"), "###0.00") + " - Valor/hs: " + Format(rs.Fields("valorhoraextra"), "###0.00")
vsql = "insert into recibo_temp (descripcion,monto) values ('" + vline + "'," + Str(rs.Fields("horasextras") * rs.Fields("valorhoraextra")) + ") "

Call EjecutarScript(vsql, pathDBMySQL)


'------
'vline = " >>>>>>>>>> Remuneración Total : "
'vsql = "insert into recibo_temp (descripcion,monto) values ('" + vline + "'," + Str(rs.Fields("importe")) + ") "
'
'Call EjecutarScript(vsql, pathDBMySQL)
'---------------------------


If Err Then
    Exit Sub
    'MsgBox Err.Description
    'Exit Sub
End If

End Sub

Private Sub vtotalSemanal_Change()
On Error Resume Next
Dim vt, vh, vhi As Double
Dim vdias As Integer
vh = Val(Me.vHoras)
vhi = 0


vdias = Me.vfhasta - Me.vfdesde
vt = Val(vtotalSemanal)

vhi = vt / vdias / vh

Me.vHorasImporte.Text = vhi


If Err Then Exit Sub
End Sub
