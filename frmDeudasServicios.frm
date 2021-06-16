VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.ShortcutBar.v13.0.0.Demo.ocx"
Begin VB.Form frmDeudasServicios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estado de  Ctas de Servcios "
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   14925
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.GroupBox GroupBox9 
      Height          =   690
      Left            =   45
      TabIndex        =   90
      Top             =   7605
      Width           =   11220
      _Version        =   851968
      _ExtentX        =   19791
      _ExtentY        =   1217
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton PusMarcarTodo 
         Height          =   240
         Left            =   90
         TabIndex        =   104
         Top             =   135
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2566
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Marcar Todo"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.TextBox vid 
         Height          =   330
         Left            =   9000
         TabIndex        =   91
         Top             =   180
         Width           =   1035
      End
      Begin MSComCtl2.DTPicker vperiodo 
         Height          =   300
         Left            =   2970
         TabIndex        =   92
         Top             =   225
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "MM/yyyy"
         Format          =   76218371
         CurrentDate     =   42479
      End
      Begin XtremeSuiteControls.PushButton PusBorrar 
         Height          =   330
         Left            =   10080
         TabIndex        =   93
         Top             =   195
         Width           =   1005
         _Version        =   851968
         _ExtentX        =   1773
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Borrar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit vimporte 
         Height          =   300
         Left            =   6300
         TabIndex        =   94
         Top             =   210
         Width           =   1485
         _Version        =   851968
         _ExtentX        =   2619
         _ExtentY        =   529
         _StockProps     =   77
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.PushButton PusAgregar 
         Height          =   330
         Left            =   7875
         TabIndex        =   95
         Top             =   195
         Width           =   1005
         _Version        =   851968
         _ExtentX        =   1773
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Agregar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PusDesmarcarTodo 
         Height          =   240
         Left            =   90
         TabIndex        =   105
         Top             =   405
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2566
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "DesmarcarTodo"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblPeríodo 
         Height          =   315
         Left            =   1980
         TabIndex        =   97
         Top             =   225
         Width           =   795
         _Version        =   851968
         _ExtentX        =   1402
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Período :"
      End
      Begin XtremeSuiteControls.Label lblImporte 
         Height          =   315
         Left            =   4545
         TabIndex        =   96
         Top             =   210
         Width           =   1725
         _Version        =   851968
         _ExtentX        =   3043
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Importe Deuda Original :"
      End
   End
   Begin XtremeSuiteControls.PushButton PusPagarEn 
      Height          =   1140
      Left            =   11385
      TabIndex        =   89
      Top             =   6345
      Width           =   3480
      _Version        =   851968
      _ExtentX        =   6138
      _ExtentY        =   2011
      _StockProps     =   79
      Caption         =   "Pagar en Caja"
      Appearance      =   6
   End
   Begin VB.CheckBox chkActivarInserción 
      Caption         =   "Activar inserción"
      Height          =   240
      Left            =   11430
      TabIndex        =   88
      Top             =   7920
      Width           =   3435
   End
   Begin XtremeSuiteControls.GroupBox framecomen 
      Height          =   1500
      Left            =   11385
      TabIndex        =   82
      Top             =   4725
      Width           =   3480
      _Version        =   851968
      _ExtentX        =   6138
      _ExtentY        =   2646
      _StockProps     =   79
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.FlatEdit vmes 
         Height          =   285
         Left            =   1845
         TabIndex        =   83
         Top             =   270
         Width           =   1410
         _Version        =   851968
         _ExtentX        =   2487
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton PushButton14 
         Height          =   330
         Left            =   945
         TabIndex        =   84
         Top             =   1080
         Width           =   2355
         _Version        =   851968
         _ExtentX        =   4154
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Agregar comentarios a recibos"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit vano 
         Height          =   285
         Left            =   1845
         TabIndex        =   86
         Top             =   630
         Width           =   1410
         _Version        =   851968
         _ExtentX        =   2487
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin VB.Label Label10 
         Caption         =   "Año: "
         Height          =   240
         Left            =   1305
         TabIndex        =   87
         Top             =   675
         Width           =   420
      End
      Begin VB.Label Label9 
         Caption         =   "Mes:"
         Height          =   240
         Left            =   1260
         TabIndex        =   85
         Top             =   360
         Width           =   465
      End
   End
   Begin RichTextLib.RichTextBox vlibredeuda 
      Height          =   750
      Left            =   0
      TabIndex        =   25
      Top             =   3915
      Width           =   14910
      _ExtentX        =   26300
      _ExtentY        =   1323
      _Version        =   393217
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      TextRTF         =   $"frmDeudasServicios.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla2 
      Height          =   2820
      Left            =   0
      TabIndex        =   24
      Top             =   4680
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   4974
      _Version        =   393216
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin XtremeSuiteControls.TabControl tab2 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14925
      _Version        =   851968
      _ExtentX        =   26326
      _ExtentY        =   6165
      _StockProps     =   68
      ItemCount       =   5
      Item(0).Caption =   "Libre Deuda"
      Item(0).ControlCount=   18
      Item(0).Control(0)=   "vcomentario"
      Item(0).Control(1)=   "vcontribuyente"
      Item(0).Control(2)=   "GroupBox1"
      Item(0).Control(3)=   "PusGenerarListado"
      Item(0).Control(4)=   "barra"
      Item(0).Control(5)=   "PusPersonas"
      Item(0).Control(6)=   "GroupBox2"
      Item(0).Control(7)=   "GroupBox3"
      Item(0).Control(8)=   "lblAgregarComentario"
      Item(0).Control(9)=   "PusCargarInformación"
      Item(0).Control(10)=   "GroupBox4"
      Item(0).Control(11)=   "Label1"
      Item(0).Control(12)=   "Label8"
      Item(0).Control(13)=   "vadesde"
      Item(0).Control(14)=   "vahasta"
      Item(0).Control(15)=   "GroDefinirPeríos"
      Item(0).Control(16)=   "GroOrdenadoPor"
      Item(0).Control(17)=   "GroupBox10"
      Item(1).Caption =   "Datos Generales"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "grilla"
      Item(1).Control(1)=   "PushButton1"
      Item(1).Control(2)=   "PusGenerar"
      Item(2).Caption =   "Indicadores Generales"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "indicadores"
      Item(3).Caption =   "Proyecciones"
      Item(3).ControlCount=   0
      Item(4).Caption =   "Pago-Link"
      Item(4).ControlCount=   3
      Item(4).Control(0)=   "GroGenerarArchivos"
      Item(4).Control(1)=   "log2"
      Item(4).Control(2)=   "TabControl2"
      Begin XtremeSuiteControls.GroupBox GroupBox10 
         Height          =   555
         Left            =   135
         TabIndex        =   98
         Top             =   2880
         Width           =   9330
         _Version        =   851968
         _ExtentX        =   16457
         _ExtentY        =   979
         _StockProps     =   79
         BackColor       =   4210752
         UseVisualStyle  =   -1  'True
         BorderStyle     =   2
         Begin XtremeSuiteControls.PushButton PusLimpiar 
            Height          =   285
            Left            =   8325
            TabIndex        =   99
            Top             =   180
            Width           =   885
            _Version        =   851968
            _ExtentX        =   1561
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Limpiar"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PusImprimir 
            Height          =   270
            Left            =   90
            TabIndex        =   100
            Top             =   150
            Width           =   1155
            _Version        =   851968
            _ExtentX        =   2037
            _ExtentY        =   476
            _StockProps     =   79
            Caption         =   "Imprimir"
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PusExportar 
            Height          =   270
            Left            =   1350
            TabIndex        =   101
            Top             =   165
            Width           =   1290
            _Version        =   851968
            _ExtentX        =   2275
            _ExtentY        =   476
            _StockProps     =   79
            Caption         =   "Exportar"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblDeudaAcumulada 
            Height          =   150
            Left            =   2880
            TabIndex        =   103
            Top             =   225
            Width           =   1725
            _Version        =   851968
            _ExtentX        =   3043
            _ExtentY        =   265
            _StockProps     =   79
            Caption         =   "Deuda Acumulada :"
            ForeColor       =   -2147483634
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
         End
         Begin XtremeShortcutBar.ShortcutCaption vsaldo 
            Height          =   345
            Left            =   4770
            TabIndex        =   102
            Top             =   90
            Width           =   3375
            _Version        =   851968
            _ExtentX        =   5953
            _ExtentY        =   609
            _StockProps     =   14
            Caption         =   "0.00"
            ForeColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            GradientColorLight=   12632256
            GradientColorDark=   16777215
            ForeColor       =   16777215
         End
      End
      Begin XtremeSuiteControls.GroupBox GroOrdenadoPor 
         Height          =   420
         Left            =   4770
         TabIndex        =   78
         Top             =   1755
         Width           =   4605
         _Version        =   851968
         _ExtentX        =   8123
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Ordenado por:"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton rdnombre 
            Height          =   150
            Left            =   1305
            TabIndex        =   79
            Top             =   180
            Width           =   960
            _Version        =   851968
            _ExtentX        =   1693
            _ExtentY        =   265
            _StockProps     =   79
            Caption         =   "Nombre"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rdcodigopersona 
            Height          =   195
            Left            =   2565
            TabIndex        =   80
            Top             =   180
            Width           =   870
            _Version        =   851968
            _ExtentX        =   1535
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Código"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rdImporte 
            Height          =   195
            Left            =   3645
            TabIndex        =   81
            Top             =   180
            Width           =   870
            _Version        =   851968
            _ExtentX        =   1535
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Importe"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GroDefinirPeríos 
         Height          =   555
         Left            =   1320
         TabIndex        =   72
         Top             =   1110
         Width           =   5655
         _Version        =   851968
         _ExtentX        =   9975
         _ExtentY        =   979
         _StockProps     =   79
         Caption         =   "Definir Períos para el calculo:"
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
         Begin MSComCtl2.DTPicker vfdesde 
            Height          =   285
            Left            =   690
            TabIndex        =   73
            Top             =   210
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            _Version        =   393216
            CalendarBackColor=   65535
            Format          =   76218369
            CurrentDate     =   42650
         End
         Begin MSComCtl2.DTPicker vfhasta 
            Height          =   285
            Left            =   2880
            TabIndex        =   74
            Top             =   210
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            _Version        =   393216
            CalendarBackColor=   65535
            Format          =   76218369
            CurrentDate     =   42650
         End
         Begin XtremeSuiteControls.PushButton PusFijar 
            Height          =   285
            Left            =   4560
            TabIndex        =   77
            Top             =   210
            Width           =   1035
            _Version        =   851968
            _ExtentX        =   1826
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Fijar 5 años"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label lblHasta 
            Caption         =   "Hasta:"
            Height          =   195
            Left            =   2340
            TabIndex        =   76
            Top             =   240
            Width           =   585
         End
         Begin VB.Label lblDesde 
            Caption         =   "Desde:"
            Height          =   195
            Left            =   150
            TabIndex        =   75
            Top             =   240
            Width           =   585
         End
      End
      Begin VB.TextBox vahasta 
         BackColor       =   &H0080FFFF&
         Height          =   315
         Left            =   5250
         TabIndex        =   69
         Top             =   1230
         Width           =   945
      End
      Begin VB.TextBox vadesde 
         BackColor       =   &H0080FFFF&
         Height          =   315
         Left            =   3450
         TabIndex        =   68
         Top             =   1230
         Width           =   885
      End
      Begin XtremeSuiteControls.TabControl TabControl2 
         Height          =   2805
         Left            =   -69820
         TabIndex        =   54
         Top             =   720
         Visible         =   0   'False
         Width           =   9075
         _Version        =   851968
         _ExtentX        =   16007
         _ExtentY        =   4948
         _StockProps     =   68
         Appearance      =   8
         Color           =   32
         PaintManager.BoldSelected=   -1  'True
         ItemCount       =   3
         SelectedItem    =   1
         Item(0).Caption =   "Paso 1: seleccionar lugar del archivo"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "GroupBox5"
         Item(1).Caption =   "Paso 2: Generar Archivos"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "GroupBox6"
         Item(2).Caption =   "Paso3: Mandar archivo a Pago Link"
         Item(2).ControlCount=   1
         Item(2).Control(0)=   "GroupBox7"
         Begin MSComDlg.CommonDialog dialogo 
            Left            =   7230
            Top             =   480
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin XtremeSuiteControls.GroupBox GroupBox5 
            Height          =   765
            Left            =   -69490
            TabIndex        =   55
            Top             =   900
            Visible         =   0   'False
            Width           =   8265
            _Version        =   851968
            _ExtentX        =   14579
            _ExtentY        =   1349
            _StockProps     =   79
            Caption         =   "Lugar donde guardar el archivo: "
            UseVisualStyle  =   -1  'True
            Begin XtremeSuiteControls.PushButton PushButton11 
               Height          =   315
               Left            =   150
               TabIndex        =   56
               Top             =   300
               Width           =   1965
               _Version        =   851968
               _ExtentX        =   3466
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "Seleccionar carpetas:"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit vpath2 
               Height          =   345
               Left            =   2430
               TabIndex        =   57
               Top             =   270
               Width           =   5655
               _Version        =   851968
               _ExtentX        =   9975
               _ExtentY        =   609
               _StockProps     =   77
               BackColor       =   16777215
               BackColor       =   16777215
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox6 
            Height          =   765
            Left            =   3720
            TabIndex        =   58
            Top             =   930
            Width           =   2685
            _Version        =   851968
            _ExtentX        =   4736
            _ExtentY        =   1349
            _StockProps     =   79
            Caption         =   "Lugar donde guardar el archivo: "
            UseVisualStyle  =   -1  'True
            Begin XtremeSuiteControls.PushButton PushButton12 
               Height          =   315
               Left            =   360
               TabIndex        =   59
               Top             =   300
               Width           =   1965
               _Version        =   851968
               _ExtentX        =   3466
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "Generar Archivo"
               UseVisualStyle  =   -1  'True
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox7 
            Height          =   765
            Left            =   -66310
            TabIndex        =   60
            Top             =   1350
            Visible         =   0   'False
            Width           =   2685
            _Version        =   851968
            _ExtentX        =   4736
            _ExtentY        =   1349
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Begin XtremeSuiteControls.PushButton PushButton13 
               Height          =   315
               Left            =   300
               TabIndex        =   61
               Top             =   300
               Width           =   1965
               _Version        =   851968
               _ExtentX        =   3466
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "Mandar archivo"
               UseVisualStyle  =   -1  'True
            End
         End
      End
      Begin XtremeSuiteControls.ListBox log2 
         Height          =   2775
         Left            =   -60610
         TabIndex        =   53
         Top             =   720
         Visible         =   0   'False
         Width           =   5205
         _Version        =   851968
         _ExtentX        =   9181
         _ExtentY        =   4895
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.GroupBox GroGenerarArchivos 
         Height          =   765
         Left            =   -64330
         TabIndex        =   49
         Top             =   1530
         Visible         =   0   'False
         Width           =   2985
         _Version        =   851968
         _ExtentX        =   5265
         _ExtentY        =   1349
         _StockProps     =   79
         Caption         =   "Generar archivos:"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.PushButton Refresh 
            Height          =   315
            Left            =   180
            TabIndex        =   50
            Top             =   330
            Width           =   825
            _Version        =   851968
            _ExtentX        =   1455
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Refresh"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PushButton9 
            Height          =   315
            Left            =   1050
            TabIndex        =   51
            Top             =   330
            Width           =   885
            _Version        =   851968
            _ExtentX        =   1561
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Extract"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PushButton10 
            Height          =   315
            Left            =   1980
            TabIndex        =   52
            Top             =   330
            Width           =   885
            _Version        =   851968
            _ExtentX        =   1561
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Extract"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.PushButton PusCargarInformación 
         Height          =   225
         Left            =   6360
         TabIndex        =   23
         Top             =   780
         Width           =   3075
         _Version        =   851968
         _ExtentX        =   5424
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Cargar información del Contribuyente"
         UseVisualStyle  =   -1  'True
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid indicadores 
         Height          =   3555
         Left            =   -69880
         TabIndex        =   22
         Top             =   420
         Visible         =   0   'False
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   6271
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin XtremeSuiteControls.PushButton PusGenerar 
         Height          =   405
         Left            =   -69880
         TabIndex        =   21
         Top             =   3510
         Visible         =   0   'False
         Width           =   825
         _Version        =   851968
         _ExtentX        =   1455
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Generar"
         UseVisualStyle  =   -1  'True
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla 
         Height          =   2985
         Left            =   -69880
         TabIndex        =   19
         Top             =   450
         Visible         =   0   'False
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   5265
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.TextBox vcontribuyente 
         Height          =   345
         Left            =   1680
         TabIndex        =   2
         Top             =   390
         Width           =   7755
      End
      Begin VB.TextBox vcomentario 
         BackColor       =   &H0080FFFF&
         Height          =   285
         Left            =   1710
         TabIndex        =   1
         Top             =   750
         Width           =   4335
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1065
         Left            =   120
         TabIndex        =   3
         Top             =   1035
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   1879
         _StockProps     =   79
         Caption         =   "Tipo:"
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
         Begin XtremeSuiteControls.RadioButton rdResumen 
            Height          =   315
            Left            =   90
            TabIndex        =   4
            Top             =   180
            Width           =   915
            _Version        =   851968
            _ExtentX        =   1614
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Resumen"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rdDetalle 
            Height          =   270
            Left            =   90
            TabIndex        =   5
            Top             =   450
            Width           =   795
            _Version        =   851968
            _ExtentX        =   1402
            _ExtentY        =   476
            _StockProps     =   79
            Caption         =   "Detalle"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
      End
      Begin XtremeSuiteControls.PushButton PusGenerarListado 
         Height          =   345
         Left            =   120
         TabIndex        =   6
         Top             =   2205
         Width           =   9315
         _Version        =   851968
         _ExtentX        =   16431
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Generar Listado"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ProgressBar barra 
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   2595
         Width           =   9330
         _Version        =   851968
         _ExtentX        =   16457
         _ExtentY        =   397
         _StockProps     =   93
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton PusPersonas 
         Height          =   345
         Left            =   135
         TabIndex        =   8
         Top             =   390
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2566
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Contribuyente"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmDeudasServicios.frx":0259
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   705
         Left            =   7020
         TabIndex        =   9
         Top             =   1035
         Width           =   2325
         _Version        =   851968
         _ExtentX        =   4101
         _ExtentY        =   1244
         _StockProps     =   79
         Caption         =   "Agrupado por:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton rdContribuyente 
            Height          =   180
            Left            =   45
            TabIndex        =   10
            Top             =   225
            Width           =   1305
            _Version        =   851968
            _ExtentX        =   2302
            _ExtentY        =   317
            _StockProps     =   79
            Caption         =   "Contribuyentes"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rdPeriodo 
            Height          =   180
            Left            =   30
            TabIndex        =   11
            Top             =   450
            Width           =   885
            _Version        =   851968
            _ExtentX        =   1561
            _ExtentY        =   317
            _StockProps     =   79
            Caption         =   "Períodos"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RadioButton5 
            Height          =   240
            Left            =   1410
            TabIndex        =   12
            Top             =   195
            Width           =   855
            _Version        =   851968
            _ExtentX        =   1508
            _ExtentY        =   423
            _StockProps     =   79
            Caption         =   "Importes"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RadZona 
            Height          =   180
            Left            =   1410
            TabIndex        =   13
            Top             =   450
            Width           =   795
            _Version        =   851968
            _ExtentX        =   1402
            _ExtentY        =   317
            _StockProps     =   79
            Caption         =   "Zona"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   420
         Left            =   1320
         TabIndex        =   14
         Top             =   1770
         Width           =   3330
         _Version        =   851968
         _ExtentX        =   5874
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Zona: "
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton rdUrbano 
            Height          =   180
            Left            =   630
            TabIndex        =   15
            Top             =   135
            Width           =   765
            _Version        =   851968
            _ExtentX        =   1349
            _ExtentY        =   317
            _StockProps     =   79
            Caption         =   "Urbano"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rdRural 
            Height          =   135
            Left            =   2475
            TabIndex        =   16
            Top             =   180
            Width           =   825
            _Version        =   851968
            _ExtentX        =   1455
            _ExtentY        =   238
            _StockProps     =   79
            Caption         =   "Rural"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RdComercio 
            Height          =   225
            Left            =   1485
            TabIndex        =   17
            Top             =   135
            Width           =   1035
            _Version        =   851968
            _ExtentX        =   1826
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "Comercio"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   405
         Left            =   -69010
         TabIndex        =   20
         Top             =   3510
         Visible         =   0   'False
         Width           =   1515
         _Version        =   851968
         _ExtentX        =   2672
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Exportar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox4 
         Height          =   3075
         Left            =   9510
         TabIndex        =   26
         Top             =   360
         Width           =   5295
         _Version        =   851968
         _ExtentX        =   9340
         _ExtentY        =   5424
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin VB.CheckBox chkMostrarTexto 
            Caption         =   "Mostrar texto"
            Height          =   195
            Left            =   3960
            TabIndex        =   67
            Top             =   2790
            Value           =   1  'Checked
            Width           =   1245
         End
         Begin XtremeSuiteControls.PushButton PusReemplazarCampos 
            Height          =   285
            Left            =   120
            TabIndex        =   41
            Top             =   2730
            Width           =   1905
            _Version        =   851968
            _ExtentX        =   3360
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Reemplazar [Campos]"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit vc 
            Height          =   345
            Index           =   0
            Left            =   1800
            TabIndex        =   27
            Top             =   180
            Width           =   2925
            _Version        =   851968
            _ExtentX        =   5159
            _ExtentY        =   609
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vc 
            Height          =   345
            Index           =   1
            Left            =   1800
            TabIndex        =   28
            Top             =   540
            Width           =   2955
            _Version        =   851968
            _ExtentX        =   5212
            _ExtentY        =   609
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vc 
            Height          =   345
            Index           =   2
            Left            =   1800
            TabIndex        =   29
            Top             =   900
            Width           =   2955
            _Version        =   851968
            _ExtentX        =   5212
            _ExtentY        =   609
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vc 
            Height          =   345
            Index           =   3
            Left            =   1800
            TabIndex        =   30
            Top             =   1260
            Width           =   2985
            _Version        =   851968
            _ExtentX        =   5265
            _ExtentY        =   609
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vc 
            Height          =   345
            Index           =   4
            Left            =   1800
            TabIndex        =   31
            Top             =   1650
            Width           =   2985
            _Version        =   851968
            _ExtentX        =   5265
            _ExtentY        =   609
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vc 
            Height          =   345
            Index           =   5
            Left            =   1800
            TabIndex        =   32
            Top             =   2010
            Width           =   2985
            _Version        =   851968
            _ExtentX        =   5265
            _ExtentY        =   609
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vc 
            Height          =   345
            Index           =   6
            Left            =   1800
            TabIndex        =   33
            Top             =   2370
            Width           =   2985
            _Version        =   851968
            _ExtentX        =   5265
            _ExtentY        =   609
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "Rural"
         End
         Begin XtremeSuiteControls.PushButton PushButton2 
            Height          =   255
            Left            =   4830
            TabIndex        =   42
            Top             =   180
            Width           =   405
            _Version        =   851968
            _ExtentX        =   714
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Act."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PushButton3 
            Height          =   255
            Left            =   4830
            TabIndex        =   43
            Top             =   510
            Width           =   405
            _Version        =   851968
            _ExtentX        =   714
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Act."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PushButton4 
            Height          =   255
            Left            =   4830
            TabIndex        =   44
            Top             =   870
            Width           =   405
            _Version        =   851968
            _ExtentX        =   714
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Act."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PushButton5 
            Height          =   255
            Left            =   4830
            TabIndex        =   45
            Top             =   1260
            Width           =   405
            _Version        =   851968
            _ExtentX        =   714
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Act."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PushButton6 
            Height          =   255
            Left            =   4860
            TabIndex        =   46
            Top             =   1650
            Width           =   405
            _Version        =   851968
            _ExtentX        =   714
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Act."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PushButton7 
            Height          =   255
            Left            =   4860
            TabIndex        =   47
            Top             =   2010
            Width           =   405
            _Version        =   851968
            _ExtentX        =   714
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Act."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PushButton8 
            Height          =   255
            Left            =   4860
            TabIndex        =   48
            Top             =   2400
            Width           =   405
            _Version        =   851968
            _ExtentX        =   714
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Act."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblCampo1 
            Height          =   345
            Left            =   180
            TabIndex        =   40
            Top             =   180
            Width           =   1605
            _Version        =   851968
            _ExtentX        =   2831
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "[Campo1] Apellido:"
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   345
            Left            =   180
            TabIndex        =   39
            Top             =   510
            Width           =   1425
            _Version        =   851968
            _ExtentX        =   2514
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "[Campo2] Calle:"
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   345
            Left            =   180
            TabIndex        =   38
            Top             =   870
            Width           =   1395
            _Version        =   851968
            _ExtentX        =   2461
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "[Campo3] Partida"
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   345
            Left            =   180
            TabIndex        =   37
            Top             =   1230
            Width           =   1365
            _Version        =   851968
            _ExtentX        =   2408
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "[Campo4]: HAS"
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   345
            Left            =   180
            TabIndex        =   36
            Top             =   1590
            Width           =   1065
            _Version        =   851968
            _ExtentX        =   1879
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "[Campo5] Dia:"
         End
         Begin XtremeSuiteControls.Label Label6 
            Height          =   345
            Left            =   180
            TabIndex        =   35
            Top             =   1950
            Width           =   1335
            _Version        =   851968
            _ExtentX        =   2355
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "[Campo6] Suma:"
         End
         Begin XtremeSuiteControls.Label Label7 
            Height          =   345
            Left            =   180
            TabIndex        =   34
            Top             =   2340
            Width           =   1365
            _Version        =   851968
            _ExtentX        =   2408
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "[Campo7] Tasa:"
         End
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   255
         Left            =   4380
         TabIndex        =   71
         Top             =   1260
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Año hasta:"
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Left            =   2490
         TabIndex        =   70
         Top             =   1260
         Width           =   915
         _Version        =   851968
         _ExtentX        =   1614
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Año desde:"
      End
      Begin XtremeSuiteControls.Label lblAgregarComentario 
         Height          =   255
         Left            =   150
         TabIndex        =   18
         Top             =   750
         Width           =   1515
         _Version        =   851968
         _ExtentX        =   2672
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Agregar comentario: "
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox8 
      Height          =   480
      Left            =   0
      TabIndex        =   62
      Top             =   3420
      Width           =   14925
      _Version        =   851968
      _ExtentX        =   26326
      _ExtentY        =   847
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.Label veid4 
         Height          =   150
         Left            =   6840
         TabIndex        =   66
         Top             =   270
         Width           =   6825
         _Version        =   851968
         _ExtentX        =   12039
         _ExtentY        =   265
         _StockProps     =   79
         Caption         =   "4"
         ForeColor       =   32768
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.Label veid3 
         Height          =   135
         Left            =   6840
         TabIndex        =   65
         Top             =   135
         Width           =   6825
         _Version        =   851968
         _ExtentX        =   12039
         _ExtentY        =   238
         _StockProps     =   79
         Caption         =   "3"
         ForeColor       =   16448
         BackColor       =   14737632
      End
      Begin XtremeSuiteControls.Label veid2 
         Height          =   135
         Left            =   45
         TabIndex        =   64
         Top             =   285
         Width           =   6690
         _Version        =   851968
         _ExtentX        =   11800
         _ExtentY        =   238
         _StockProps     =   79
         Caption         =   "2"
         ForeColor       =   255
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.Label veid1 
         Height          =   150
         Left            =   30
         TabIndex        =   63
         Top             =   120
         Width           =   6690
         _Version        =   851968
         _ExtentX        =   11800
         _ExtentY        =   265
         _StockProps     =   79
         Caption         =   "1"
         ForeColor       =   16711680
         BackColor       =   14737632
      End
   End
End
Attribute VB_Name = "frmDeudasServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vcaducidad, vcp, vvtitulo As String
Dim vtotal9 As Double

Dim recibosApagar(50) As String


Private Sub chkActivarInserción_Click()

If Me.chkActivarInserción.Value Then
    Me.framecomen.Enabled = True
Else
    Me.framecomen.Enabled = False
End If

End Sub

Private Sub Form_Load()

Me.Width = 15015
Me.Height = 8730

vperiodo.Value = Date - Day(Date) + 1

tab2.SelectedItem = 0

init

End Sub


Private Sub init()
Dim v As String

On Error Resume Next


Me.vano = Year(Now())

Me.vlibredeuda.Text = "La COMUNA DE WHEELWRIGHT CERTIFICA QUE EL SR/A.: [campo1] , domiciliado en [campo2] , actual propietario del inmueble cuya partida inmobiliaria es la siguiente: [campo3], con una extensión de: [campo4] has, respectivamente, adeuda al dia: [campo5] , en concepto de capital e intereses conforme a la Ordenanza Tributaria en vigencia, la suma de  [campo6],  por la Tasa [campo7], que corresponde a los períodos que se detallan en el Anexo que junto al presente se acompaña."

set_indicadores

v = Me.vlibredeuda.Text

If UCase(LeerXml("Cliente")) = "COMUNAWW" Then
    Me.vlibredeuda.Text = Replace$(v, "JUNCAL", "WHEELWRIGHT")
Else
    Me.vlibredeuda.Text = Replace$(v, "WHEELWRIGHT", "JUNCAL")
End If

vfdesde.Value = Date
vfhasta.Value = Date


If Err Then Exit Sub
End Sub


Private Sub set_indicadores()
Dim vsql As String

If UCase$(LeerXml("Cliente")) = UCase$("ComunaWw") Then
        
        vsql = " select concat('Acumulado Urbano (Orig, Intereses, Total) : ', ORIGINAL, ' + ', Interes, ' = ', Deuda) As c From deudaurbanatotal "
        
        Me.veid1.Caption = traerDatos2(vsql, "c", ConnComunaDB.ConnectionString)
        
        vsql = " select concat('Acumulado RURAL (Orig, Intereses, Total) : ', ORIGINAL, ' + ', Interes, ' = ', Deuda) As c From deudaruraltotal "
        
        Me.veid2.Caption = traerDatos2(vsql, "c", ConnComunaDB.ConnectionString)
End If

End Sub

Function sqlTotal(vzona As String, pago As Boolean) As String
On Error Resume Next

sqlTotal = "select  Apellido, '' as Periodo, sum(ImporteTotal1) as ImporteTotal1, sum(Intereses) as Intereses,  sum(Deuda) as Deuda" + _
" from deudarural t Where t.zona = '" + vzona + "' Group By t.ID_Persona order by Apellido"

If Err Then Exit Function
End Function



Function sqlruralresumen(vzona As String) As String
On Error Resume Next


Dim vorderby As String


If Me.rdnombre.Value = True Then vorderby = "Nombre"
If Me.rdcodigopersona.Value = True Then vorderby = "ID_Persona"
If Me.rdImporte.Value = True Then vorderby = "Deuda desc"




If vzona = "Rural" Then

        If UCase$(LeerXml("Cliente")) = UCase$("ComunaWw") Then
                sqlruralresumen = "select * from deudaRuralResumida order by " + vorderby
        Else
        
                sqlruralresumen = "select concat(apellido, ' [ ',count(Periodo),' ]' ) as Apellido,  sum(ImporteTotal1) as ImporteTotal1, sum(Intereses) as Intereses,  sum(Deuda) as Deuda" + _
                " from deudarural t Where t.zona = '" + vzona + "'" + vcaducidad + " Group By t.ID_Persona order by Apellido"

                 sqlruralresumen = "select apellido,  Nombre, count(Periodo) as Cantidad, '' as Periodo, Hectareas, sum(ImporteTotal1) as ImporteTotal1, sum(Intereses) as Intereses,  sum(Deuda) as Deuda" + _
                " from deudarural t Where t.zona = '" + vzona + "'" + vcaducidad + " Group By t.ID_Persona order by Apellido"



        End If

End If


If vzona = "Urbano" Then


        If UCase$(LeerXml("Cliente")) = UCase$("ComunaWw") Then
                 
                 sqlruralresumen = "select * from deudaurbanaresumida order by " + vorderby
        
        Else
                
                sqlruralresumen = "select  Apellido,count(Periodo) as Cantidad, '' as Periodo, sum(ImporteTotal1) as ImporteTotal1, sum(Intereses) as Intereses,  sum(Deuda) as Deuda" + _
                " from deudaurbana t Where t.zona = '" + vzona + "'" + vcaducidad + "  Group By t.ID_Persona order by Apellido"

        End If

End If



If vzona = "Comercio" Then


        If UCase$(LeerXml("Cliente")) = UCase$("ComunaWw") Then
                 
                 sqlruralresumen = "select * from deudaurbanaresumida order by 1 "
        
        Else
                
                sqlruralresumen = "select  Apellido,count(Periodo) as Cantidad, '' as Periodo, sum(ImporteTotal1) as ImporteTotal1, sum(Intereses) as Intereses,  sum(Deuda) as Deuda" + _
                " from deudaurbana t Where t.zona = '" + vzona + "'" + vcaducidad + "  Group By t.ID_Persona order by Apellido"

        End If

End If


If Err Then Exit Function
End Function


Private Sub grilla2_Click()

Dim i As Long

i = grilla2.Row

Me.vid.Text = grilla2.TextMatrix(i, 1)

End Sub

Private Sub grilla2_DblClick()
Dim vr As Integer

vr = grilla2.Row

If grilla2.TextMatrix(vr, 0) = "X" Then
    grilla2.TextMatrix(vr, 0) = ""
Else
    grilla2.TextMatrix(vr, 0) = "X"
End If

End Sub

Private Sub PusAgregar_Click()
Dim vsql, vcampo, vValor, vzona, vvalores As String



If MsgBox("Confirma el ingreso ?", vbYesNo) = vbNo Then
    Exit Sub
End If

If rdRural.Value Then
    vzona = "Rural"
End If


If Me.rdUrbano Then
    vzona = "Urbano"
End If


' __ todo ________________________________
' poner la insersión en comunadb
'_________________________________________



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
    Exit Sub
End If


vsql = "delete from t_recibos where Nrorecibo = " + Me.vid.Text

Dim rec2 As New Recordset


Call rec2.Open(vsql, ConnComunaDB, adOpenDynamic, adLockPessimistic)
    
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

If UCase(LeerXml("Cliente")) = UCase("ComunaWw") Then
    cargar_datos_ComunaDB
    Exit Sub
End If



'----------------------------




'vsql = "select t.ID_Zona  as c from  t_contribuyente t where t.id_persona = " + Str(Me.vcontribuyente.Tag)
'vzona = Val(EsNulo(traerDatos2(vsql, "c", pathDBMySQLComuna)))

'vsql = "select t.MetrosFrente as c from  t_contribuyente t where t.id_persona = " + Str(Me.vcontribuyente.Tag)
'vmfrente = Val(EsNulo(traerDatos2(vsql, "c", pathDBMySQLComuna)))

'vsql = "select t.Terreno  as c from  t_contribuyente t where t.id_persona = " + Str(Me.vcontribuyente.Tag)
'vmterreno = Val(EsNulo(traerDatos2(vsql, "c", pathDBMySQLComuna)))





'--------------------



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
    vc(3).Text = vmterreno
     
   
    
End If

Me.vcomentario.Text = vdatos


Call llenar_campos


If Err Then Exit Sub
End Sub

Private Sub cargar_datos_ComunaDB()
On Error Resume Next
Dim vsql As String
Dim vdatos As String
Dim vzona As Integer

Dim vmfrente, vmterreno As Double



vsql = "select t.id_zonas  as c from  contribuyentes_rurales t where  t.id_contribuyentes = " + Str(Me.vcontribuyente.Tag)
vzona = Val(EsNulo(traerDatos2(vsql, "c", pathDBMySQLComuna)))

vsql = "select t.hectareas as c from  contribuyentes_rurales t where t.id_contribuyentes = " + Str(Me.vcontribuyente.Tag)
vmfrente = Val(EsNulo(traerDatos2(vsql, "c", pathDBMySQLComuna)))

'vsql = "select t.Terreno  as c from  t_contribuyente t where t.ID_contribuyente = " + Str(Me.vcontribuyente.Tag)
'vmterreno = Val(EsNulo(traerDatos2(vsql, "c", pathDBMySQLComuna)))



    vdatos = ""

    vdatos = vdatos + "Hectareas : [" + Str(vmfrente) + "]"
    
    If Val(vc(3).Text) = 0 Then vc(3).Text = vmfrente

    Me.vcomentario.Text = vdatos


If Err Then Exit Sub

End Sub


Private Sub PusDesmarcarTodo_Click()
On Error Resume Next
Dim i, j  As Integer

With grilla2
    
For i = 1 To .Rows - 1
     .TextMatrix(i, 0) = ""
Next

End With

If Err Then Exit Sub
End Sub

Private Sub PusExportar_Click()
Call grillaToExcel(Me.grilla2)
End Sub

Private Sub PusFijar_Click()
Dim v As String

' -----------
Me.vadesde = Year(Date) - 5
Me.vahasta = Year(Date)


v = Str(Day(Date)) + "/" + Str(Month(Date)) + "/" + Str(Year(Date) - 5)

vfdesde = CDate(v)


End Sub

Public Sub PusGenerarListado_Click()
On Error Resume Next

Dim vsql, vvalue, vcampos  As String
Dim rec As New Recordset
Dim vdatosgrales As String

init

vcampos = "c1,c2,c3,c4,c5"

vcaducidad = ""

If Not (Me.vadesde.Text = "") Then
    vcaducidad = vcaducidad + " and year(periodo) >= " + Str(vadesde.Text)
End If

If Not (Me.vadesde.Text = "") Then
    vcaducidad = vcaducidad + " and year(periodo) <= " + Str(vahasta.Text)
End If

vcaducidad = ""

If Not vfdesde.Value = vfhasta.Value And UCase$(LeerXml("Cliente")) = "COMUNAWW" Then
    vcaducidad = ""
    vcaducidad = vcaducidad + " and (fecha_vencimiento2) >= '" + strfechaMySQL(vfdesde) + "' "
    vcaducidad = vcaducidad + " and (fecha_vencimiento2) <= '" + strfechaMySQL(vfhasta) + "'"
End If



If Not vfdesde.Value = vfhasta.Value And Not UCase$(LeerXml("Cliente")) = "COMUNAWW" Then
    vcaducidad = ""
    vcaducidad = vcaducidad + " and (periodo) >= '" + strfechaMySQL(vfdesde) + "' "
    vcaducidad = vcaducidad + " and (periodo) <= '" + strfechaMySQL(vfhasta) + "'"
End If


Me.barra.Value = 0
vsql = ""

        vsql = "delete from Rural"
        Call EjecutarScript(vsql, PathDBListados)

vsql = ""

If Me.rdResumen Then

    If Me.rdUrbano Then
        vsql = sqlruralresumen("Urbano")
    End If


    If Me.rdRural Then
        vsql = sqlruralresumen("Rural")
    End If
    
    
    If Me.RdComercio Then
        vsql = "select * from deudacomerciototal order by 2"
    End If
  

End If


If Me.rdDetalle Then
    vsql = " select * from deudaRural where ID_Persona = " + Me.vcontribuyente.Tag + vcaducidad
    
     If Me.rdUrbano Then
        vsql = " select * from deudaUrbana where ID_Persona = " + Me.vcontribuyente.Tag + vcaducidad
    
    End If


    If Me.rdRural Then
        vsql = " select * from deudaRural where ID_Persona = " + Me.vcontribuyente.Tag + vcaducidad
    
    End If
    
    If Me.RdComercio Then
        vsql = " select * from deudacomercio where id_comercios = " + Me.vcontribuyente.Tag + " order by 1 desc"
        vvtitulo = "Nro. recibo"
    End If

    
End If



With rec


    If UCase(LeerXml("Cliente")) = UCase("ComunaWw") Then
            vsql = vsql + " "
    Else
    
        If Not Me.rdResumen Then vsql = vsql + " order by Periodo "
    
    
    End If
    
    Call .Open(vsql, ConnComunaDB, adOpenDynamic, adLockReadOnly)
    barra.Max = 0
    barra.Max = .RecordCount
    
    Dim vtotal As String
    
    vtotal = 0
    
    If Not barra.Max > 0 Then
        MsgBox "No hay datos para mostar"
        Exit Sub
    End If
    
    'Dim vdatosgrales As String
    
    If UCase(LeerXml("Cliente")) = UCase("ComunaWw") Then
   
                             
                             
                Dim vcant As String
                
                'Do Until .EOF Or (barra.Value + 1 > barra.Max)
                
                Dim vsql11, vpartidas  As String
                
                Do Until .EOF
                                
                        barra.Value = barra.Value + 1
                        
                         If Me.rdResumen.Value = True Then
                         
                            vcant = Str(getCantDeuda(.Fields("id_persona")))
                            vcant = " - [" + vcant + "]"
                            'vcant = ""
                         
                            vdatosgrales = fdatosgrales(.Fields("id_persona")) + _
                            " P: (" + Str(.Fields("periodos")) + ")"
                            
                            
                            vvalue = vvalue + "'" + _
                            Str(.Fields("ID_Persona")) + " - " + _
                            Left(.Fields("Nombre"), 100) + vcant + "',"
                            vvalue = vvalue + "'" + vdatosgrales + "',"
                            vvalue = vvalue + Str(.Fields("Original")) + ","
                            vvalue = vvalue + Str(.Fields("Intereses")) + ","
                            
                            vtotal9 = .Fields("Original") + .Fields("Intereses")
                            vvalue = vvalue + Str(vtotal9)
                            
                            ' todo 2019
                            ' vvalue = vvalue + Str(.Fields("Deuda"))
                            
                           
                            
                        Else
                            vvalue = vvalue + "'" + _
                            Str(.Fields("ID_Persona")) + " - " + _
                            Left(Replace(Me.vcontribuyente, "'", " "), 100) + " ',"
                            vvalue = vvalue + "'" + Format(.Fields("periodo_anomes"), "mm-yyyy") + " - Recibo: " + .Fields("nro_recibo") + _
                            " ',"
                            vvalue = vvalue + Str(.Fields("importe_total2")) + ","
                            vvalue = vvalue + Str(.Fields("Intereses")) + ","
                            
                            ' vvalue = vvalue + Str(.Fields("Total"))
                            
                            vtotal9 = .Fields("importe_total2") + .Fields("Intereses")
                            
                            vvalue = vvalue + Str(vtotal9)
                            
                        End If
                        
                            'vtotal9 = .Fields("importe_total2") + .Fields("Intereses")
                        
                            ' vtotal = vtotal + .Fields("Total")
                            
                            vtotal = vtotal + vtotal9
                        
                            barra.Value = barra.Value + 1
                    
                              
                            vsql = "insert into Rural (" + vcampos + ") values (" + vvalue + ")"
                            Call EjecutarScript(vsql, PathDBListados)
                        
                            Debug.Print "  >>>>>>>>> " + vsql
                        
                            vvalue = ""
                            
                            .MoveNext
                            
                Loop
   
    Else
    
                Dim vnombre As String
    
                Do Until .EOF
                
               
                    If Not Me.rdDetalle.Value And Not UCase(LeerXml("Cliente")) = UCase("ComunaWw") Then
                
                        vnombre = Left(.Fields("Apellido"), 50) + " " + Left(.Fields(1), 10) + "  [" + Str(.Fields("Cantidad")) + "]"
                
                    Else
                
                         vnombre = Left(.Fields("Apellido"), 50) + " " + Left(.Fields(1), 10)
                    
                    End If
                
                    vvalue = vvalue + "'" + vnombre + "',"
                    vvalue = vvalue + "'" + Format(.Fields("Periodo"), "mmm-yyyy") + " - Hras: " _
                    + Format(.Fields("Hectareas")) + "',"
                    vvalue = vvalue + Str(.Fields("ImporteTotal1")) + ","
                    vvalue = vvalue + Str(.Fields("Intereses")) + ","
                    vvalue = vvalue + Str(.Fields("Deuda"))
                    
                    vtotal = vtotal + .Fields("Deuda")
                
                    barra.Value = barra.Value + 1
                      
                      vsql = "insert into Rural (" + vcampos + ") values (" + vvalue + ")"
                    Call EjecutarScript(vsql, PathDBListados)
                
                    vvalue = ""
                    
                    .MoveNext
                    
                Loop
                
    End If
                
    'vsql = "insert into Rural (" + vcampos + ") values (" + vvalores + ")"
    'Call EjecutarScript(vsql, PathDBListados)
     
   Me.vsaldo.Caption = Format(vtotal, "###,###,##0.00")
   
   Me.vc(5).Text = Format(vtotal, "###,###,##0.00")
    
   'Set Me.g2e.Recordset = rec
   
   Set grilla2.DataSource = rec.DataSource
   Call recalculartotalesenGrilla
    
End With


barra.Value = 0
Me.PusImprimir.Enabled = True


Call llenar_campos

Call PusReemplazarCampos_Click

If Err Then Exit Sub
End Sub

Private Sub recalculartotalesenGrilla()

Dim j, i As Integer


For i = 0 To grilla2.Rows - 1
    Debug.Print grilla2.TextMatrix(i, 0)
    Debug.Print grilla2.TextMatrix(i, 1)
    Debug.Print grilla2.TextMatrix(i, 2)
    
    Debug.Print grilla2.TextMatrix(i, 3)
    Debug.Print grilla2.TextMatrix(i, 4)
    Debug.Print grilla2.TextMatrix(i, 5)
    
    Debug.Print grilla2.TextMatrix(i, 6)
    Debug.Print grilla2.TextMatrix(i, 7)
    Debug.Print "1) " + grilla2.TextMatrix(i, 8)
    
    grilla2.TextMatrix(i, 8) = Val(grilla2.TextMatrix(i, 6)) + Val(grilla2.TextMatrix(i, 7))
    
    Debug.Print "2) " + grilla2.TextMatrix(i, 8)
    
Next

End Sub

Function fdatosgrales(vid As Long) As String
On Error Resume Next

Dim vv, vsql  As String

If UCase(LeerXml("Cliente")) = "COMUNAWW" Then

    vsql = "select *  from personas where id_personas = " + Str(vid)
    


    vv = "D:" + traerDatos2(vsql, "direccion", pathDBMySQLComuna)
    
    
    vsql = "select *  from contribuyentes_rurales where id_contribuyentes = " + Str(vid)
    
    vv = vv + " -  H:[" + traerDatos2(vsql, "hectareas", pathDBMySQLComuna) + "]"
    

End If

fdatosgrales = vv

If Err Then
    fdatosgrales = ""
End If

End Function


Function fsql2() As String
Dim v As String

v = ""

If rdUrbano.Value = True Then

    v = v + "select *   from  contribuyentes t "
    v = v + " inner join personas p on p.id_personas = t.id_contribuyentes  where p.ID_Personas = " + Me.vcontribuyente.Tag

Else
  
    v = v + "select *   from  contribuyentes_rurales t "
    v = v + " inner join personas p on p.id_personas = t.id_contribuyentes  where p.ID_Personas = " + Me.vcontribuyente.Tag

End If

fsql2 = v
End Function


Function fsql1() As String
Dim v As String

v = ""
v = v + "select Apellido,  Localidad, PadronTasaAnt, calle, numero   from  t_contribuyente t "
v = v + " inner join t_personas p on t.ID_persona = p.ID_persona  where p.ID_Persona = " + Me.vcontribuyente.Tag

fsql1 = v
End Function


Function get_Metros(vid As String) As String
Dim v As String
Dim m, m2 As String

    v = ""
    v = v + "select metros     from  contribuyentes_tasas t "
    v = v + " inner join personas p on p.ID_personas = t.id_contribuyentes  where id_ubicaciones =1 and   p.ID_Personas = " + Me.vcontribuyente.Tag

m = "M. Frente: " + traerDatos2(v, "metros", pathDBMySQLComuna)

  
    v = ""
    v = v + "select metros      from  contribuyentes_tasas t "
    v = v + " inner join personas p on p.ID_personas = t.id_contribuyentes  where id_ubicaciones =2 and   p.ID_Personas = " + Me.vcontribuyente.Tag

    m2 = traerDatos2(v, "metros", pathDBMySQLComuna)

    If Not m2 = "" Then
        m = m + " Lateral:  " + traerDatos2(v, "metros", pathDBMySQLComuna)
    End If

get_Metros = m

End Function

Private Sub llenar_campos()

Dim vsql As String

If UCase$(LeerXml("Cliente")) = UCase$("Comunaww") Then

            
            If rdUrbano.Value = True Then
                     
                    Me.vc(0).Text = traerDatos2(fsql2, "Apellido", pathDBMySQLComuna) + ", " + traerDatos2(fsql2, "Nombre", pathDBMySQLComuna)
                    Me.vc(1).Text = traerDatos2(fsql2, "direccion", pathDBMySQLComuna)
                    Me.vc(2).Text = traerDatos2(fsql2, "partida", pathDBMySQLComuna)
                    Me.vc(3).Text = get_Metros(Me.vcontribuyente.Tag)

            Else
                    Me.vc(0).Text = traerDatos2(fsql2, "Apellido", pathDBMySQLComuna) + ", " + traerDatos2(fsql2, "Nombre", pathDBMySQLComuna)
                    Me.vc(1).Text = traerDatos2(fsql2, "direccion", pathDBMySQLComuna)
                    Me.vc(2).Text = traerDatos2(fsql2, "impuesto_inmobiliario", pathDBMySQLComuna)
                    Me.vc(3).Text = traerDatos2(fsql2, "hectareas", pathDBMySQLComuna)
            End If


            'Me.vc(3).Text = traerDatos2(fsql1, "Apellido", pathDBMySQLComuna)
            'Me.vc(0).Text = traerDatos2(fsql1, "Apellido", pathDBMySQLComuna)
            'Me.vc(0).Text = traerDatos2(fsql1, "Apellido", pathDBMySQLComuna)
            
            Me.vc(4).Text = CDate(Now())
Else
            Me.vc(0).Text = traerDatos2(fsql1, "Apellido", pathDBMySQLComuna)
            Me.vc(1).Text = traerDatos2(fsql1, "calle", pathDBMySQLComuna)
            Me.vc(2).Text = traerDatos2(fsql1, "PadronTasaAnt", pathDBMySQLComuna)
            'Me.vc(3).Text = traerDatos2(fsql1, "Apellido", pathDBMySQLComuna)
            'Me.vc(0).Text = traerDatos2(fsql1, "Apellido", pathDBMySQLComuna)
            'Me.vc(0).Text = traerDatos2(fsql1, "Apellido", pathDBMySQLComuna)
            
            Me.vc(4).Text = CDate(Now())
End If

If UCase(LeerXml("Cliente")) = "COMUNAWW" Then
    
    vsql = "select hectareas from contribuyentes_rurales where activo = 'S' and  id_contribuyentes = " + Str(Me.vcontribuyente.Tag)
    
    vc(3).Text = getSqlTotal(vsql, pathDBMySQLComuna)
    
    vsql = "select impuesto_inmobiliario from contribuyentes_rurales where activo = 'S' and id_contribuyentes = " + Str(Me.vcontribuyente.Tag)
    
    vc(2).Text = getSqlTotal(vsql, pathDBMySQLComuna)

End If

End Sub


Private Sub PushButton11_Click()
On Error Resume Next
        
        Me.dialogo.ShowOpen
        
        vpath2.Text = Me.dialogo.FileName
        
    
If Err Then Exit Sub
End Sub

Private Sub PushButton12_Click()


'---------------------------------
Pago_Link.init ' inicializo las variables por defoult
'----------------------------------

Pago_Link.abrir_archivo (Refresh_nombre("EEE", 1))

'Call Pago_Link.Refresh_nombre("EEE", 1)
Call Pago_Link.Refresh_Inicial
Call Pago_Link.Refresh_Datos
Call Pago_Link.Refresh_Final

Pago_Link.cerrar_archivo

'------------------------------------

Pago_Link.abrir_archivo (Pago_Link.Control_nombre("EEE", 1))

'Call Pago_Link.Control_nombre("EEE", 1)
Call Pago_Link.Control_Inicial
Call Pago_Link.Control_Datos
Call Pago_Link.Control_Final

Pago_Link.cerrar_archivo
'------------------------------------

Pago_Link.abrir_archivo (Pago_Link.Extract_nombre("EEE"))

'Call Pago_Link.Extract_nombre("EEE")

Call Pago_Link.Extract_Inicial
Call Pago_Link.Extract_Datos
Call Pago_Link.Extract_Final

Pago_Link.cerrar_archivo

End Sub

Private Sub PushButton14_Click()

Dim rec As New Recordset

Dim vsql, vsql2, vmesano As String

vmesano = Trim$(Me.vmes.Text) + "/" + Trim$(Me.vano)


vsql2 = "select nro_recibo as c from recibo_resumen where periodo_anomes = '" + vmesano + "'"


Dim r As String
r = traerDatos2(vsql2, "c", pathDBMySQL)

If r = "" Then

    MsgBox "El perìodo no fue liquidado", vbInformation
    
    Exit Sub
End If

vsql = "select * from deudaruralresumida "

Call rec.Open(vsql, ConnComunaDB, adOpenDynamic, adLockPessimistic)

Dim vid As Long


Do While rec.EOF

    vid = rec.Fields("ID_Persona")
    

    rec.MoveNext
Loop

End Sub

Private Sub PushButton2_Click()
    Call fact(1)
End Sub


Private Sub fact(i As Integer)

If vc(i - 1).Text = "" Then Exit Sub

Dim v As String
v = Me.vlibredeuda.Text
Debug.Print v

Me.vlibredeuda.Text = Replace(v, "[campo" + Trim(Str(i)) + "]", vc(i - 1).Text)
    
Debug.Print v
Debug.Print Chr(13)
Debug.Print Me.vlibredeuda.Text



End Sub

Private Sub PushButton3_Click()
Call fact(2)
End Sub

Private Sub PushButton4_Click()
Call fact(3)
End Sub

Private Sub PushButton5_Click()
Call fact(4)
End Sub

Private Sub PushButton6_Click()
Call fact(5)
End Sub

Private Sub PushButton7_Click()
Call fact(6)
End Sub



Private Sub PushButton8_Click()
Call fact(7)
End Sub

Private Sub PusImprimir_Click()

Dim vmensaje As String

            Unload Mantenimiento
            Load Mantenimiento
             
            MsgBox "Prepare la impresora ...", vbInformation, "Mensaje ..."
            
            
            Unload Mantenimiento
            Load Mantenimiento
                 
            vmensaje = "Contribuyente: " + Trim(UCase(Me.vcontribuyente)) + "  -  "
            
           ' drWsurIntereses.Sections("TituloEmpresa").Controls("ecomentario").Caption = UCase(vcomentario)
            
            If chkMostrarTexto.Value = 1 Then
            
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

Private Sub PusMarcarTodo_Click()
On Error Resume Next
Dim i, j  As Integer

With grilla2
    
For i = 1 To .Rows - 1
     .TextMatrix(i, 0) = "X"
Next

End With
If Err Then Exit Sub
End Sub

Private Sub PusPagarEn_Click()
On Error Resume Next

Dim i, j As Integer
Dim valor, vnrorecibo As String
Dim vsql, vcodigo_barra As String

Dim vacio(50) As String

Dim vaux, vaux2 As String

'recibosApagar() = vacio()

i = 0
j = 0

For i = 0 To 49
    recibosApagar(i) = ""
Next

For i = 0 To grilla2.Rows - 1

    
    
    vnrorecibo = grilla2.TextMatrix(i, 2)
    
    vsql = "select * from recibo_resumen where nro_recibo = '" + vnrorecibo + "'"
    
    vcodigo_barra = traerDatos2(vsql, "cod_barra", ConnComunaDB)
    
    Dim periodo, nrocontri As String
    
    periodo = traerDatos2(vsql, "periodo_anomes", ConnComunaDB)
    nrocontri = traerDatos2(vsql, "id_contribuyentes", ConnComunaDB)
    
    
    valor = vcodigo_barra + vbTab + Me.vcontribuyente.Text + vbTab + grilla2.TextMatrix(i, 4) + vbTab + " " + vbTab + Format(grilla2.TextMatrix(i, 8), "#######0.00") + vbTab + grilla2.TextMatrix(i, 2) + vbTab + periodo + vbTab + nrocontri
    Debug.Print (valor)
    
    If grilla2.TextMatrix(i, 0) = "X" Then
        recibosApagar(j) = valor
        Debug.Print (recibosApagar(j))
        j = j + 1
    Else
        recibosApagar(j) = ""
    End If
Next

' pasa el arreglo que contiene los detalles de pagos a la grilla
frmCargarCodigoBarra.setrecibos recibosApagar()


If Err Then
    frmCargarCodigoBarra.setrecibos recibosApagar()
End If

End Sub

Private Sub PusPersonas_Click()

Me.vcontribuyente.Tag = ""


Me.vcontribuyente.Text = ""

If UCase(LeerXml("Cliente")) = UCase("ComunaWw") Then

   ' Call fbuscarGrilla("personas", "apellido", "id_personas", Me.vcontribuyente.Name, Me, "nombre", True)
    
    If Me.rdUrbano Then
    
        Dim vsql As String
    
        
        ''vsql = " (select p.id_personas, apellido, nombre, t.direccion as direccion, p.direccion as d2  from personas p" + _
        ''" inner join contribuyentes t on t.id_personas = p.id_personas) a  "
     
        '' Call fbuscarGrilla(vsql, "apellido", "id_personas", Me.vcontribuyente.Name, Me, "nombre", True)
         
         
        vsql = " (select t.id_contribuyentes, apellido, nombre, t.direccion as direccion, p.direccion as d2  from personas p" + _
        " inner join contribuyentes t on t.id_personas = p.id_personas) a  "
     
         Call fbuscarGrilla(vsql, "apellido", "id_contribuyentes", Me.vcontribuyente.Name, Me, "nombre", True)
         
         
    End If
    
    
    If Me.rdRural Then
    
    
    
        Dim vsql22 As String
    
       ' vsql22 = " (select p.id_personas, apellido, nombre, t.direccion as direccion, p.direccion as d2  from personas p" + _
        " inner join contribuyentes t on t.id_personas = p.id_personas) a  "
     
        vsql22 = " (select id_contribuyentes, p.id_personas, apellido, nombre, t.direccion as direccion, p.direccion as d2  from personas p" + _
        " inner join contribuyentes t on t.id_personas = p.id_personas) a  "
     
       
         
         Call fbuscarGrilla(vsql22, "apellido", "id_contribuyentes", Me.vcontribuyente.Name, Me, "nombre", True)
    

    End If
    
    
    ' ema:

Else

    vsql22 = "(select c.ID_contribuyente, c.ID_persona, t.Apellido, t.Nombre, c.DireccionCobro " + _
    " from t_personas t  inner join t_contribuyente c on t.id_persona = c.ID_persona) a"

    'Call fbuscarGrilla("t_personas", "Apellido", "ID_Persona", Me.vcontribuyente.Name, Me, "Nombre", True)
     Call fbuscarGrilla(vsql22, "Apellido", "ID_contribuyente", Me.vcontribuyente.Name, Me, "Nombre", True)

    ' ema:
End If

vcp = "p"




End Sub

Private Sub PusReemplazarCampos_Click()
On Error Resume Next


Call fact(1)
Call fact(2)
Call fact(3)
Call fact(4)
Call fact(5)
Call fact(6)
Call fact(7)
Call fact(8)

Dim v As String

If Me.rdUrbano.Value Then
            v = Me.vlibredeuda.Text
            If UCase(LeerXml("Cliente")) = "COMUNAWW" Then
                Me.vlibredeuda.Text = Replace$(v, "JUNCAL", "WHEELWRIGHT")
            Else
                Me.vlibredeuda.Text = Replace$(v, "WHEELWRIGHT", "JUNCAL")
            End If
End If


If Err Then Exit Sub

End Sub

Private Sub rdRural_Click()
vc(6).Text = "Rural"
End Sub

Private Sub rdUrbano_Click()
vc(6).Text = "Urbana"
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
'Call PusPersonas_Click
End Sub

Function getCantDeuda(vid As Long) As Integer
On Error Resume Next

getCantDeuda = 0

Dim vsql, vres As String

vsql = "select count(*) as c  from recibo_resumen t where t.id_estados = 'IM' and id_contribuyentes =" + Str(vid) + " group by t.id_contribuyentes"

vres = traerDatos2(vsql, "c", ConnComunaDB)

getCantDeuda = Val(vres)

If Err < 0 Then
    getCantDeuda = 0
End If

End Function


