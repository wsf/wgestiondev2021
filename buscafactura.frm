VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{9746E3DA-06E1-4D26-9CE4-D9F6411A9C70}#1.0#0"; "SMGA_OcxTxt2008.ocx"
Object = "{FF19AA0C-2968-41B8-A906-E80997A9C394}#208.0#0"; "WSAFIPFEOCX.ocx"
Begin VB.Form frmBuscarFactura 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado de Documentos de Ventas"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   510
   ClientWidth     =   17850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   17850
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   13560
      TabIndex        =   27
      Top             =   3570
      Visible         =   0   'False
      Width           =   2295
      Begin VB.CheckBox chkNoImpreso 
         Caption         =   "No Impreso"
         Height          =   225
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Value           =   1  'Checked
         Width           =   1125
      End
      Begin VB.CheckBox chkImpreso 
         Caption         =   "Impreso"
         Height          =   225
         Left            =   1320
         TabIndex        =   28
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
   End
   Begin XtremeSuiteControls.TabControl Tab 
      Height          =   7605
      Left            =   45
      TabIndex        =   38
      Top             =   570
      Width           =   17685
      _Version        =   851968
      _ExtentX        =   31194
      _ExtentY        =   13414
      _StockProps     =   68
      ItemCount       =   8
      SelectedItem    =   3
      Item(0).Caption =   "Filtar datos"
      Item(0).ControlCount=   12
      Item(0).Control(0)=   "Picture1"
      Item(0).Control(1)=   "GroupBox8"
      Item(0).Control(2)=   "GBTipoMovimiento"
      Item(0).Control(3)=   "Command2"
      Item(0).Control(4)=   "PushButton23"
      Item(0).Control(5)=   "vcodEmpresa"
      Item(0).Control(6)=   "vdescEmpresa"
      Item(0).Control(7)=   "lblDocumento(9)"
      Item(0).Control(8)=   "vDesRepartidor"
      Item(0).Control(9)=   "vcodRepartidor"
      Item(0).Control(10)=   "PushButton24"
      Item(0).Control(11)=   "lblDocumento(10)"
      Item(1).Caption =   "Grupos"
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "GroupBox10"
      Item(1).Control(1)=   "GroupBox11"
      Item(1).Control(2)=   "GroupBox12"
      Item(1).Control(3)=   "GroupBox13"
      Item(2).Caption =   "Tipo Impresión"
      Item(2).ControlCount=   4
      Item(2).Control(0)=   "GroupBox1"
      Item(2).Control(1)=   "ComboBox2"
      Item(2).Control(2)=   "GroupBox7"
      Item(2).Control(3)=   "ComboBox4"
      Item(3).Caption =   "Ver datos"
      Item(3).ControlCount=   19
      Item(3).Control(0)=   "KlexDocumentos"
      Item(3).Control(1)=   "GroupBox14"
      Item(3).Control(2)=   "Label9"
      Item(3).Control(3)=   "bdetalle"
      Item(3).Control(4)=   "Adodc1"
      Item(3).Control(5)=   "PushButton12"
      Item(3).Control(6)=   "PushButton13"
      Item(3).Control(7)=   "TabControl1"
      Item(3).Control(8)=   "vltotal"
      Item(3).Control(9)=   "vlsaldo"
      Item(3).Control(10)=   "vlpagado"
      Item(3).Control(11)=   "barra2"
      Item(3).Control(12)=   "vlblSaldoCtaCte"
      Item(3).Control(13)=   "lblSaldoReal"
      Item(3).Control(14)=   "PusIrA"
      Item(3).Control(15)=   "log2"
      Item(3).Control(16)=   "GroCantidadDe"
      Item(3).Control(17)=   "PusValidarEn"
      Item(3).Control(18)=   "FlaVerDatos"
      Item(4).Caption =   "Retenciones-Percepciones"
      Item(4).Tooltip =   "Liquida retenciones"
      Item(4).ControlCount=   2
      Item(4).Control(0)=   "Label2"
      Item(4).Control(1)=   "MSHFlexGrid1"
      Item(5).Caption =   "Gráficos"
      Item(5).ControlCount=   1
      Item(5).Control(0)=   "PushButton8"
      Item(6).Caption =   "Ínidice"
      Item(6).ControlCount=   1
      Item(6).Control(0)=   "indices"
      Item(7).Caption =   "Conciliar CtasCtes con Facturas"
      Item(7).ControlCount=   3
      Item(7).Control(0)=   "gconciliacion"
      Item(7).Control(1)=   "PusImprimir"
      Item(7).Control(2)=   "lblerrores"
      Begin XtremeSuiteControls.FlatEdit FlaVerDatos 
         Height          =   255
         Left            =   15270
         TabIndex        =   206
         Top             =   360
         Width           =   2280
         _Version        =   851968
         _ExtentX        =   4022
         _ExtentY        =   450
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "Ver datos AFIP (cae y fecha)"
      End
      Begin XtremeSuiteControls.PushButton PusValidarEn 
         Height          =   285
         Left            =   8550
         TabIndex        =   205
         Top             =   6210
         Width           =   3030
         _Version        =   851968
         _ExtentX        =   5345
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Validar en AFIP"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroCantidadDe 
         Height          =   1050
         Left            =   8550
         TabIndex        =   200
         Top             =   6525
         Width           =   3030
         _Version        =   851968
         _ExtentX        =   5345
         _ExtentY        =   1852
         _StockProps     =   79
         Caption         =   "Cantidad de items factura"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.FlatEdit vs7 
            Height          =   285
            Left            =   2295
            TabIndex        =   204
            Top             =   495
            Width           =   645
            _Version        =   851968
            _ExtentX        =   1138
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   255
            Text            =   "4000"
            BackColor       =   255
         End
         Begin XtremeSuiteControls.RadioButton rbdc 
            Height          =   330
            Left            =   135
            TabIndex        =   201
            Top             =   180
            Width           =   2580
            _Version        =   851968
            _ExtentX        =   4551
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Doc. cortos"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbdl 
            Height          =   330
            Left            =   135
            TabIndex        =   202
            Top             =   450
            Width           =   2580
            _Version        =   851968
            _ExtentX        =   4551
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Doc. largo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RadTodos 
            Height          =   240
            Left            =   135
            TabIndex        =   203
            Top             =   765
            Width           =   2580
            _Version        =   851968
            _ExtentX        =   4551
            _ExtentY        =   423
            _StockProps     =   79
            Caption         =   "Todos"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid log2 
         Height          =   3705
         Left            =   480
         TabIndex        =   180
         Top             =   1440
         Width           =   16755
         _ExtentX        =   29554
         _ExtentY        =   6535
         _Version        =   393216
         Cols            =   10
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   10
      End
      Begin XtremeSuiteControls.PushButton PusIrA 
         Height          =   285
         Left            =   15720
         TabIndex        =   176
         Top             =   6600
         Width           =   1785
         _Version        =   851968
         _ExtentX        =   3149
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Ir a Cta Cte"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ProgressBar barra2 
         Height          =   165
         Left            =   90
         TabIndex        =   173
         Top             =   390
         Width           =   9540
         _Version        =   851968
         _ExtentX        =   16828
         _ExtentY        =   291
         _StockProps     =   93
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton PusImprimir 
         Height          =   345
         Left            =   -69820
         TabIndex        =   168
         Top             =   540
         Visible         =   0   'False
         Width           =   2715
         _Version        =   851968
         _ExtentX        =   4789
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Imprimir planilla de conciliación"
         UseVisualStyle  =   -1  'True
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gconciliacion 
         Height          =   6405
         Left            =   -69880
         TabIndex        =   167
         Top             =   1050
         Visible         =   0   'False
         Width           =   17355
         _ExtentX        =   30612
         _ExtentY        =   11298
         _Version        =   393216
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         Height          =   285
         Left            =   -67300
         TabIndex        =   151
         Top             =   810
         Visible         =   0   'False
         Width           =   525
      End
      Begin XtremeSuiteControls.TabControl TabControl1 
         Height          =   1275
         Left            =   60
         TabIndex        =   134
         Top             =   6270
         Width           =   8355
         _Version        =   851968
         _ExtentX        =   14737
         _ExtentY        =   2249
         _StockProps     =   68
         ItemCount       =   2
         Item(0).Caption =   "Pagos de Documentos"
         Item(0).ControlCount=   10
         Item(0).Control(0)=   "PushButton10"
         Item(0).Control(1)=   "barra"
         Item(0).Control(2)=   "Label8"
         Item(0).Control(3)=   "Label7"
         Item(0).Control(4)=   "Label6"
         Item(0).Control(5)=   "Label5"
         Item(0).Control(6)=   "Label3"
         Item(0).Control(7)=   "Label4"
         Item(0).Control(8)=   "vImporteSeleccionado"
         Item(0).Control(9)=   "vtotal"
         Item(1).Caption =   "Factura +"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "GroupBox15"
         Begin XtremeSuiteControls.PushButton PushButton10 
            Height          =   465
            Left            =   3840
            TabIndex        =   135
            ToolTipText     =   "Volver a imputar el pago a los documentos según el importe ingresado."
            Top             =   690
            Width           =   1725
            _Version        =   851968
            _ExtentX        =   3043
            _ExtentY        =   820
            _StockProps     =   79
            Caption         =   "Cancelar Doc."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.ProgressBar barra 
            Height          =   225
            Left            =   60
            TabIndex        =   136
            Top             =   330
            Width           =   7980
            _Version        =   851968
            _ExtentX        =   14076
            _ExtentY        =   397
            _StockProps     =   93
            Text            =   "Barra"
         End
         Begin XtremeSuiteControls.GroupBox GroupBox15 
            Height          =   735
            Left            =   -69940
            TabIndex        =   143
            Top             =   360
            Visible         =   0   'False
            Width           =   9525
            _Version        =   851968
            _ExtentX        =   16801
            _ExtentY        =   1296
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Begin XtremeSuiteControls.CheckBox chksolomarcados 
               Height          =   345
               Left            =   3300
               TabIndex        =   182
               Top             =   240
               Width           =   2655
               _Version        =   851968
               _ExtentX        =   4683
               _ExtentY        =   609
               _StockProps     =   79
               Caption         =   "Modifica solamente los marcados"
               UseVisualStyle  =   -1  'True
            End
            Begin VB.TextBox vncdesde 
               Height          =   285
               Left            =   2235
               TabIndex        =   145
               Top             =   270
               Width           =   915
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Ejecutar"
               Height          =   285
               Left            =   6180
               TabIndex        =   144
               Top             =   270
               Width           =   1695
            End
            Begin VB.Label Label17 
               Caption         =   "Fijar Nro. Comprobante  Inicio:"
               Height          =   255
               Left            =   75
               TabIndex        =   146
               Top             =   330
               Width           =   2895
            End
         End
         Begin VB.Label vtotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000003&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1034
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   2280
            TabIndex        =   148
            Top             =   630
            Width           =   1395
         End
         Begin VB.Label vImporteSeleccionado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000003&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "###,###,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1034
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   2280
            TabIndex        =   147
            Top             =   930
            Width           =   1395
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Importe Total Seleccionado:"
            Height          =   225
            Left            =   30
            TabIndex        =   142
            Top             =   660
            Width           =   2085
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Importe a pagar:"
            Height          =   225
            Left            =   90
            TabIndex        =   141
            Top             =   930
            Width           =   1995
         End
         Begin XtremeSuiteControls.Label Label5 
            Height          =   315
            Left            =   6165
            TabIndex        =   140
            Top             =   600
            Width           =   2085
            _Version        =   851968
            _ExtentX        =   3678
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Selección automática"
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
         Begin XtremeSuiteControls.Label Label6 
            Height          =   225
            Left            =   6165
            TabIndex        =   139
            Top             =   930
            Width           =   2085
            _Version        =   851968
            _ExtentX        =   3678
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "Selección manual"
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
         Begin XtremeSuiteControls.Label Label7 
            Height          =   255
            Left            =   5625
            TabIndex        =   138
            Top             =   930
            Width           =   345
            _Version        =   851968
            _ExtentX        =   609
            _ExtentY        =   450
            _StockProps     =   79
            BackColor       =   65280
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
         Begin XtremeSuiteControls.Label Label8 
            Height          =   315
            Left            =   5625
            TabIndex        =   137
            Top             =   600
            Width           =   345
            _Version        =   851968
            _ExtentX        =   609
            _ExtentY        =   556
            _StockProps     =   79
            BackColor       =   16711680
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
      End
      Begin XtremeSuiteControls.PushButton PushButton12 
         Height          =   285
         Left            =   15240
         TabIndex        =   132
         Top             =   0
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Pinta todo"
         ForeColor       =   0
         BackColor       =   65535
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox14 
         Height          =   465
         Left            =   11730
         TabIndex        =   123
         Top             =   7110
         Width           =   5835
         _Version        =   851968
         _ExtentX        =   10292
         _ExtentY        =   820
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton RadioButton1 
            Height          =   255
            Left            =   120
            TabIndex        =   124
            Top             =   150
            Width           =   885
            _Version        =   851968
            _ExtentX        =   1561
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Todos"
            UseVisualStyle  =   -1  'True
            Appearance      =   6
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RadioButton2 
            Height          =   255
            Left            =   1470
            TabIndex        =   125
            Top             =   150
            Width           =   1815
            _Version        =   851968
            _ExtentX        =   3201
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Adeudados Deudas"
            ForeColor       =   255
            UseVisualStyle  =   -1  'True
            Appearance      =   6
         End
         Begin XtremeSuiteControls.RadioButton RadioButton3 
            Height          =   135
            Left            =   3420
            TabIndex        =   126
            Top             =   210
            Width           =   1275
            _Version        =   851968
            _ExtentX        =   2249
            _ExtentY        =   238
            _StockProps     =   79
            Caption         =   "Pagos"
            ForeColor       =   49152
            UseVisualStyle  =   -1  'True
            Appearance      =   6
         End
      End
      Begin VB.ListBox indices 
         Height          =   5715
         Left            =   -69430
         TabIndex        =   121
         Top             =   900
         Visible         =   0   'False
         Width           =   16605
      End
      Begin XtremeSuiteControls.PushButton PushButton8 
         Height          =   405
         Left            =   -69730
         TabIndex        =   119
         Top             =   7020
         Visible         =   0   'False
         Width           =   2355
         _Version        =   851968
         _ExtentX        =   4154
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Actualizar datos"
         Appearance      =   2
      End
      Begin XtremeSuiteControls.GroupBox GroupBox11 
         Height          =   885
         Left            =   -69790
         TabIndex        =   101
         Top             =   1350
         Visible         =   0   'False
         Width           =   17115
         _Version        =   851968
         _ExtentX        =   30189
         _ExtentY        =   1561
         _StockProps     =   79
         Caption         =   "Filtro:"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.PushButton PushButton6 
            Height          =   315
            Left            =   3420
            TabIndex        =   104
            Top             =   360
            Width           =   465
            _Version        =   851968
            _ExtentX        =   820
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit vCarticulo 
            Height          =   315
            Left            =   1860
            TabIndex        =   103
            Top             =   360
            Width           =   1515
            _Version        =   851968
            _ExtentX        =   2672
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vDarticulo 
            Height          =   315
            Left            =   3960
            TabIndex        =   105
            Top             =   360
            Width           =   8295
            _Version        =   851968
            _ExtentX        =   14631
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   315
            Left            =   180
            TabIndex        =   102
            Top             =   360
            Width           =   1635
            _Version        =   851968
            _ExtentX        =   2884
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "> Artículo / Servicio: "
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox10 
         Height          =   795
         Left            =   -69820
         TabIndex        =   97
         Top             =   540
         Visible         =   0   'False
         Width           =   17085
         _Version        =   851968
         _ExtentX        =   30136
         _ExtentY        =   1402
         _StockProps     =   79
         Caption         =   "Agrupados por:"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton rdArticulo 
            Height          =   375
            Left            =   3870
            TabIndex        =   98
            Top             =   270
            Width           =   1065
            _Version        =   851968
            _ExtentX        =   1879
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Artículo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rdPersona 
            Height          =   375
            Left            =   5700
            TabIndex        =   99
            Top             =   270
            Width           =   1065
            _Version        =   851968
            _ExtentX        =   1879
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Personas"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbSinAgruar 
            Height          =   375
            Left            =   1800
            TabIndex        =   100
            Top             =   270
            Width           =   1755
            _Version        =   851968
            _ExtentX        =   3096
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Sin agrupar"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbMeses 
            Height          =   375
            Left            =   7620
            TabIndex        =   117
            Top             =   270
            Width           =   1065
            _Version        =   851968
            _ExtentX        =   1879
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Meses"
            ForeColor       =   0
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbAnos 
            Height          =   375
            Left            =   9240
            TabIndex        =   118
            Top             =   270
            Width           =   765
            _Version        =   851968
            _ExtentX        =   1349
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Años"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid KlexDocumentos 
         Height          =   5535
         Left            =   150
         TabIndex        =   87
         Top             =   630
         Width           =   17415
         _ExtentX        =   30718
         _ExtentY        =   9763
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   4210688
         BackColorSel    =   16711680
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5325
         Left            =   -69850
         ScaleHeight     =   5325
         ScaleWidth      =   17385
         TabIndex        =   39
         Top             =   780
         Visible         =   0   'False
         Width           =   17385
         Begin XtremeSuiteControls.PushButton PushButton20 
            Height          =   285
            Left            =   6570
            TabIndex        =   188
            Top             =   1470
            Width           =   345
            _Version        =   851968
            _ExtentX        =   609
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit vrubro 
            Height          =   285
            Left            =   6990
            TabIndex        =   184
            Top             =   1470
            Width           =   4065
            _Version        =   851968
            _ExtentX        =   7170
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin WSAFIPFEOCX.WSAFIPFEx vfe 
            Left            =   11370
            Top             =   1110
            _ExtentX        =   1402
            _ExtentY        =   2037
         End
         Begin VB.Frame FraTiposDe 
            Caption         =   "Tipos de documento"
            ForeColor       =   &H00000000&
            Height          =   1005
            Left            =   120
            TabIndex        =   59
            Top             =   3540
            Width           =   14235
            Begin VB.CheckBox chkNotaCC 
               Caption         =   "Nota de Crédito C"
               Height          =   255
               Left            =   9120
               TabIndex        =   191
               Top             =   330
               Value           =   1  'Checked
               Width           =   1750
            End
            Begin VB.CheckBox chkFacturaX 
               Caption         =   "Factura X"
               Height          =   255
               Left            =   7170
               TabIndex        =   166
               Top             =   660
               Value           =   1  'Checked
               Width           =   1215
            End
            Begin XtremeSuiteControls.PushButton PushButton14 
               Height          =   345
               Left            =   12240
               TabIndex        =   149
               Top             =   180
               Width           =   1845
               _Version        =   851968
               _ExtentX        =   3254
               _ExtentY        =   609
               _StockProps     =   79
               Caption         =   "Desmarcar todo"
               UseVisualStyle  =   -1  'True
            End
            Begin VB.CheckBox chkFacturaC 
               Caption         =   "Factura C"
               Height          =   255
               Left            =   5550
               TabIndex        =   88
               Top             =   690
               Value           =   1  'Checked
               Width           =   1215
            End
            Begin VB.CheckBox chkPresupuesto 
               Caption         =   "Presupuesto"
               Height          =   255
               Left            =   5790
               TabIndex        =   68
               Top             =   330
               Value           =   1  'Checked
               Width           =   1335
            End
            Begin VB.CheckBox chkDocNo 
               Caption         =   "Doc. no válido como factura"
               Height          =   255
               Left            =   150
               TabIndex        =   67
               Top             =   690
               Value           =   1  'Checked
               Width           =   2355
            End
            Begin VB.CheckBox chkRemito 
               Caption         =   "Remito"
               Height          =   255
               Left            =   2640
               TabIndex        =   66
               Top             =   690
               Value           =   1  'Checked
               Width           =   885
            End
            Begin VB.CheckBox chkNotasDe 
               Caption         =   "Notas de Débito"
               Height          =   255
               Left            =   1290
               TabIndex        =   65
               Top             =   330
               Value           =   1  'Checked
               Width           =   1455
            End
            Begin VB.CheckBox chkMonotributo 
               Caption         =   "Factura B"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   4650
               TabIndex        =   64
               Top             =   330
               Value           =   1  'Checked
               Width           =   1125
            End
            Begin VB.CheckBox chkFacturaA 
               Caption         =   "Factura A"
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   150
               TabIndex        =   63
               Top             =   330
               Value           =   1  'Checked
               Width           =   1395
            End
            Begin VB.CheckBox chkNotaCA 
               Caption         =   "Notas de Crédito A"
               Height          =   255
               Left            =   2880
               TabIndex        =   62
               Top             =   330
               Value           =   1  'Checked
               Width           =   1750
            End
            Begin VB.CheckBox chkNotaCB 
               Caption         =   "Notas de Crédito B"
               Height          =   255
               Left            =   7170
               TabIndex        =   61
               Top             =   330
               Value           =   1  'Checked
               Width           =   1750
            End
            Begin VB.CheckBox chkOtros 
               Caption         =   "Otros Comprobantes"
               Height          =   255
               Left            =   3660
               TabIndex        =   60
               Top             =   690
               Value           =   1  'Checked
               Width           =   1750
            End
            Begin XtremeSuiteControls.PushButton PushButton15 
               Height          =   345
               Left            =   12240
               TabIndex        =   150
               Top             =   570
               Width           =   1845
               _Version        =   851968
               _ExtentX        =   3254
               _ExtentY        =   609
               _StockProps     =   79
               Caption         =   "Marcar todo"
               UseVisualStyle  =   -1  'True
            End
         End
         Begin XtremeSuiteControls.PushButton PushButton1 
            Height          =   315
            Left            =   5040
            TabIndex        =   40
            Top             =   390
            Width           =   495
            _Version        =   851968
            _ExtentX        =   873
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.GroupBox GroupBox4 
            Height          =   495
            Left            =   150
            TabIndex        =   41
            Top             =   1500
            Width           =   5085
            _Version        =   851968
            _ExtentX        =   8969
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Estado del documento:"
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
            Begin XtremeSuiteControls.ComboBox cmbEstadoDocumento 
               Height          =   315
               Left            =   2010
               TabIndex        =   42
               Top             =   180
               Width           =   2925
               _Version        =   851968
               _ExtentX        =   5159
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
               Text            =   "Todos los estados"
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox3 
            Height          =   525
            Left            =   150
            TabIndex        =   43
            Top             =   2130
            Width           =   4275
            _Version        =   851968
            _ExtentX        =   7541
            _ExtentY        =   926
            _StockProps     =   79
            Caption         =   "Nro. Comprobante:"
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
            Begin XtremeSuiteControls.FlatEdit txtNDesde 
               Height          =   285
               Left            =   840
               TabIndex        =   44
               Top             =   240
               Width           =   1245
               _Version        =   851968
               _ExtentX        =   2196
               _ExtentY        =   503
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.FlatEdit txtNHasta 
               Height          =   285
               Left            =   2880
               TabIndex        =   45
               Top             =   225
               Width           =   1215
               _Version        =   851968
               _ExtentX        =   2143
               _ExtentY        =   503
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin VB.Label Label11 
               Caption         =   "> Desde :"
               Height          =   195
               Left            =   90
               TabIndex        =   47
               Top             =   270
               Width           =   765
            End
            Begin VB.Label Label10 
               Caption         =   "> Hasta :"
               Height          =   195
               Left            =   2130
               TabIndex        =   46
               Top             =   270
               Width           =   705
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox2 
            Height          =   495
            Left            =   150
            TabIndex        =   48
            Top             =   840
            Width           =   5085
            _Version        =   851968
            _ExtentX        =   8969
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Fechas:"
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
            Begin VB.CheckBox chkFechaTodas 
               Caption         =   "Todas"
               Height          =   195
               Left            =   60
               TabIndex        =   49
               Top             =   210
               Value           =   1  'Checked
               Width           =   765
            End
            Begin Aplisoft_CajasDeTexto.TxF dtpFecha 
               Height          =   285
               Index           =   0
               Left            =   1710
               TabIndex        =   50
               Top             =   180
               Width           =   1245
               _ExtentX        =   2196
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
            Begin Aplisoft_CajasDeTexto.TxF dtpFecha 
               Height          =   285
               Index           =   1
               Left            =   3780
               TabIndex        =   51
               Top             =   180
               Width           =   1305
               _ExtentX        =   2302
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
            Begin VB.Label lblDocumento 
               Caption         =   "> Hasta:"
               Height          =   195
               Index           =   3
               Left            =   3030
               TabIndex        =   53
               Top             =   240
               Width           =   615
            End
            Begin VB.Label lblDocumento 
               Caption         =   "> Desde:"
               Height          =   195
               Index           =   2
               Left            =   990
               TabIndex        =   52
               Top             =   210
               Width           =   705
            End
         End
         Begin XtremeSuiteControls.GroupBox GroAgrupadoPor 
            Height          =   525
            Left            =   120
            TabIndex        =   54
            Top             =   4680
            Width           =   5985
            _Version        =   851968
            _ExtentX        =   10557
            _ExtentY        =   926
            _StockProps     =   79
            Caption         =   "Agrupado por:"
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
            Begin XtremeSuiteControls.RadioButton RadNoAgrupado 
               Height          =   225
               Left            =   1260
               TabIndex        =   55
               Top             =   180
               Width           =   1365
               _Version        =   851968
               _ExtentX        =   2408
               _ExtentY        =   397
               _StockProps     =   79
               Caption         =   "No agrupado"
               UseVisualStyle  =   -1  'True
               Value           =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RadProvincias 
               Height          =   225
               Left            =   2835
               TabIndex        =   56
               Top             =   180
               Width           =   1365
               _Version        =   851968
               _ExtentX        =   2408
               _ExtentY        =   397
               _StockProps     =   79
               Caption         =   "Provincias"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RadPersona 
               Height          =   225
               Left            =   4185
               TabIndex        =   57
               Top             =   180
               Width           =   1665
               _Version        =   851968
               _ExtentX        =   2937
               _ExtentY        =   397
               _StockProps     =   79
               Caption         =   "Cliente o Proveedor"
               UseVisualStyle  =   -1  'True
            End
         End
         Begin XtremeSuiteControls.FlatEdit txtCliente 
            Height          =   315
            Left            =   3120
            TabIndex        =   58
            Top             =   0
            Width           =   11775
            _Version        =   851968
            _ExtentX        =   20770
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit txtEmpleado 
            Height          =   315
            Left            =   1260
            TabIndex        =   69
            Top             =   390
            Width           =   3735
            _Version        =   851968
            _ExtentX        =   6588
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit1 
            Height          =   315
            Left            =   6450
            TabIndex        =   70
            Top             =   390
            Width           =   4095
            _Version        =   851968
            _ExtentX        =   7223
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.PushButton PushButton2 
            Height          =   315
            Left            =   10620
            TabIndex        =   71
            Top             =   390
            Width           =   405
            _Version        =   851968
            _ExtentX        =   714
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.GroupBox GroupBox5 
            Height          =   495
            Left            =   120
            TabIndex        =   72
            Top             =   2880
            Width           =   11025
            _Version        =   851968
            _ExtentX        =   19447
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Tipos de pedidos:"
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
            Begin XtremeSuiteControls.PushButton ref 
               Height          =   285
               Left            =   10470
               TabIndex        =   164
               Top             =   210
               Width           =   495
               _Version        =   851968
               _ExtentX        =   873
               _ExtentY        =   503
               _StockProps     =   79
               Caption         =   "..."
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.ComboBox cmbTiposPedidos 
               Height          =   315
               Left            =   1770
               TabIndex        =   73
               Top             =   180
               Width           =   2865
               _Version        =   851968
               _ExtentX        =   5054
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
               Text            =   "Todos"
            End
            Begin XtremeSuiteControls.FlatEdit vreferenciaPedido 
               Height          =   285
               Left            =   6660
               TabIndex        =   129
               Top             =   210
               Width           =   3735
               _Version        =   851968
               _ExtentX        =   6588
               _ExtentY        =   503
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               Caption         =   "Referencia de pedidos exportados:"
               Height          =   435
               Left            =   0
               TabIndex        =   131
               Top             =   0
               Width           =   1665
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               Caption         =   "Referencia de pedidos exportados:"
               ForeColor       =   &H00C00000&
               Height          =   435
               Left            =   4770
               TabIndex        =   128
               Top             =   120
               Width           =   1665
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox9 
            Height          =   495
            Left            =   5490
            TabIndex        =   91
            Top             =   870
            Width           =   5535
            _Version        =   851968
            _ExtentX        =   9763
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Fecha Vencimiento:"
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
            Begin VB.CheckBox chkFechaVencimiento 
               Caption         =   "Todas"
               Height          =   195
               Left            =   0
               TabIndex        =   92
               Top             =   210
               Value           =   1  'Checked
               Width           =   765
            End
            Begin Aplisoft_CajasDeTexto.TxF dtpFecha 
               Height          =   285
               Index           =   2
               Left            =   2280
               TabIndex        =   93
               Top             =   180
               Width           =   1245
               _ExtentX        =   2196
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
            Begin Aplisoft_CajasDeTexto.TxF dtpFecha 
               Height          =   285
               Index           =   3
               Left            =   4200
               TabIndex        =   94
               Top             =   180
               Width           =   1305
               _ExtentX        =   2302
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
            Begin VB.Label lblDocumento 
               Caption         =   "> Desde:"
               Height          =   195
               Index           =   6
               Left            =   1530
               TabIndex        =   96
               Top             =   210
               Width           =   705
            End
            Begin VB.Label lblDocumento 
               Caption         =   "> Hasta:"
               Height          =   195
               Index           =   5
               Left            =   3570
               TabIndex        =   95
               Top             =   240
               Width           =   615
            End
         End
         Begin XtremeSuiteControls.FlatEdit txtcodigoCliente 
            Height          =   315
            Left            =   1260
            TabIndex        =   152
            Top             =   0
            Width           =   1245
            _Version        =   851968
            _ExtentX        =   2196
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.GroupBox GroupBox16 
            Height          =   495
            Left            =   5520
            TabIndex        =   153
            Top             =   2160
            Width           =   5535
            _Version        =   851968
            _ExtentX        =   9763
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Fecha de Pago estimada"
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
            Begin VB.CheckBox chkFechaPago 
               Caption         =   "Todas"
               Height          =   195
               Left            =   0
               TabIndex        =   154
               Top             =   210
               Value           =   1  'Checked
               Width           =   765
            End
            Begin Aplisoft_CajasDeTexto.TxF dtpFecha 
               Height          =   285
               Index           =   4
               Left            =   2280
               TabIndex        =   155
               Top             =   180
               Width           =   1245
               _ExtentX        =   2196
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
            Begin Aplisoft_CajasDeTexto.TxF dtpFecha 
               Height          =   285
               Index           =   5
               Left            =   4200
               TabIndex        =   156
               Top             =   180
               Width           =   1305
               _ExtentX        =   2302
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
            Begin VB.Label lblDocumento 
               Caption         =   "> Hasta:"
               Height          =   195
               Index           =   8
               Left            =   3570
               TabIndex        =   158
               Top             =   240
               Width           =   615
            End
            Begin VB.Label lblDocumento 
               Caption         =   "> Desde:"
               Height          =   195
               Index           =   7
               Left            =   1530
               TabIndex        =   157
               Top             =   210
               Width           =   705
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox17 
            Height          =   525
            Left            =   6450
            TabIndex        =   159
            Top             =   4680
            Width           =   4725
            _Version        =   851968
            _ExtentX        =   8334
            _ExtentY        =   926
            _StockProps     =   79
            Caption         =   "Situación del saldo de los documentos:"
            ForeColor       =   255
            UseVisualStyle  =   -1  'True
            Begin XtremeSuiteControls.RadioButton rbPago 
               Height          =   225
               Left            =   1140
               TabIndex        =   160
               Top             =   240
               Width           =   795
               _Version        =   851968
               _ExtentX        =   1402
               _ExtentY        =   397
               _StockProps     =   79
               Caption         =   "Pagos"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton rbAdeudado 
               Height          =   225
               Left            =   1950
               TabIndex        =   161
               Top             =   240
               Width           =   1095
               _Version        =   851968
               _ExtentX        =   1931
               _ExtentY        =   397
               _StockProps     =   79
               Caption         =   "Adeudado"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton rbParcial 
               Height          =   225
               Left            =   3030
               TabIndex        =   162
               Top             =   210
               Width           =   1665
               _Version        =   851968
               _ExtentX        =   2937
               _ExtentY        =   397
               _StockProps     =   79
               Caption         =   "Parcialmente pago"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton rbTodos 
               Height          =   225
               Left            =   90
               TabIndex        =   163
               Top             =   240
               Width           =   885
               _Version        =   851968
               _ExtentX        =   1561
               _ExtentY        =   397
               _StockProps     =   79
               Caption         =   "Todos"
               UseVisualStyle  =   -1  'True
               Value           =   -1  'True
            End
         End
         Begin XtremeSuiteControls.FlatEdit vsubrubro 
            Height          =   285
            Left            =   6990
            TabIndex        =   185
            Top             =   1800
            Width           =   4065
            _Version        =   851968
            _ExtentX        =   7170
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.PushButton PushButton21 
            Height          =   285
            Left            =   6570
            TabIndex        =   189
            Top             =   1800
            Width           =   345
            _Version        =   851968
            _ExtentX        =   609
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label19 
            Height          =   285
            Left            =   5340
            TabIndex        =   187
            Top             =   1800
            Width           =   1095
            _Version        =   851968
            _ExtentX        =   1931
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Sub Rubro:"
            ForeColor       =   4194368
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
         Begin XtremeSuiteControls.Label Label18 
            Height          =   285
            Left            =   5310
            TabIndex        =   186
            Top             =   1470
            Width           =   1155
            _Version        =   851968
            _ExtentX        =   2037
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Rubro:"
            ForeColor       =   4194368
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
         Begin VB.Label lblDocumento 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "> Persona :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   76
            Top             =   60
            Width           =   945
         End
         Begin VB.Label lblDocumento 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "> Interesado: "
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   -720
            TabIndex        =   75
            Top             =   420
            Width           =   1755
         End
         Begin VB.Label lblDocumento 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "> Chofer:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   5280
            TabIndex        =   74
            Top             =   450
            Width           =   1035
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   765
         Left            =   -69850
         TabIndex        =   79
         Top             =   720
         Visible         =   0   'False
         Width           =   12705
         _Version        =   851968
         _ExtentX        =   22410
         _ExtentY        =   1349
         _StockProps     =   79
         Caption         =   "Tipos de pedidos:"
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
         Begin XtremeSuiteControls.ComboBox ComboBox1 
            Height          =   315
            Left            =   270
            TabIndex        =   80
            Top             =   270
            Width           =   12105
            _Version        =   851968
            _ExtentX        =   21352
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "Todos"
         End
      End
      Begin XtremeSuiteControls.ComboBox ComboBox2 
         Height          =   315
         Left            =   -69790
         TabIndex        =   81
         Top             =   960
         Visible         =   0   'False
         Width           =   12105
         _Version        =   851968
         _ExtentX        =   21352
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "Todos"
      End
      Begin XtremeSuiteControls.GroupBox GroupBox7 
         Height          =   825
         Left            =   -69760
         TabIndex        =   82
         Top             =   1770
         Visible         =   0   'False
         Width           =   12645
         _Version        =   851968
         _ExtentX        =   22304
         _ExtentY        =   1455
         _StockProps     =   79
         Caption         =   "Tipos de pedidos:"
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
         Begin XtremeSuiteControls.ComboBox ComboBox3 
            Height          =   315
            Left            =   180
            TabIndex        =   83
            Top             =   270
            Width           =   12105
            _Version        =   851968
            _ExtentX        =   21352
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "Todos"
         End
      End
      Begin XtremeSuiteControls.ComboBox ComboBox4 
         Height          =   315
         Left            =   -69790
         TabIndex        =   84
         Top             =   960
         Visible         =   0   'False
         Width           =   12105
         _Version        =   851968
         _ExtentX        =   21352
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "Todos"
      End
      Begin XtremeSuiteControls.GroupBox GroupBox8 
         Height          =   555
         Left            =   -69790
         TabIndex        =   89
         Top             =   6150
         Visible         =   0   'False
         Width           =   3615
         _Version        =   851968
         _ExtentX        =   6376
         _ExtentY        =   979
         _StockProps     =   79
         Caption         =   "Ordenar el listado por el campo:"
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
         Begin XtremeSuiteControls.ComboBox vordenadoPor 
            Height          =   315
            Left            =   0
            TabIndex        =   90
            Top             =   210
            Width           =   3585
            _Version        =   851968
            _ExtentX        =   6324
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "Fecha"
         End
      End
      Begin XtremeSuiteControls.GroupBox GBTipoMovimiento 
         Height          =   555
         Left            =   -69940
         TabIndex        =   106
         Top             =   6810
         Visible         =   0   'False
         Width           =   17595
         _Version        =   851968
         _ExtentX        =   31036
         _ExtentY        =   979
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton cmdFiltrar 
            Height          =   375
            Left            =   60
            TabIndex        =   107
            Top             =   120
            Width           =   17505
            _Version        =   851968
            _ExtentX        =   30877
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Buscar"
            UseVisualStyle  =   -1  'True
            Picture         =   "buscafactura.frx":0000
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox12 
         Height          =   735
         Left            =   -69970
         TabIndex        =   108
         Top             =   6840
         Visible         =   0   'False
         Width           =   17595
         _Version        =   851968
         _ExtentX        =   31036
         _ExtentY        =   1296
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton PushButton7 
            Height          =   435
            Left            =   90
            TabIndex        =   109
            Top             =   210
            Width           =   17535
            _Version        =   851968
            _ExtentX        =   30930
            _ExtentY        =   767
            _StockProps     =   79
            Caption         =   "Buscar"
            UseVisualStyle  =   -1  'True
            Picture         =   "buscafactura.frx":059A
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox13 
         Height          =   795
         Left            =   -69790
         TabIndex        =   110
         Top             =   2430
         Visible         =   0   'False
         Width           =   17085
         _Version        =   851968
         _ExtentX        =   30136
         _ExtentY        =   1402
         _StockProps     =   79
         Caption         =   "Ordernar el listado por:"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton rbOPersona 
            Height          =   375
            Left            =   3810
            TabIndex        =   111
            Top             =   270
            Width           =   1065
            _Version        =   851968
            _ExtentX        =   1879
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Personas"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbOCantidad 
            Height          =   375
            Left            =   5700
            TabIndex        =   112
            Top             =   270
            Width           =   2745
            _Version        =   851968
            _ExtentX        =   4842
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Cantidad de unidades vendidas"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbOArticulo 
            Height          =   375
            Left            =   2010
            TabIndex        =   113
            Top             =   240
            Width           =   1755
            _Version        =   851968
            _ExtentX        =   3096
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Articulo"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbOTotales 
            Height          =   375
            Left            =   8790
            TabIndex        =   114
            Top             =   270
            Width           =   2745
            _Version        =   851968
            _ExtentX        =   4842
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Importe de las Ventas/Compras"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6585
         Left            =   -69880
         TabIndex        =   116
         Top             =   960
         Visible         =   0   'False
         Width           =   17415
         _ExtentX        =   30718
         _ExtentY        =   11615
         _Version        =   393216
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   10380
         Top             =   5820
         Width           =   2505
         _ExtentX        =   4419
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
         Caption         =   "bfactura"
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
      Begin MSAdodcLib.Adodc bdetalle 
         Height          =   330
         Left            =   14190
         Top             =   5730
         Width           =   2505
         _ExtentX        =   4419
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
         Caption         =   "bdetalle"
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
      Begin XtremeSuiteControls.PushButton PushButton13 
         Height          =   285
         Left            =   16320
         TabIndex        =   133
         Top             =   0
         Width           =   1185
         _Version        =   851968
         _ExtentX        =   2090
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Despintar todo"
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton23 
         Height          =   345
         Left            =   -67300
         TabIndex        =   192
         Top             =   390
         Visible         =   0   'False
         Width           =   525
         _Version        =   851968
         _ExtentX        =   926
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "< F6>"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit vcodEmpresa 
         Height          =   315
         Left            =   -68590
         TabIndex        =   193
         Top             =   390
         Visible         =   0   'False
         Width           =   1245
         _Version        =   851968
         _ExtentX        =   2196
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4210752
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit vdescEmpresa 
         Height          =   345
         Left            =   -66700
         TabIndex        =   194
         Top             =   360
         Visible         =   0   'False
         Width           =   3975
         _Version        =   851968
         _ExtentX        =   7011
         _ExtentY        =   609
         _StockProps     =   77
         ForeColor       =   4210752
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.FlatEdit vDesRepartidor 
         Height          =   315
         Left            =   -59950
         TabIndex        =   196
         Top             =   390
         Visible         =   0   'False
         Width           =   4965
         _Version        =   851968
         _ExtentX        =   8758
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4210752
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vcodRepartidor 
         Height          =   315
         Left            =   -61360
         TabIndex        =   197
         Top             =   390
         Visible         =   0   'False
         Width           =   795
         _Version        =   851968
         _ExtentX        =   1402
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4210752
         BackColor       =   16777215
         BackColor       =   16777215
      End
      Begin XtremeSuiteControls.PushButton PushButton24 
         Height          =   315
         Left            =   -60520
         TabIndex        =   198
         Top             =   390
         Visible         =   0   'False
         Width           =   495
         _Version        =   851968
         _ExtentX        =   873
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "<F7>"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label lblDocumento 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "> Vendedor :"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   10
         Left            =   -62500
         TabIndex        =   120
         Top             =   450
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lblDocumento 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "> Empresa  :"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   9
         Left            =   -69700
         TabIndex        =   195
         Top             =   420
         Visible         =   0   'False
         Width           =   945
      End
      Begin XtremeSuiteControls.Label lblSaldoReal 
         Height          =   225
         Left            =   11850
         TabIndex        =   175
         Top             =   6630
         Width           =   1845
         _Version        =   851968
         _ExtentX        =   3254
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Saldo real Ctas Ctes:"
      End
      Begin VB.Label vlblSaldoCtaCte 
         Appearance      =   0  'Flat
         BackColor       =   &H00747474&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   13740
         TabIndex        =   174
         Top             =   6570
         Width           =   1785
      End
      Begin VB.Label vlpagado 
         Appearance      =   0  'Flat
         BackColor       =   &H00747474&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   13740
         TabIndex        =   172
         Top             =   6210
         Width           =   1785
      End
      Begin VB.Label vlsaldo 
         Appearance      =   0  'Flat
         BackColor       =   &H00747474&
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   15720
         TabIndex        =   171
         Top             =   6210
         Width           =   1785
      End
      Begin VB.Label vltotal 
         Appearance      =   0  'Flat
         BackColor       =   &H00747474&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   11820
         TabIndex        =   170
         Top             =   6210
         Width           =   1785
      End
      Begin XtremeSuiteControls.Label lblerrores 
         Height          =   315
         Left            =   -66460
         TabIndex        =   169
         Top             =   600
         Visible         =   0   'False
         Width           =   13725
         _Version        =   851968
         _ExtentX        =   24209
         _ExtentY        =   556
         _StockProps     =   79
         ForeColor       =   255
      End
      Begin XtremeSuiteControls.Label Label9 
         Height          =   225
         Left            =   11670
         TabIndex        =   127
         Top             =   6900
         Width           =   5715
         _Version        =   851968
         _ExtentX        =   10081
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Puede selección documentos a pagar haciendo clic en las filas correspondiente"
         ForeColor       =   8421504
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   435
         Left            =   -69850
         TabIndex        =   115
         Top             =   480
         Visible         =   0   'False
         Width           =   13725
         _Version        =   851968
         _ExtentX        =   24209
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Espacio para el listado de las Retenciones y Persepciones"
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox6 
      Height          =   585
      Left            =   0
      TabIndex        =   30
      Top             =   -180
      Width           =   17775
      _Version        =   851968
      _ExtentX        =   31353
      _ExtentY        =   1032
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton PusDetalles 
         Height          =   375
         Index           =   2
         Left            =   6480
         TabIndex        =   199
         Top             =   180
         Width           =   1080
         _Version        =   851968
         _ExtentX        =   1905
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Detalles"
         UseVisualStyle  =   -1  'True
         Picture         =   "buscafactura.frx":0B34
      End
      Begin XtremeSuiteControls.PushButton PushButton22 
         Height          =   375
         Left            =   12510
         TabIndex        =   190
         Top             =   180
         Width           =   405
         _Version        =   851968
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "FE+"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton19 
         Height          =   375
         Left            =   15630
         TabIndex        =   181
         Top             =   180
         Width           =   405
         _Version        =   851968
         _ExtentX        =   714
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Log"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton16 
         Height          =   375
         Left            =   12990
         TabIndex        =   178
         Top             =   180
         Width           =   975
         _Version        =   851968
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Marcar todo"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PusExell 
         Height          =   375
         Left            =   4980
         TabIndex        =   177
         Top             =   180
         Width           =   525
         _Version        =   851968
         _ExtentX        =   926
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Excel"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton11 
         Height          =   375
         Left            =   3960
         TabIndex        =   130
         Top             =   180
         Width           =   1035
         _Version        =   851968
         _ExtentX        =   1826
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Facturar +"
         ForeColor       =   255
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
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton9 
         Height          =   375
         Left            =   60
         TabIndex        =   122
         Top             =   180
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Volver"
         UseVisualStyle  =   -1  'True
         Picture         =   "buscafactura.frx":0F11
      End
      Begin XtremeSuiteControls.PushButton PushButton5 
         Height          =   375
         Left            =   16050
         TabIndex        =   86
         Top             =   180
         Width           =   825
         _Version        =   851968
         _ExtentX        =   1455
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Pintar"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton CmdVerDetalle 
         Height          =   375
         Left            =   1980
         TabIndex        =   31
         Top             =   180
         Width           =   1155
         _Version        =   851968
         _ExtentX        =   2037
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Ver detalle"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton CmdBorrar 
         Height          =   375
         Left            =   3150
         TabIndex        =   32
         Top             =   180
         Width           =   825
         _Version        =   851968
         _ExtentX        =   1455
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Borrar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdImprimir 
         Height          =   375
         Index           =   0
         Left            =   5520
         TabIndex        =   33
         Top             =   180
         Width           =   945
         _Version        =   851968
         _ExtentX        =   1667
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
         Picture         =   "buscafactura.frx":1318
      End
      Begin XtremeSuiteControls.PushButton btn_Recambio 
         Height          =   375
         Left            =   930
         TabIndex        =   37
         Top             =   180
         Visible         =   0   'False
         Width           =   975
         _Version        =   851968
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Volquete"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PusCerrar 
         Height          =   375
         Index           =   1
         Left            =   16890
         TabIndex        =   77
         Top             =   180
         Width           =   825
         _Version        =   851968
         _ExtentX        =   1455
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cerrar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   375
         Left            =   7695
         TabIndex        =   78
         Top             =   180
         Width           =   1380
         _Version        =   851968
         _ExtentX        =   2434
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cambiar estado del documento"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   375
         Left            =   9135
         TabIndex        =   85
         Top             =   180
         Visible         =   0   'False
         Width           =   1035
         _Version        =   851968
         _ExtentX        =   1826
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Anular Documento"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdImprimir 
         Height          =   375
         Index           =   1
         Left            =   6270
         TabIndex        =   36
         Top             =   180
         Width           =   975
         _Version        =   851968
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Factura"
         Enabled         =   0   'False
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton CmdEjecutarCobro 
         Height          =   375
         Left            =   10110
         TabIndex        =   35
         Top             =   180
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Ejec. Cobro"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdAnularFactura 
         Height          =   375
         Left            =   11370
         TabIndex        =   34
         Top             =   180
         Width           =   435
         _Version        =   851968
         _ExtentX        =   767
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Anular Doc."
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton17 
         Height          =   375
         Left            =   13980
         TabIndex        =   179
         Top             =   180
         Width           =   1245
         _Version        =   851968
         _ExtentX        =   2196
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Desmarcar todo"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton18 
         Height          =   375
         Left            =   11820
         TabIndex        =   183
         Top             =   180
         Width           =   645
         _Version        =   851968
         _ExtentX        =   1138
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "WS FE"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin MSAdodcLib.Adodc bfactura 
      Height          =   630
      Left            =   10440
      Top             =   7980
      Visible         =   0   'False
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   1111
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
      Caption         =   "bfactura"
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
   Begin MSAdodcLib.Adodc bfacturas_impagas 
      Height          =   330
      Left            =   7920
      Top             =   8280
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
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
      Caption         =   "bfacturas_impagas"
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
   Begin TabDlg.SSTab TabDocumentos 
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   2790
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   1085
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Documentos en Gral."
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Documentos Impagos"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label13"
      Tab(1).Control(1)=   "Label14"
      Tab(1).Control(2)=   "Label15"
      Tab(1).Control(3)=   "DataGrid1"
      Tab(1).Control(4)=   "txtRepartoImpagos"
      Tab(1).Control(5)=   "cmdImprimeFD"
      Tab(1).Control(6)=   "txtClienteImpagos"
      Tab(1).Control(7)=   "cmdImprimeF"
      Tab(1).Control(8)=   "cmdImprimeD"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Documentos Pagos"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DataGrid2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Totales"
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "GroT"
      Tab(3).ControlCount=   1
      Begin VB.CommandButton cmdImprimeD 
         Caption         =   "Imprimir Doc. Impagos"
         Height          =   495
         Left            =   -63690
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Imprimir"
         Top             =   4560
         UseMaskColor    =   -1  'True
         Width           =   1755
      End
      Begin VB.CommandButton cmdImprimeF 
         Caption         =   "Imprimir Fact. A Impagas"
         Height          =   495
         Left            =   -65550
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Imprimir"
         Top             =   4560
         UseMaskColor    =   -1  'True
         Width           =   1875
      End
      Begin VB.TextBox txtClienteImpagos 
         Height          =   285
         Left            =   -72360
         TabIndex        =   6
         Top             =   4650
         Width           =   4035
      End
      Begin VB.CommandButton cmdImprimeFD 
         Caption         =   "Imprimir Doc/Fact Impagos"
         Height          =   495
         Left            =   -67740
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Imprimir"
         Top             =   4560
         UseMaskColor    =   -1  'True
         Width           =   2205
      End
      Begin VB.TextBox txtRepartoImpagos 
         Height          =   285
         Left            =   -74790
         TabIndex        =   4
         Top             =   4650
         Width           =   2115
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2595
         Left            =   -74880
         TabIndex        =   1
         Top             =   420
         Width           =   13035
         _ExtentX        =   22992
         _ExtentY        =   4577
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Appearance      =   0
         BackColor       =   16777215
         HeadLines       =   2
         RowHeight       =   15
         FormatLocked    =   -1  'True
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "ÚltimoDeFecha"
            Caption         =   "ÚltimoDeFecha"
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
            DataField       =   "ÚltimoDeNcomprobante"
            Caption         =   "ÚltimoDeNcomprobante"
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
         BeginProperty Column02 
            DataField       =   "ÚltimoDeTotal"
            Caption         =   "ÚltimoDeTotal"
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
         BeginProperty Column03 
            DataField       =   "ÚltimoDePagado"
            Caption         =   "ÚltimoDePagado"
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
         BeginProperty Column04 
            DataField       =   "Remito"
            Caption         =   "Remito"
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
         BeginProperty Column05 
            DataField       =   "ÚltimoDeCodigo"
            Caption         =   "ÚltimoDeCodigo"
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
         BeginProperty Column06 
            DataField       =   "ÚltimoDeNombre"
            Caption         =   "ÚltimoDeNombre"
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
         BeginProperty Column07 
            DataField       =   "ÚltimoDereparto"
            Caption         =   "ÚltimoDereparto"
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
         BeginProperty Column08 
            DataField       =   "ÚltimoDetipo"
            Caption         =   "ÚltimoDetipo"
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   4635
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   13035
         _ExtentX        =   22992
         _ExtentY        =   8176
         _Version        =   393216
         AllowUpdate     =   -1  'True
         Appearance      =   0
         BackColor       =   16777215
         HeadLines       =   2
         RowHeight       =   15
         FormatLocked    =   -1  'True
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
         ColumnCount     =   27
         BeginProperty Column00 
            DataField       =   "Ncomprobante"
            Caption         =   "Nº C."
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
            DataField       =   "Fecha"
            Caption         =   "Fecha"
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
         BeginProperty Column02 
            DataField       =   "Hora"
            Caption         =   "Hora"
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
         BeginProperty Column03 
            DataField       =   "Codigo"
            Caption         =   "Codigo"
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
         BeginProperty Column04 
            DataField       =   "Nombre"
            Caption         =   "Nombre"
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
         BeginProperty Column05 
            DataField       =   "Domicilio"
            Caption         =   "Domicilio"
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
         BeginProperty Column06 
            DataField       =   "Cod_repartidor"
            Caption         =   "C. Rep."
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
         BeginProperty Column07 
            DataField       =   "Repartidor"
            Caption         =   "Repartidor"
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
         BeginProperty Column08 
            DataField       =   "Localidad"
            Caption         =   "Localidad"
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
         BeginProperty Column09 
            DataField       =   "Telefono"
            Caption         =   "Telefono"
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
         BeginProperty Column10 
            DataField       =   "Iva"
            Caption         =   "Iva"
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
         BeginProperty Column11 
            DataField       =   "Tiva2"
            Caption         =   "Tiva2"
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
         BeginProperty Column12 
            DataField       =   "Cventa"
            Caption         =   "C: Venta"
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
         BeginProperty Column13 
            DataField       =   "Remito"
            Caption         =   "Remito"
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
         BeginProperty Column14 
            DataField       =   "Subtotal"
            Caption         =   "Subtotal"
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
         BeginProperty Column15 
            DataField       =   "Tiva"
            Caption         =   "Tiva"
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
         BeginProperty Column16 
            DataField       =   "Total"
            Caption         =   "Total cdo."
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
         BeginProperty Column17 
            DataField       =   "Total_ctacte"
            Caption         =   "Total ctacte"
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
         BeginProperty Column18 
            DataField       =   "Comentario"
            Caption         =   "Comentario"
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
         BeginProperty Column19 
            DataField       =   "Descuento"
            Caption         =   "Descuento"
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
         BeginProperty Column20 
            DataField       =   "Impuesto"
            Caption         =   "Impuesto"
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
         BeginProperty Column21 
            DataField       =   "cuit"
            Caption         =   "cuit"
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
         BeginProperty Column22 
            DataField       =   "Impreso"
            Caption         =   "Impreso"
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
         BeginProperty Column23 
            DataField       =   "tipo"
            Caption         =   "tipo"
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
         BeginProperty Column24 
            DataField       =   "id"
            Caption         =   "id"
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
         BeginProperty Column25 
            DataField       =   "vvremito"
            Caption         =   "vvremito"
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
         BeginProperty Column26 
            DataField       =   "codigo_num"
            Caption         =   "codigo_num"
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
            BeginProperty Column09 
            EndProperty
            BeginProperty Column10 
            EndProperty
            BeginProperty Column11 
            EndProperty
            BeginProperty Column12 
            EndProperty
            BeginProperty Column13 
            EndProperty
            BeginProperty Column14 
            EndProperty
            BeginProperty Column15 
            EndProperty
            BeginProperty Column16 
            EndProperty
            BeginProperty Column17 
            EndProperty
            BeginProperty Column18 
            EndProperty
            BeginProperty Column19 
            EndProperty
            BeginProperty Column20 
            EndProperty
            BeginProperty Column21 
            EndProperty
            BeginProperty Column22 
            EndProperty
            BeginProperty Column23 
            EndProperty
            BeginProperty Column24 
            EndProperty
            BeginProperty Column25 
            EndProperty
            BeginProperty Column26 
            EndProperty
         EndProperty
      End
      Begin XtremeSuiteControls.GroupBox GroT 
         Height          =   2295
         Left            =   -74760
         TabIndex        =   11
         Top             =   600
         Width           =   4575
         _Version        =   851968
         _ExtentX        =   8070
         _ExtentY        =   4048
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin VB.Label lblTotal 
            Alignment       =   2  'Center
            Caption         =   "Total"
            ForeColor       =   &H00000080&
            Height          =   225
            Left            =   3240
            TabIndex        =   26
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblIV 
            Alignment       =   2  'Center
            Caption         =   "I.V.A."
            ForeColor       =   &H00000080&
            Height          =   225
            Left            =   1680
            TabIndex        =   25
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblTiposDe 
            Alignment       =   2  'Center
            Caption         =   "Tipos de Doc."
            ForeColor       =   &H00000080&
            Height          =   225
            Left            =   240
            TabIndex        =   24
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label ind 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1770
            TabIndex        =   23
            Top             =   1680
            Width           =   1155
         End
         Begin VB.Label inc 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1770
            TabIndex        =   22
            Top             =   1320
            Width           =   1155
         End
         Begin VB.Label imono 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1770
            TabIndex        =   21
            Top             =   990
            Width           =   1155
         End
         Begin VB.Label ifactaLabel 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1770
            TabIndex        =   20
            Top             =   660
            Width           =   1155
         End
         Begin VB.Label tnd 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3000
            TabIndex        =   19
            Top             =   1680
            Width           =   1155
         End
         Begin VB.Label tnc 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3000
            TabIndex        =   18
            Top             =   1320
            Width           =   1155
         End
         Begin VB.Label tmono 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3000
            TabIndex        =   17
            Top             =   990
            Width           =   1155
         End
         Begin VB.Label tfacta 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   3000
            TabIndex        =   16
            Top             =   660
            Width           =   1155
         End
         Begin VB.Label lblNotaDeb 
            Caption         =   "Nota de Débito :"
            Height          =   255
            Left            =   270
            TabIndex        =   15
            Top             =   1710
            Width           =   1305
         End
         Begin VB.Label lblNotaDe 
            Caption         =   "Nota de Crédito :"
            Height          =   255
            Left            =   270
            TabIndex        =   14
            Top             =   1350
            Width           =   1245
         End
         Begin VB.Label lblMonotributo 
            Caption         =   "Monotributo :"
            Height          =   255
            Left            =   270
            TabIndex        =   13
            Top             =   1020
            Width           =   1155
         End
         Begin VB.Label lblFacturaA 
            Caption         =   "Factura A:"
            Height          =   255
            Left            =   270
            TabIndex        =   12
            Top             =   690
            Width           =   1155
         End
      End
      Begin VB.Label Label15 
         Caption         =   "> Impresión de documentos de ventas impagos por tipos (Remitos y Facturas)"
         Height          =   315
         Left            =   -67740
         TabIndex        =   10
         Top             =   4290
         Width           =   5655
      End
      Begin VB.Label Label14 
         Caption         =   "Código de Cliente : "
         Height          =   255
         Left            =   -72330
         TabIndex        =   7
         Top             =   4440
         Width           =   2175
      End
      Begin VB.Label Label13 
         Caption         =   "Código de Repato: "
         Height          =   255
         Left            =   -74790
         TabIndex        =   3
         Top             =   4410
         Width           =   2175
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox18 
      Height          =   225
      Left            =   90
      TabIndex        =   165
      Top             =   360
      Width           =   17655
      _Version        =   851968
      _ExtentX        =   31141
      _ExtentY        =   397
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
   End
   Begin VB.Menu acciones 
      Caption         =   "Acciones"
      Begin VB.Menu fpc 
         Caption         =   "Facturarle a un CLIENTE las facturas seleccionadas de los PROVEEDORES"
      End
   End
   Begin VB.Menu auto 
      Caption         =   "Automatizaciones"
      Begin VB.Menu factmensual 
         Caption         =   "Activar. Facturar mensualmente"
      End
      Begin VB.Menu dfacmensual 
         Caption         =   "Desactivar. Facturación mensual"
      End
   End
End
Attribute VB_Name = "frmBuscarFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vsql, vwhere, vgrupo, vorden, vsql2 As String
Dim vdocseleccionados As String
Public viene, vVieneRemito, cpFactura, CP As String
Public vieneCobro As Boolean
Dim vid As String
Dim vsqlpago(100) As String

Dim recibosApagar(50) As String

Dim vsqlpagoAuto(100) As String
Dim vshift As Long
Dim vvtotal, vvpagado, vvsaldo As Double
Dim vvpagoPacial As Boolean
Dim vgrow As Long
Public vdocmarcados As Long
Public vTipoComprobante As Long
Public venro, veLetra As String
Public veptovta, veDocumento As String
Dim vf As vfactura
Public vnroFA, vnroFB As Long
Dim vidVendedor As Long
Dim vecondicion As String
Dim cadena_errores As String

Dim vnroCodigoBarra As String

Dim vvmes, vvano  As Long

Dim arr_as(500) As String

Dim ii As Long


Dim vpap2, vvalorespap2, vvalorespap2_temp As String
Dim varreglo_pago2(400) As String

Dim varreglo_pago2_temp(400) As String


Public Sub BuscarCliente(codCli As String)
    On Error Resume Next
    
    Dim rsClientes As New ADODB.Recordset, sqlClientes As String

    sqlClientes = "SELECT * FROM " + CP + " WHERE (codigo = '" & codCli & "') or (nombre LIKE '%" & codCli & "%')"

    With rsClientes
        Call .Open(sqlClientes, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If .EOF Then
            frmBuscarCliente.Show
            frmBuscarCliente.o = 3
            frmBuscarCliente.txtClientes = txtCliente.Text
            frmBuscarCliente.txtClientes.SetFocus
        End If
    
        If Not .EOF = True Then
        
            If TabDocumentos.tab = 0 Then
            
                txtCliente.Text = EsNulo(.Fields("Nombre").Value)
                txtCliente.Tag = EsNulo(.Fields("Codigo").Value)
            
            Else

                txtClienteImpagos.Text = EsNulo(.Fields("Nombre").Value)
                txtClienteImpagos.Tag = EsNulo(.Fields("Codigo").Value)
            
            End If
        
        End If
    
    End With
    
    sqlClientes = ""
    
    If rsClientes.State = 1 Then
        rsClientes.Close
        Set rsClientes = Nothing
    End If
    
    If Err Then GrabarLog "BuscarCliente" & codCli, Err.Number & " " & Err.Description, Me.Name
End Sub
Function BuscarRepartidor(vCodigoRepartidor As String) As String
On Error Resume Next

    If Trim(vCodigoRepartidor) = "" Then Exit Function
    
    Dim rsRepartidor As New ADODB.Recordset
    Dim sqlRepartidor As String

    sqlRepartidor = "SELECT codigo, nombre FROM empleados WHERE (codigo = '" & vCodigoRepartidor & "')"
    
    With rsRepartidor
        Call .Open(sqlRepartidor, ConnDDBB, adOpenStatic, adLockReadOnly)
    
        If Not .EOF = True Then
            BuscarRepartidor = .Fields("Nombre").Value
        Else
            BuscarRepartidor = ""
        End If
    
    End With
    
    sqlRepartidor = ""
    
    rsRepartidor.Close
    Set rsRepartidor = Nothing
    
    If Err Then GrabarLog "BuscarRepartidor", Err.Number & " " & Err.Description, Me.Name
End Function

Public Sub calTotales()
    On Error Resume Next

    Dim vtfacta, vtmono, vtnd, vtnc As Double
    Dim vifacta, vimono, vind, vinc As Double

    bfactura.Refresh

    vtfacta = 0
    vtmono = 0
    vtnd = 0
    vtnc = 0

    vifacta = 0
    vimono = 0
    vind = 0
    vinc = 0

    Do Until bfactura.Recordset.EOF
    
        If bfactura.Recordset("tipo") = "Fact A" Then
            vtfacta = vtfacta + bfactura.Recordset("total_ctacte") + bfactura.Recordset("total_cdo")
            vifacta = vifacta + bfactura.Recordset("tiva")
        End If
    
        If bfactura.Recordset("tipo") = "Fact B" Then
            vtmono = vtmono + bfactura.Recordset("total_ctacte") + bfactura.Recordset("total_cdo")
            vimono = vimono + bfactura.Recordset("tiva")
        End If
    
        If bfactura.Recordset("tipo") = "Nota C" Then
            vtnc = vtnc + bfactura.Recordset("total_ctacte") + bfactura.Recordset("total_cdo")
            vinc = vinc + bfactura.Recordset("tiva")
        End If
    
        If bfactura.Recordset("tipo") = "Nota D" Then
            vtnd = vtnd + bfactura.Recordset("total_ctacte") + bfactura.Recordset("total_cdo")
            vind = vind + bfactura.Recordset("tiva")
        End If
    
        bfactura.Recordset.MoveNext
    Loop

    tfacta.Caption = Format(vtfacta, "########0.000")
    tmono.Caption = Format(vtmono, "########0.000")
    tnc.Caption = Format(vtnc, "########0.000")
    tnd.Caption = Format(vtnd, "########0.000")

    ifactaLabel.Caption = Format(vifacta, "########0.000")
    imono.Caption = Format(vimono, "########0.000")
    inc.Caption = Format(vinc, "########0.000")
    ind.Caption = Format(vind, "########0.000")

    If Err Then GrabarLog "caltotales", Left(vsql, 99), Me.Name
End Sub
Public Sub cmdAnularF_Click()
    On Error Resume Next

    If MsgBox("Esta seguro que desea anular la factura", vbExclamation + vbYesNo, "Mensaje ...") = vbYes Then

        With bfactura

            If .Recordset.RecordCount = 0 Then Exit Sub
            If Right(.Recordset("tipo"), 3) = "(A)" Then
                MsgBox "El doc. seleccionado ya ha sido anulado.", vbInformation, "Mensaje ..."
                Exit Sub
            End If

            .Recordset("tipo") = bfactura.Recordset("tipo") + "(A)"
            .Recordset.Update
        End With

        MsgBox "Factura Anulada", vbInformation, "Mensaje ..."
    End If

    If Err Then GrabarLog "cmdAnularF_Click", Err.Number & " " & Err.Description, Me.Name
End Sub


Private Sub pasarAIE()

frmIngresosEgresos.vpagoPacial = vvpagoPacial

frmIngresosEgresos.vobservacion.Text = vdocseleccionados

frmIngresosEgresos.vtotalcontrol.Text = Me.vtotal.Caption
frmIngresosEgresos.setvsqlPago vsqlpago()
frmIngresosEgresos.setvsqlPagoAuto vsqlpagoAuto()

frmIngresosEgresos.vtotalseleccionado = Val(Me.vtotal.Caption)

frmIngresosEgresos.actualizarTotales


'frmCobros.vdocSeleccionado.Text = vdocseleccionados  ' comentario
'frmCobros.txtMontoTotalPendienteSeleccionado = Me.vtotal.Caption
'frmCobros.setvsqlPago vsqlpago()
'frmCobros.setvsqlPagoAuto vsqlpagoAuto()

End Sub


Private Sub pasarACobros()

frmCobros.vdocSeleccionado.Text = vdocseleccionados  ' comentario
frmCobros.txtMontoTotalPendienteSeleccionado = Me.vtotal.Caption
frmCobros.setvsqlPago vsqlpago()
frmCobros.setvsqlPagoAuto vsqlpagoAuto()
frmCobros.setvsqlPago2 varreglo_pago2()
frmCobros.setvsqlPago2_temp varreglo_pago2_temp()


End Sub

Private Sub bfactura_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
MsgBox Description
End Sub

Public Sub cmdBorrar_Click()
    On Error Resume Next
    
    If Not Me.vreferenciaPedido = "" Then
    
        If MsgBox("Quiere borrar todos los pedidos exportados con referencia: " + Trim(Me.vreferenciaPedido), vbYesNo) = vbYes Then
        
            Call borrarPedidosExportados(Trim(Me.vreferenciaPedido))
            Exit Sub
        
        End If
    
    End If
    
    
    frmTransaccionMantenimiento.vnrointerno = Val(Me.KlexDocumentos.TextMatrix(Me.KlexDocumentos.Row, 11))
    frmTransaccionMantenimiento.Show
    
    Unload Me
    Exit Sub
    
    '---------------- panic !  Acomodar esto
    
    Dim vRemitoABorrar As Long, vnroasiento As Long

    If MsgBox(" ¿ Esta seguro que desea dar de baja este movimiento ?  ", vbInformation + vbYesNo, "Mensaje ...") = vbNo Then Exit Sub
    
    With Me.KlexDocumentos
        If Not (.Rows = 1) And Not Val(.TextMatrix(.Row, 13)) = 0 Then
            vRemitoABorrar = Val(.TextMatrix(.Row, 13))
            vnroasiento = Val(.TextMatrix(.Row, 14))
            
        Else
            Exit Sub
        End If
    End With
    
    '---------------- borro movimiento en cuentas corrientes ------------------------
    Call BorrarBase("Factura WHERE (remito = " & vRemitoABorrar & ")", pathDBMySQL)
    
    '---------------- borro movimiento en cuentas corrientes ------------------------
    Call BorrarBase("CuentasCorrientes WHERE (remito = " & vRemitoABorrar & ")", pathDBMySQL)
    
    '---------------- borro movimiento en caja diaria -------------------------
    Call BorrarBase("Caja WHERE (remito = " & Val(vRemitoABorrar) & ")", pathDBMySQL)

    '---------------- Borro el Asiento ---------------------------------------
    Call BorrarBase(" Asientos WHERE Numero = " & Val(vnroasiento) & "", pathDBMySQL)

    '---------------- Borro el Asiento ---------------------------------------
    Call BorrarBase(" AsientosDetalle WHERE Numero = " & Val(vnroasiento) & "", pathDBMySQL)

    '---------------- borro el libro de iva -------------------------
    'Panic: Controlar que el IVA NO ESTE CERRADO (IvaVenta)
    Call BorrarBase("IvaFacturaVenta WHERE (remito = " & Val(vRemitoABorrar) & ")", pathDBMySQL)
    
    '--------------- borro los detalles -------------------------
    Dim rsFDetalle As New ADODB.Recordset, sqlFDetalle As String
    
    sqlFDetalle = "SELECT * FROM fdetalle WHERE (remito = " & Val(vRemitoABorrar) & ")"
    
    With rsFDetalle
        Call .Open(sqlFDetalle, ConnDDBB, adOpenStatic, adLockReadOnly)

        Do Until .EOF = True
            Call ModificarStock(1, .Fields("cantidad").Value, .Fields("codigo").Value)
            .MoveNext
        Loop
    
    End With
    
    Call BorrarBase("Fdetalle WHERE (remito = " & Val(vRemitoABorrar) & ")", pathDBMySQL)
    
    cmdFiltrar_Click
    
    If Err Then GrabarLog "cmdBorrar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Public Sub borrarPedidosExportados(vref As String)
Dim vsql As String
Dim vnrointerno As Long
Dim i As Long

barra.Max = Me.KlexDocumentos.Rows - 1
barra.Value = 1
For i = 1 To barra.Max

    barra.Value = barra.Value + 1
    
    vsql = "select nrointerno from factura where idfactura= " + Me.KlexDocumentos.TextMatrix(i, 1) + " and refexportpedidos = '" + vref + "'"

    vnrointerno = Val(traerDatos2(vsql, "nrointerno", pathDBMySQL))
    
    If Not vnrointerno = 0 Then borrarTodosLosModulos (vnrointerno)


Next


'vsql = "delete factura,fdetalle, cuentascorrientes from factura " & _
" inner join fdetalle  on  factura.remito = fdetalle.remito" & _
" inner join cuentascorrientes  on  factura.NroInterno = cuentascorrientes.NroInterno" & _
" inner join ivafacturaventa   on factura.NroInterno  = ivafacturaventa.nrointerno" & _
" Where Factura.refexportpedidos = '" + vref + "'"


'Call EjecutarScript(vsql, pathDBMySQL)

End Sub



Function validarFiltro() As Boolean
Dim vmensaje As String
validarFiltro = True

vmensaje = ""

If Me.chkFacturaA.Value = Me.chkMonotributo.Value And Not Me.vreferenciaPedido = "" And False Then
    vmensaje = vmensaje + Chr(13) + "> Debe seleccionar solamente un solo tipo de factura para los pedidos facturado con referencia : " + Me.vreferenciaPedido
End If

If Not vmensaje = "" Then
    MsgBox vmensaje
    validarFiltro = False
End If

End Function

Private Sub set_mostrar_filtros()

vecondicion = ""

If Not Me.vdescEmpresa = "" Then vecondicion = vecondicion + " > Empresa: " + Me.vdescEmpresa + Chr(13)
If Not Me.txtCliente = "" Then vecondicion = vecondicion + " > Cliente: " + txtCliente + Chr(13)
If Not Me.vDesRepartidor = "" Then vecondicion = vecondicion + " > Vendedor: " + vDesRepartidor + Chr(13)
If Not Me.txtEmpleado = "" Then vecondicion = vecondicion + " > Empleado: " + txtEmpleado + Chr(13)
If Not Me.chkFechaTodas.Value = 1 Then vecondicion = vecondicion + " > Desde : " + Str(Me.dtpFecha(0)) + " - Hasta: " + Str(Me.dtpFecha(1))


End Sub


Public Sub cmdFiltrar_Click()
On Error Resume Next

Dim vsqlOrdenado, vsqlAgrupado, vvcampo As String

    
    If Not validarFiltro Then Exit Sub
    
    vsqlOrdenado = ""
    vsqlAgrupado = ""
    vvcampo = "*"
    

    vsql = ""
    
    
   'If Not Me.vrubro.Tag = "" Then
    
   ' If Not Me.vsubrubro.Tag = "" Then
    

    set_mostrar_filtros
    
    
    If Not Trim(txtCliente.Tag) = "" Then vsql = vsql + " AND (codigo = '" & Trim(txtCliente.Tag) & "')"
    'If Not Trim(me.te.Text) = "" Then vSQL = vSQL + " AND (Cod_Repartidor = '" & Trim(txtEmpleado.Tag) & "')"

    If Not (txtNDesde.Text = "") And Not (txtNHasta.Text = "") Then
        vsql = vsql + " AND (ncomprobante >= " & txtNDesde.Text & " AND NComprobante <= " & txtNHasta.Text & ")"
    End If


    If Not Me.cmbEstadoDocumento.Text = "Todos los estados" Then vsql = vsql + " and (estadodocumento = '" + Trim(Me.cmbEstadoDocumento.Text) + "')"


    vsql = vsql + " and (1<1 "
    
    If chkFacturaA.Value = 1 Then vsql = vsql + " OR (tipo = 'Fact A')"
    If chkMonotributo.Value = 1 Then vsql = vsql + " OR (tipo = 'Fact B')"
    If chkFacturaC.Value = 1 Then vsql = vsql + " OR (tipo = 'Fact C')"
    If chkNotasDe.Value = 1 Then vsql = vsql + " OR (Tipo = 'Nota D')"
    If chkRemito.Value = 1 Then vsql = vsql + " OR (Tipo = 'Remito')"
    If chkDocNo.Value = 1 Then vsql = vsql + " OR (Tipo = 'Documento')"
    If chkPresupuesto.Value = 1 Then vsql = vsql + " OR (Tipo = 'Presupuesto')"
    If chkNotaCA.Value = 1 Then vsql = vsql + " OR (tipo = 'Nota C' AND (Letra = 'A'))"
    If chkNotaCB.Value = 1 Then vsql = vsql + " OR (tipo = 'Nota C' AND (Letra = 'B'))"
    If chkNotaCC.Value = 1 Then vsql = vsql + " OR (tipo = 'Nota C' AND (Letra = 'C'))"
    
    If chkOtros.Value = 1 Then vsql = vsql + " OR (Tipo = 'Otros')"
    
    If chkFacturaX.Value = 1 Then vsql = vsql + " OR (tipo = 'Fact X')"
    
    vsql = vsql + ")"
    

    If chkFechaTodas.Value = 0 Then
        'If Trim(dtpFecha(0).Text) = True And fhasta.Enabled = True Then
            vsql = vsql + " and Fecha >= '" & strfechaMySQL(dtpFecha(0).Value) + "' and fecha <= '" & strfechaMySQL(dtpFecha(1).Value) + "'"
        
        'End If
    End If
 
 
 If Not LeerXml("Puesto") = "Empresas" Then
 
     If Not vcodRepartidor.Text = "" Then vsql = vsql + " and (Codigo  in " + Trim(Procedimientos.getRepartidor2idFactura(Str(vidVendedor))) + ") "
 
 Else
 
     If Not vcodRepartidor.Text = "" Then vsql = vsql + " and (idfactura in " + Trim(Procedimientos.getRepartidor2idFactura(Me.vcodRepartidor.Text)) + ") "
 
 End If
    

    If Not vcodEmpresa.Text = "" Then vsql = vsql + " and (idfactura in " + Trim(Procedimientos.getEmpresa2idFactura(Me.vcodEmpresa.Text)) + ") "
  

    If Not Me.chkFechaPago.Value = 1 Then
            vsql = vsql + " and fechapago >= '" & strfechaMySQL(dtpFecha(4).Value) + "' and fechapago <= '" & strfechaMySQL(dtpFecha(5).Value) + "'"
    End If
 
 
 
    If Not Me.vreferenciaPedido.Text = "" Then
        
        vsql = vsql + " and factura.refexportpedidos = '" + Trim(Me.vreferenciaPedido) + "'"
    
    End If
 
 
  If Me.rbAdeudado Then
        'If Trim(dtpFecha(0).Text) = True And fhasta.Enabled = True Then
            vsql = vsql + " and (total-pagoparcial) > 0.1 "
        'End If
    End If
 
 
  If chkFechaVencimiento.Value = 0 Then
        'If Trim(dtpFecha(0).Text) = True And fhasta.Enabled = True Then
            vsql = vsql + " and FechaVencimiento >= '" & strfechaMySQL(dtpFecha(2).Value) + "' and FechaVencimiento <= '" & strfechaMySQL(dtpFecha(3).Value) + "'"
        'End If
    End If
 
vsql2 = vsql
 
    If RadPersona.Value Then
        vsqlAgrupado = " group by codigo"
        vvcampo = vid + ",Tipo,TipoMovimiento, Letra, PuntoDeVenta, NComprobante, NroRemito, fecha, Codigo, Nombre, CUIT, sum(total) as Total, NroInterno, comentario, remito,NroAsiento,cantidadvolquetes,idlistachoferes,estadodocumento,pagoparcial,cantidadvolquetedevuelto,pagoparcial"
           
    End If
    
    If Me.vordenadoPor = "Importe" Then
        vsqlOrdenado = " order by Total"
   End If
    
    If vordenadoPor = "Fecha" Then
        vsqlOrdenado = " order by Fecha"
    End If
    
    If vordenadoPor = "Codigo" Then
        vsqlOrdenado = " order by codigo"
    End If
    
    
    
    If Not Me.vreferenciaPedido = "" Then vsqlOrdenado = " order by ncomprobante asc"
     
   
    With bfactura
        .ConnectionString = pathDBMySQL
        
        If Me.RadProvincias Then
            .RecordSource = "SELECT  " + CP + ".Provincia, sum(" + cpFactura + ".Total) AS TotalFacturado From  ivafacturaventa t INNER JOIN " + cpFactura + " ff ON (t.Remito=ff.Remito) INNER JOIN " + CP + " cc ON (cc.Codigo=ff.Codigo) where 1=1 " + vsql + " Group By  cc.Provincia"
            .Refresh
            Set KlexDocumentos.Recordset = .Recordset
            
        Else
            .RecordSource = "SELECT " + vvcampo + ", (total - case when pagoparcial is null then 0 " + _
            " else pagoparcial end) as saldos1   FROM " + cpFactura + " WHERE (1=1" & vsql & ")" + vsqlAgrupado + vsqlOrdenado
            .Refresh
            
            CargarDocumentos
            
           ' Set kk.Recordset = .Recordset
        End If

    
    End With
    
    Me.tab.Item(3).Selected = True  ' cambia a la solapa de datos
    
 '   fpintarGrilla
 
    Call formatGrillaDoc(10, "##########0.00")
    Call formatGrillaDoc(20, "##########0.00")
    
    cmdImprimir(0).Tag = ""
    
     log2.Visible = False
     

    If Err Then GrabarLog "cmdFiltrar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub seleccionartodo(o As Boolean)
Dim i As Long

With Me.KlexDocumentos
Me.barra.Max = .Rows - 1
Me.barra.Value = 0


        For i = 1 To .Rows - 1
            Me.barra.Value = Me.barra.Value + 1
             
            If o Then
                vdocmarcados = vdocmarcados + 1
        
                .TextMatrix(i, 0) = "X"
            Else
                .TextMatrix(i, 0) = " "
                vdocmarcados = vdocmarcados - 1
        
            End If
        Next
        
End With

End Sub

Function validariImpreFE() As Boolean
Dim i, j As Long
Dim vmensaje As String

validariImpreFE = True
vmensaje = ""

With Me.KlexDocumentos
        For i = 1 To .Rows - 1
        
            If Trim(.TextMatrix(i, 9)) = "" And Not Trim(.TextMatrix(i, 0)) = "" And Not Trim(.TextMatrix(i, 21)) = "Documento" Then
                vmensaje = "Hay facturas documentos sin cuit" + Chr(13) + "Está marcado en rojo."
                .TextMatrix(i, 9) = "???????"
                .Row = i
                .Col = 9
                .CellBackColor = vbRed
            End If

        Next
End With

If Not vmensaje = "" Then
    MsgBox vmensaje
    validariImpreFE = False
End If

End Function


Private Sub fpintarGrilla()
Dim i, j As Long
Dim v1, v2, vtipo  As String



For i = 1 To Me.KlexDocumentos.Rows - 1

v1 = Me.KlexDocumentos.TextMatrix(i, 19)
v2 = Me.KlexDocumentos.TextMatrix(i, 15)
vtipo = Me.KlexDocumentos.TextMatrix(i, 21)


For j = 1 To KlexDocumentos.Cols - 1
Me.KlexDocumentos.Row = i
Me.KlexDocumentos.Col = j

If v1 = "Retirado" And v2 = "Pagado" Then Me.KlexDocumentos.CellBackColor = vbBlue
If v1 = "Retirado" And v2 = "Adeudado" Then Me.KlexDocumentos.CellBackColor = vbRed

If v1 = "Recambio" And v2 = "Pagado" Then Me.KlexDocumentos.CellBackColor = vbGreen
If v1 = "Recambio" And v2 = "Adeudado" Then Me.KlexDocumentos.CellBackColor = vbYellow - 5

If vtipo = "Documento" Then Me.KlexDocumentos.CellForeColor = vbGreen
If vtipo = "Nota C" Then Me.KlexDocumentos.CellForeColor = vbRed
If vtipo = "Fact A" Then Me.KlexDocumentos.CellForeColor = vbBlack
If vtipo = "Fact B" Then Me.KlexDocumentos.CellForeColor = vbBlue
If vtipo = "Fact D" Then Me.KlexDocumentos.CellForeColor = vbRed

Next

Next

End Sub


Private Sub pintarlinea(l As Long)
Dim i, j As Long

Me.KlexDocumentos.Row = l

For j = 1 To KlexDocumentos.Cols - 1
    
    Me.KlexDocumentos.Col = j
    Me.KlexDocumentos.CellBackColor = vbYellow + 30

Next

End Sub

Private Sub CargarDocumentos() ' ema: pone el filtro de la tabla factura en el data grid
    On Error Resume Next
    
    Dim i As Long
    
    
    KlexDocumentos.Visible = False
    
    With bfactura
    
        If Not .Recordset.EOF = True Then
            .Recordset.MoveFirst
        Else
            MsgBox "No hay datos para mostrar", vbInformation, "Consulta..."
            Exit Sub
        End If
        
        
        FormatoGrilla (.Recordset.RecordCount) ' ema: formatea las columnas de la grilla
        
        i = 1
        
        KlexDocumentos.Visible = True
        
        
        vvtotal = 0
        vvpagado = 0
        vvsaldo = 0
        
        
        If Not .Recordset.RecordCount > 0 Then Exit Sub
        Do Until .Recordset.EOF = True
        
            KlexDocumentos.TextMatrix(i, 0) = ""
            
            
            KlexDocumentos.TextMatrix(i, 1) = EsNulo(.Recordset(vid).Value)
            KlexDocumentos.TextMatrix(i, 2) = EsNulo(.Recordset("Tipo").Value)
            KlexDocumentos.TextMatrix(i, 3) = EsNulo(.Recordset("Letra").Value)
            KlexDocumentos.TextMatrix(i, 4) = EsNulo(.Recordset("PuntoDeVenta").Value)
            KlexDocumentos.TextMatrix(i, 5) = EsNulo(.Recordset("NComprobante").Value)
            KlexDocumentos.TextMatrix(i, 6) = EsNulo(.Recordset("Fecha").Value)
            KlexDocumentos.TextMatrix(i, 7) = EsNulo(.Recordset("Codigo").Value)
            KlexDocumentos.TextMatrix(i, 8) = EsNulo(.Recordset("Nombre").Value)
            
            KlexDocumentos.TextMatrix(i, 9) = EsNulo(.Recordset("Cuit").Value)
            
        
            KlexDocumentos.TextMatrix(i, 10) = Format(.Recordset("Total").Value, "######0.00")
            KlexDocumentos.TextMatrix(i, 11) = EsNulo(.Recordset("NroInterno").Value)
            KlexDocumentos.TextMatrix(i, 12) = EsNulo(.Recordset("Comentario").Value)
            KlexDocumentos.TextMatrix(i, 13) = EsNulo(.Recordset("Remito").Value)
            KlexDocumentos.TextMatrix(i, 14) = EsNulo(.Recordset("NroAsiento").Value)
            KlexDocumentos.TextMatrix(i, 24) = EsNulo(.Recordset("mensual").Value)
            
            
            ' ----------- agregado para volquete ------
            KlexDocumentos.TextMatrix(i, 15) = EsNulo(.Recordset("estadodocumento").Value)
            
            
           ' KlexDocumentos.TextMatrix(i, 16) = "*" + (.Recordset("idlistachoferes").Value)
           ' KlexDocumentos.TextMatrix(i, 17) = EsNulo(.Recordset("cantidadvolquetes").Value)
            
            'KlexDocumentos.TextMatrix(i, 18) = EsNulo(.Recordset("cantidadvolquetedevuelto").Value)
             
            KlexDocumentos.TextMatrix(i, 19) = EsNulo(.Recordset("tipopedido").Value)
            KlexDocumentos.TextMatrix(i, 20) = EsNulo(.Recordset("pagoparcial").Value)
            '--------------------------------------------
            
             KlexDocumentos.TextMatrix(i, 21) = EsNulo(.Recordset("Tipo").Value)
            
            KlexDocumentos.TextMatrix(i, 22) = Val(KlexDocumentos.TextMatrix(i, 10)) - Val(KlexDocumentos.TextMatrix(i, 20))
            
            KlexDocumentos.TextMatrix(i, 25) = getvdesEmpresa(.Recordset(vid).Value)
            
            KlexDocumentos.TextMatrix(i, 26) = getvdesRepartidor(.Recordset(vid).Value)
            
            
            
            vvtotal = vvtotal + Val(KlexDocumentos.TextMatrix(i, 10))
            vvpagado = vvpagado + Val(KlexDocumentos.TextMatrix(i, 20))
            vvsaldo = vvtotal + vvpagado
            
             
            KlexDocumentos.TextMatrix(i, 23) = EsNulo(.Recordset("fechapago").Value)

            
            KlexDocumentos.Row = i
            KlexDocumentos.Col = 10
             
             If Not Val(KlexDocumentos.TextMatrix(i, 22)) < 0.1 Then KlexDocumentos.CellForeColor = vbRed
            
            'klexDocumentos.TextMatrix(i, 13) = EsNulo(.Recordset("Endoso").Value)
            
            'klexDocumentos.TextMatrix(i, 15) = EsNulo(.Recordset("FechaAcreditacion").Value)
            
            'klexDocumentos.TextMatrix(i, 17) = EsNulo(.Recordset("NroInterno").Value)
            'klexDocumentos.TextMatrix(i, 18) = EsNulo(.Recordset("Observaciones").Value)
            
             If .Recordset("estado2") = "Impreso" Then pintarlinea (i)
                
                
            .Recordset.MoveNext
        
        
            i = i + 1
        Loop
        
        
        
        KlexDocumentos.TopRow = i - 1
        .Refresh
        
        
        vltotal = Format(vvtotal, "###,###,##0.00")
        vlpagado = Format(vvpagado, "###,###,##0.00")
        vlsaldo = Format(vvtotal - vvpagado, "###,###,##0.00")
        
        
        'If .Recordset.EOF = True Then .Recordset.MoveLast
    End With
    
If Err Then GrabarLog "CargarDocumentos", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub CmdEjecutarCobro_Click()
On Error Resume Next

    If (bfactura.Recordset.EOF > True) Then
        With frmCobros
            .NroComprobante = EsNulo(bfactura.Recordset("NComprobante").Value)
            .tipoComprobante = EsNulo(bfactura.Recordset("Tipo").Value)
            .fechaDocumento = EsNulo(bfactura.Recordset("Fecha").Value)
            .remito = EsNulo(bfactura.Recordset("remito").Value)
            .codCliente = EsNulo(bfactura.Recordset("codigo").Value)
           
            'si el cliente no coincide se debe alertar y no continuar el proceso
            If (.codCliente <> EsNulo(bfactura.Recordset("codigo").Value)) And (.codCliente <> "") Then
                MsgBox "El documento seleccionado corresponde a un cliente distindo del que se está cobrando, revise su operación."
            Else
        
                .BuscarDatosOperacionesCliente EsNulo(bfactura.Recordset("codigo").Value), EsNulo(bfactura.Recordset("remito").Value)
                .esComprobanteAutomatico = False
        
                .txtNroComprobante.Text = EsNulo(bfactura.Recordset("NComprobante").Value)
                .txtTipoComp.Text = EsNulo(bfactura.Recordset("Tipo").Value)
                            
                .Show
            End If
            
        End With
        
        Unload Me
    Else
        MsgBox "Debe seleccionar un documento de venta", vbInformation, "WGestion"
    End If
    
If Err Then GrabarLog "CmdEjecutarCobro_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub cmdImprimeD_Click()
On Error Resume Next

If Err Then GrabarLog "cmdImprimeD_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Public Sub cmdImprimeF_Click()
On Error Resume Next
    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox "    Prepare la impresora    ", vbInformation, "Mensaje ..."
    
    With Mantenimiento.rsfacturas_impagas_factA
        If Not Trim(txtRepartoImpagos.Text) = "" Then
            .Filter = "reparto = '" & Trim(txtRepartoImpagos.Tag) & "'"
        Else
            .Filter = "Codigo = '" & Trim(txtClienteImpagos.Tag) & "'"
        End If
        .Sort = "Nombre ASC"

    End With
    'With drfacturas_impagas_factA
    '    .Sections("TituloEmpresa").Controls("vreparto").Caption = "Reparto nro: " & Trim(txtRepartoImpagos.Text) & ""
    '    .Show
    'End With
If Err Then GrabarLog "cmdImprimeF_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Public Sub cmdImprimeFD_Click()
On Error Resume Next
    
    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox "    Prepare la impresora    ", vbInformation, "Mensaje ..."
    
    With Mantenimiento.rsfacturas_impagas
        If Not Trim(txtRepartoImpagos.Text) = "" Then
            .Filter = "reparto = '" & txtRepartoImpagos.Tag & "'"
        Else
            .Filter = "Codigo = '" & Trim(txtClienteImpagos.Tag) & "'"
            .Sort = "Nombre ASC"
        End If

    End With
    
    'With drfacturas_impagas
    '    .Sections("TituloEmpresa").Controls("vreparto").Caption = "Reparto nro: " + Trim(txtRepartoImpagos.Text)
    '    .Show
    'End With

If Err Then GrabarLog "cmdImprimeFD_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Public Sub imprimirEstadisticaMensual()
On Error Resume Next
        MsgBox "    Prepare la impresora    ", vbInformation, "Mensaje ..."
        
        Unload Mantenimiento
        Load Mantenimiento
        
        
        
        With Mantenimiento.rsEstadisticasMes
            If .State = 1 Then .Close
            
            .Source = bfactura.RecordSource    ' Emma: le pasa lo que está en bfactura al conector del datareport
            
            If .State = 0 Then .Open
            .Close
            .Open
        End With
        
        
        
        
        With drEstadisticasMensuales
           ' .Show
           ' Exit Sub
        
            '.DataMember = bfactura.RecordSource
            
            
            
'            .Sections("detalle").Controls("mes").DataField = "mes"
'            .Sections("detalle").Controls("mes").DataMember = bfactura.RecordSource
'
'            .Sections("detalle").Controls("ano").DataField = "ano"
'            .Sections("detalle").Controls("ano").DataMember = bfactura.RecordSource
'
'            .Sections("detalle").Controls("facturado").DataField = "facturado"
'            .Sections("detalle").Controls("facturado").DataMember = bfactura.RecordSource
'
'            .Sections("detalle").Controls("cobrado").DataField = "cobrado"
'            .Sections("detalle").Controls("cobrado").DataMember = bfactura.RecordSource
'
'            .Sections("detalle").Controls("porcentaje").DataField = "porcentaje"
'            .Sections("detalle").Controls("porcentaje").DataMember = bfactura.RecordSource
'
'            .Sections("detalle").Controls("sfacturado").DataField = "sfacturado"
'            .Sections("detalle").Controls("sfacturado").DataMember = bfactura.RecordSource
'
'            .Sections("detalle").Controls("scobrado").DataField = "scobrado"
'            .Sections("detalle").Controls("scobrado").DataMember = bfactura.RecordSource
'
'            .Sections("detalle").Controls("sporcentaje").DataField = "sporcentaje"
'            .Sections("detalle").Controls("sporcentaje").DataMember = bfactura.RecordSource
            
            
          '  .Refresh
            
            .Show
        
        End With
        
        MousePointer = vbDefault

If Err Then Exit Sub
End Sub


Public Sub cmdImprimir_Click(Index As Integer)
On Error Resume Next
    
' imprime grupos o factura

If Me.rbMeses Then
    imprimirEstadisticaMensual
    Exit Sub
End If

If Me.rbAnos Then
    imprimirEstadisticaMensual
    Exit Sub
End If

If Me.rdArticulo.Value Then
   vsql = sqlArtFactVenta(vwhere, vgrupo, vorden)

   With bfactura
        .ConnectionString = pathDBMySQL
        .RecordSource = vsql
        .Refresh
         Set KlexDocumentos.Recordset = .Recordset
    End With

End If

If cmdImprimir(0).Tag = "grupos" Then
    imprimirEstadistica
    Exit Sub
End If
    
    
       If Index = 0 Then
        MousePointer = vbHourglass
    
        MsgBox "Prepare la impresora", vbInformation, "Mensaje ..."
        
        Unload Mantenimiento
        Load Mantenimiento
        
        With Mantenimiento.rsldoc
            If .State = 1 Then .Close
            
            .Source = bfactura.RecordSource    ' Emma: le pasa lo que está en bfactura al conector del datareport
            
            If .State = 0 Then .Open
            .Close
            .Open
        End With
        
        With drListadoFacturaventa
            .Sections("TituloEmpresa").Controls("econdicion").Caption = vecondicion
            .Show
        End With
        
        MousePointer = vbDefault
    Else
        Dim vImportePagoPesos As Double

        If (vImpresoras.vNombreImpresora = "Hasar") Then
            With KlexDocumentos
                If Not (KlexDocumentos.Rows = 1) Then
                    If MsgBox("Desea Imprimir la Factura " & EsNulo(.TextMatrix(.Row, 5)) & " ?", vbInformation + vbYesNo, "Mensaje ...") = vbYes Then
                        vImportePagoPesos = TraerDato("CuentasCorrientes", "Remito = " & Val(.TextMatrix(.Row, 13)) & " AND Debito = 0", "Credito")
                        Call ImprimirHasar(Val(.TextMatrix(.Row, 13)), Val(vImportePagoPesos))
                        vImportePagoPesos = 0
                    End If
                End If
            End With
        End If
    End If
    
    If Err Then GrabarLog "cmdImprimir_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub imprimirEstadistica()
On Error Resume Next
    
        MousePointer = vbHourglass
   
        MsgBox "Prepare la impresora", vbInformation, "Mensaje ..."
        
        Unload Mantenimiento
        Load Mantenimiento
        
        With Mantenimiento.rsEstadisticas
            If .State = 1 Then .Close
            
            .Source = bfactura.RecordSource    ' Emma: le pasa lo que está en bfactura al conector del datareport
            
            If .State = 0 Then .Open
            .Close
            .Open
        End With
        
        With drEstadisticaVenta
            If Me.rdPersona.Value Then
               ' .Sections("detalle").Controls("texto6").DataField = "nombre"
            End If
            .Show
        End With
         MousePointer = vbDefault
    
If Err Then GrabarLog "cmdImprimir_Click", Err.Number & " " & Err.Description, Me.Name
End Sub


Private Sub ImprimirHasar(vRemitoAImprimir As Long, vMontoEnEF As Double)
On Error Resume Next

    Dim FS As String
    
    FS = Chr$(28) '// Separador de campos del comando

    Dim rsImprimirHasar As New ADODB.Recordset, sqlImprimirHasar As String
    
    MsgBox "Prepare la Impresora ", vbInformation, "Mensaje ..."
    
    sqlImprimirHasar = "SELECT * FROM ImpresionFactura WHERE (Remito = " & Val(vRemitoAImprimir) & ")"

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
                Call .DatosCliente(EsNulo(rsImprimirHasar.Fields("Nombre").Value), EsNuloGuion(rsImprimirHasar.Fields("NroDocumento").Value), TIPO_DNI, CONSUMIDOR_FINAL, Left(EsNulo(rsImprimirHasar.Fields("CodigoPostal").Value) & "-" & EsNuloGuion(rsImprimirHasar.Fields("Localidad").Value) & "     -     " & EsNuloGuion(rsImprimirHasar.Fields("Direccion").Value), 50))
        
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

            Case "Remito"
                '
                
        End Select
        
        Dim rsDetalleHasar As New ADODB.Recordset, sqlDetalleHasar As String
        
        sqlDetalleHasar = "SELECT * FROM FDetalle WHERE (Remito = " & Val(vRemitoAImprimir) & ") ORDER BY idFDetalle ASC"
        
        If rsDetalleHasar.State = 1 Then rsDetalleHasar.Close
        rsDetalleHasar.CursorLocation = adUseClient
        
        Call rsDetalleHasar.Open(sqlDetalleHasar, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not rsDetalleHasar.State = 1 Then
            If rsDetalleHasar.EOF = True Then
            
            End If
        End If
    
        Do Until rsDetalleHasar.EOF = True
            
            Select Case EsNulo(rsImprimirHasar.Fields("Tipo").Value)
        
                Case "Fact A", "Ticket-Factura", "Fact B"
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
        
        End Select
        
    End With
    
    MousePointer = vbDefault
    
If Err Then
    Call GrabarLog("ImprimirHasar", Err.Number & " " & Err.Description, Me.Caption)
    Call MsgBox("Error Impresora:" & Err.Description, vbCritical, "Errores")
Else
   
End If
End Sub
Function ControlarApertura(vnumeroremito As Long) As Boolean
    On Error Resume Next
    
    With frmRemito.bfactura
        .Refresh
        .Recordset.Find ("remito = " & vnumeroremito)
       
        ControlarApertura = .Recordset.EOF
    
    End With

    If Err Then GrabarLog "ControlarApertura" & vnumeroremito, Err.Number & " " & Err.Description, Me.Name
End Function
Function CopiarDocumentos() As Boolean
Dim i As Long
On Error Resume Next

    With bfactura
        'En el caso que no hayan sido encontradas facturas!!!
        If .Recordset.EOF = True Then
            CopiarDocumentos = True
            Exit Function
        Else
            CopiarDocumentos = False
        End If
        
        BorrarBase "Temp_documentos", pathDBMySQL
        
        .Recordset.MoveFirst
        
        Dim rsTempDocumentos As New ADODB.Recordset, sqlTempDocumentos As String
        
        sqlTempDocumentos = "SELECT * FROM Temp_documentos WHERE 1=2"
        
        Call rsTempDocumentos.Open(sqlTempDocumentos, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        Do Until .Recordset.EOF = True
            rsTempDocumentos.AddNew
            
            For i = 0 To (.Recordset.Fields.Count - 2)
                rsTempDocumentos.Fields(i).Value = .Recordset(i).Value
                rsTempDocumentos.Fields("saldo").Value = rsTempDocumentos.Fields("Total").Value - rsTempDocumentos.Fields("Paga").Value
            Next i
            
            rsTempDocumentos.Update
            .Recordset.MoveNext
        Loop
        
    End With
    
    If rsTempDocumentos.State = 1 Then
        rsTempDocumentos.Close
        Set rsTempDocumentos = Nothing
    End If
    
If Err Then GrabarLog "CopiarDocumentos", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub cmdVerDetalle_Click()
On Error Resume Next
    Me.WindowState = 0
    MousePointer = vbHourglass
    Dim frm As Object
    Dim vrow As Long
    
    vrow = Me.KlexDocumentos.Row
    
    If vConfigGral.vIncluyeResto = False Then
        
        If cpFactura = "Factura" Then 'With frmRemito
            Set frm = Forms.Add("frmRemito")
            
        End If
        
        If cpFactura = "pFactura" Then 'With frmCompras
            Set frm = Forms.Add("frmCompras")
           ' With frmCompras
        End If
            
 With frm

            .Show
            .Habilitar (True)
            .vGrabaModo = 1  'indica que Exit Sub una eventual modificación
           ' .CargarRemito (bfactura.Recordset("Remito").Value)
           
           ' es la variable que permitirá cargar los choferes
            .vlistachofer = Right(Me.KlexDocumentos.TextMatrix(vrow, 16), Len(Me.KlexDocumentos.TextMatrix(vrow, 16)) - 1)
           
           
            .CargarRemito (Me.KlexDocumentos.TextMatrix(vrow, 13))
    
    
            '.vTipoDocumento = bfactura.Recordset("Tipo").Value
            .vTipoDocumento = Me.KlexDocumentos.TextMatrix(vrow, 20)
            
            
    
            '.txtEmpleados(0).Text = bfactura.Recordset("cod_repartidor").value
            '.txtEmpleados(1).Text = BuscarRepartidor(bfactura.Recordset("cod_repartidor").value)
            '.vncomprobante = bfactura.Recordset("Ncomprobante").value
            'If bfactura.Recordset("Tipo").value = "Documento" Then .HabilitarDocAFactura (True)
    
            .txtSubtotal.SetFocus
            .txtIva(0).SetFocus
            .txtIva(1).SetFocus
            .txtIva(2).SetFocus
            .txtPDescuento.SetFocus
            .txtDescuento.SetFocus
            .txtImpuesto.SetFocus
            .txtTotal.SetFocus
            .txtDetalle(0).SetFocus
            
            'Call .CargarChoferAGrilla(Right(Me.KlexDocumentos.TextMatrix(vrow, 16), Len(Me.KlexDocumentos.TextMatrix(vrow, 16)) - 1))
            
            
        End With
    
    Else
    
        'Panic: Controlar Que la Mesa No este Abierta y Darle Una Ubicacion Correcta Por Ejemplo Mesa Nro 1
        If Not Trim(TraerDato("TempMesas", "Remito = " & bfactura.Recordset("Remito").Value & "", "idTempMesas")) = "" Then
            MsgBox "La Mesa se Encuenta Abierta, No puede Cargarla/Modificarla desde este Modulo", vbExclamation, "Mensaje ..."
            MousePointer = vbDefault
            Exit Sub
        End If
        
        MsgBox "Nro de Mesa Asignada : " & ControlarMesasAbiertas
        
        With frmRemitoResto
            .Show
            .Habilitar (True)
            .vGrabaModo = 1
            .CargarRemito (bfactura.Recordset("remito").Value)
    
            .vTipoDocumento = bfactura.Recordset("Tipo").Value
            .txtClientes(0).Tag = bfactura.Recordset("codigo").Value
    
            .txtEmpleado(0).Text = EsNulo(bfactura.Recordset("cod_repartidor").Value)
            .txtEmpleado(1).Text = EsNulo(bfactura.Recordset("Repartidor").Value)
    
            .vnrocomprobante = bfactura.Recordset("Ncomprobante").Value
    
            If bfactura.Recordset("Tipo").Value = "Documento" Then .HabilitarDocAFactura (True)
    
            .txtSubtotal.SetFocus
            .txtIva(0).SetFocus
            .txtIva(1).SetFocus
            .txtIva(2).SetFocus
            .txtPDescuento.SetFocus
            .txtDescuento.SetFocus
            .txtImpuesto.SetFocus
            .txtTotal.SetFocus
            .txtDetalle(0).SetFocus
    
        End With
    
    End If
    
    MousePointer = vbDefault
    
   ' Unload frmRemito
    
    'If ConfigRemito(3) = True Then Unload Me
    
    If Err Then GrabarLog "cmdSeleccionar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Function ControlarMesasAbiertas() As Long
On Error Resume Next
    
    Dim rsMesasAbiertas As New ADODB.Recordset, sqlMesasAbiertas As String
    
    sqlMesasAbiertas = "SELECT M.*, TM.Remito, TM.idMesas AS NroMesa FROM Mesas M LEFT JOIN TempMesas TM ON M.idMesas=TM.idMesas ORDER BY idMesas ASC"
    
    With rsMesasAbiertas
        If .State = 1 Then .Close
        Call .Open(sqlMesasAbiertas, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If (.State = 1) And Not (.EOF = True) Then .MoveFirst
                    
        ControlarMesasAbiertas = 0
                    
        Do Until .EOF = True
            If IsNull(.Fields("Remito").Value) = True Then
                ControlarMesasAbiertas = .Fields("idMesas").Value
                Exit Do
            End If
            .MoveNext
        Loop
                    
            
    End With
    
    sqlMesasAbiertas = ""

    If rsMesasAbiertas.State = 1 Then
        rsMesasAbiertas.Close
        Set rsMesasAbiertas = Nothing
    End If
    
If Err Then GrabarLog "ControlarMesasAbiertas", Err.Number & " " & Err.Description, Me.Caption
End Function

Private Sub Command1_Click()
Dim i, j, vnro As Long
Dim vsql As String

i = 0
i = Me.KlexDocumentos.Row

vnro = Me.vncdesde

If MsgBox("Está seguro de querer cambiarlos números de comprobantes a partir del comprobante " + Me.KlexDocumentos.TextMatrix(i, 5) + " por " + Me.vncdesde, vbYesNo) = vbNo Then
    Exit Sub
End If

If Not ValFacturasImprimir Then Exit Sub


For j = i To Me.KlexDocumentos.Rows - 1

If chksolomarcados.Value Then

            If Not Me.KlexDocumentos.TextMatrix(j, 0) = "" Then
            
                        vsql = "update " + Trim(cpFactura) + " set ncomprobante = " + Str(vnro) + " where id" + Trim(cpFactura) + " = " + Me.KlexDocumentos.TextMatrix(j, 1)
                        Call EjecutarScript(vsql, pathDBMySQL)
                        
                        Me.KlexDocumentos.TextMatrix(j, 5) = vnro
                        
                        vnro = vnro + 1
            End If

Else
        vsql = "update " + Trim(cpFactura) + " set ncomprobante = " + Str(vnro) + " where id" + Trim(cpFactura) + " = " + Me.KlexDocumentos.TextMatrix(j, 1)
        Call EjecutarScript(vsql, pathDBMySQL)
            
        Me.KlexDocumentos.TextMatrix(j, 5) = vnro
            
        vnro = vnro + 1
End If


Next


End Sub

Private Sub Command2_Click()
Call fbuscarGrilla(Me.CP, "Nombre", "Codigo", Me.txtCliente.Name, Me) ' ema:
End Sub

Private Sub chkFechaTodas_Click()
On Error Resume Next

    dtpFecha(0).Enabled = Not CBool(chkFechaTodas.Value)
    dtpFecha(1).Enabled = Not CBool(chkFechaTodas.Value)
    
    dtpFecha(0).Value = Date
    dtpFecha(1).Value = Date
        
If Err Then GrabarLog "chkFechaTodas_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub chkFechaVencimiento_Click()
On Error Resume Next

    dtpFecha(2).Enabled = Not CBool(chkFechaVencimiento.Value)
    dtpFecha(3).Enabled = Not CBool(chkFechaVencimiento.Value)
    
    dtpFecha(2).Value = Date
    dtpFecha(3).Value = Date
End Sub

Private Sub dfacmensual_Click()
Dim vid2 As Long
Dim vsql As String

vid2 = Me.KlexDocumentos.TextMatrix(Me.KlexDocumentos.RowSel, 1)

vsql = "update " + Me.cpFactura + " set mensual=false where " + vid + " = " + Str(vid2)

Call EjecutarScript(vsql, pathDBMySQL)

End Sub

Private Sub dtpFecha_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
    
        Select Case Index
        
            Case 0
                dtpFecha(1).SetFocus
                
            Case 1
                cmdFiltrar.SetFocus
        
        End Select
    
    End If

If Err Then GrabarLog "dtpFecha_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub factmensual_Click()
Dim vid2 As Long
Dim vsql As String

vid2 = Me.KlexDocumentos.TextMatrix(Me.KlexDocumentos.RowSel, 1)

vsql = "update " + Me.cpFactura + " set mensual=true where " + vid + " = " + Str(vid2)

Call EjecutarScript(vsql, pathDBMySQL)

End Sub

Private Sub FlaVerDatos_Change()
Call ver_datos_afip
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbShiftMask Then

    vshift = 1

End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbShiftMask Then

    vshift = 0

End If
End Sub

Public Sub Form_Load()
    On Error Resume Next

    With Me
        .Top = 0
        .Left = 0
        .Width = 17970
        .Height = 8610
        .dtpFecha(0).Enabled = True
        .dtpFecha(1).Enabled = True
        .dtpFecha(0).Value = Date
        .dtpFecha(1).Value = Date
    End With

    If viene = "ctacte" Then
        bfactura.RecordSource = "SELECT * FROM " + cpFactura + " WHERE (remito = " & vVieneRemito & ")"
        bfactura.Refresh
        viene = ""
    Else
       ' cmdFiltrar_Click
    End If
    
    Call CentrarFormulario(Me)
    

    finit

    
    If Err Then GrabarLog "Form_load", Err.Number & " " & Err.Description, Me.Name
End Sub

Public Sub finit()
' ema: las init van siempre con esta función


Me.log2.Visible = False
Me.log2.Cols = 10

Me.tab.SelectedItem = 0

'---------
Me.cmbEstadoDocumento.Clear
Me.cmbEstadoDocumento.AddItem "Adeudado"
Me.cmbEstadoDocumento.AddItem "Pagado"
Me.cmbEstadoDocumento.AddItem "Quebranto"
Me.cmbEstadoDocumento.AddItem "Pendiente"
Me.cmbEstadoDocumento.AddItem "Todos los estados"
Me.cmbEstadoDocumento.Text = "Todos los estados"

'---------------------------
Me.cmbTiposPedidos.Clear
Me.cmbTiposPedidos.AddItem "Retirado"
Me.cmbTiposPedidos.AddItem "No Retirado"
Me.cmbTiposPedidos.AddItem "Recambio"
Me.cmbTiposPedidos.AddItem "Todos"
Me.cmbTiposPedidos.Text = "Todos"
'---------------------------

Me.vordenadoPor.Clear
Me.vordenadoPor.AddItem "Importe"
Me.vordenadoPor.AddItem "Fecha"
Me.vordenadoPor.AddItem "Codigo"
Me.vordenadoPor = "Fecha"
'---------------------------

finitConciliacion
    
    If cpFactura = "factura" Then
    
        vid = "idFactura"
        cpFactura = "Factura"
    Else
        vid = "idPFactura"
        cpFactura = "pFactura"
    End If

If UCase(LeerXml("Puesto")) = "ASOCIAL" Then
    Me.chkPresupuesto.Caption = "Ficha Social"
End If


End Sub
Private Sub finitConciliacion()
On Error Resume Next
Dim r As New ADODB.Recordset
Dim vsql As String


vsql = "select a.codigo, a.nombre, format(a.saldo,'###,###,##0.00') , format(b.saldo,'###,###,##0.00')  from ((select codigo,nombre,sum(debito) - sum(credito) as saldo from pcuentascorrientes group by codigo) as a)" + _
" Inner Join " + _
" ((select codigo, sum(t.total) - sum(t.pagoparcial) as saldo from pfactura t group by codigo ) as b) " + _
" on a.codigo = b.codigo " + _
" Where abs(a.saldo - b.saldo) > 1"

Call r.Open(vsql, ConnDDBB, adOpenStatic, adLockReadOnly)

Set Me.gconciliacion.DataSource = r.DataSource

Me.lblerrores.Caption = "Hay " + Str(r.RecordCount) + " errores de conciliación. Verifique cada una de las cuentas de proveedores"

If Err Then Exit Sub
End Sub

Private Sub FormatoGrilla(vCantidadRenglones As Long)
On Error Resume Next

    Dim i As Long

    With KlexDocumentos
        .FixedRows = 1
        .FixedCols = 0
           
        .Cols = 28
        .Rows = vCantidadRenglones + 1
        
        If vCantidadRenglones = 1 Then
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
                .ColWidth(i) = 0
            Next
        End If
        
        .TextMatrix(0, 0) = "Sel"
        .ColWidth(0) = 250
        
        .TextMatrix(0, 1) = "idFactura"
        .ColWidth(1) = 0
               
        .TextMatrix(0, 2) = "Tipo"
        .ColWidth(2) = 500
        
        .TextMatrix(0, 3) = "Letra"
        .ColWidth(3) = 500
        
        .TextMatrix(0, 4) = "P. De Venta"
        .ColWidth(4) = 500
        
        .TextMatrix(0, 5) = "Nº Comp."
        .ColWidth(5) = 750
        
        .TextMatrix(0, 6) = "Fecha"
        .ColWidth(6) = 1000
        
        .TextMatrix(0, 7) = "Cod. Cliente"
        .ColWidth(7) = 1250
        
        .TextMatrix(0, 8) = "Cliente"
        .ColWidth(8) = 2500
        
        .TextMatrix(0, 9) = "Cuit."
        .ColWidth(9) = 1500
        
        .TextMatrix(0, 10) = "Total"
        .ColWidth(10) = 1000
        '.ColDisplayFormat(10) = "##,##0.00"
        .ColAlignment(10) = 6
                
        .TextMatrix(0, 11) = "Nro.Interno"
        .ColWidth(11) = 750
        
        .TextMatrix(0, 12) = "Observ."
        .ColWidth(12) = 1000
        
        .TextMatrix(0, 13) = "Remito"
        .ColWidth(13) = 0
        
        .TextMatrix(0, 14) = "Asiento"
        .ColWidth(14) = 0
        
        .TextMatrix(0, 15) = "Est.Doc."
        .ColWidth(15) = 600
        
        .TextMatrix(0, 16) = "Choferes"
        .ColWidth(16) = 600
        
        .TextMatrix(0, 17) = "Volquetes"
        .ColWidth(17) = 600
        
        
        .TextMatrix(0, 18) = "Devol."
        .ColWidth(18) = 600
        
        
        .TextMatrix(0, 19) = "Est.Volq."
        .ColWidth(18) = 600
        
        .TextMatrix(0, 20) = "Pagos"
        .ColWidth(19) = 600
                '.ColAlignment(19) = 1
        '.ColDisplayFormat(19) = "##,##0.00"
        
        
        .TextMatrix(0, 21) = "T.Doc."
        .ColWidth(21) = 600
        
        .TextMatrix(0, 22) = "Saldos"
        .ColWidth(22) = 1200
        
        
        .TextMatrix(0, 23) = "F.Pago"
        .ColWidth(23) = 1200
        
        .TextMatrix(0, 24) = "Mensual"
        .ColWidth(24) = 500
        
        .TextMatrix(0, 25) = "Empresa"
        .ColWidth(24) = 500
        
        .TextMatrix(0, 26) = "Repartidor"
        .ColWidth(24) = 500
        
        '.TextMatrix(0, 18) = "Observaciones"
        '.ColWidth(18) = 2500
        
        '.TextMatrix(0, 17) = "Importe"
        '.ColWidth(17) = 1500
        '.FormatString = "##,##0.00"
        '.ColDisplayFormat(17) = "#0.000"
        '.ColAlignment(17) = vbAlignRight



       ' .BackColorAlternate = 14737632
    End With
    
If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub ver_datos_afip()
Dim vrow, vcol, vnro_sucursal, vnro_comprobante, vnro_tipo_doc As Integer
Dim vtipo_doc, vLetra As String


vrow = Me.KlexDocumentos.Row
vnro_comprobante = Me.KlexDocumentos.TextMatrix(vrow, 5)
vnro_sucursal = Me.KlexDocumentos.TextMatrix(vrow, 4)
vtipo_doc = Me.KlexDocumentos.TextMatrix(vrow, 2)
vLetra = Me.KlexDocumentos.TextMatrix(vrow, 2)

If vtipo_doc = "Fact A" And vLetra = "A" Then vnro_tipo_doc = 1
If vtipo_doc = "Nota D" And vLetra = "A" Then vnro_tipo_doc = 2
If vtipo_doc = "Nota C" And vLetra = "A" Then vnro_tipo_doc = 3
If vtipo_doc = "Recibo" And vLetra = "A" Then vnro_tipo_doc = 4

If vtipo_doc = "Fact B" And vLetra = "B" Then vnro_tipo_doc = 6
If vtipo_doc = "Nota D" And vLetra = "B" Then vnro_tipo_doc = 7
If vtipo_doc = "Nota C" And vLetra = "B" Then vnro_tipo_doc = 8
If vtipo_doc = "Recibo" And vLetra = "B" Then vnro_tipo_doc = 9


If vtipo_doc = "Fact C" And vLetra = "C" Then vnro_tipo_doc = 11
If vtipo_doc = "Nota D" And vLetra = "C" Then vnro_tipo_doc = 12
If vtipo_doc = "Nota C" And vLetra = "C" Then vnro_tipo_doc = 13
If vtipo_doc = "Recibo" And vLetra = "C" Then vnro_tipo_doc = 15


With frmPrincipal.fe
        Dim bResultado As Boolean
        
        bResultado = .F1CompConsultarS(CInt(vnro_sucursal), vtipo_doc, vnro_comprobante)
       
       
       If .UltimoMensajeError = "" Then
          MsgBox "CAE consultado: " + .F1RespuestaDetalleCae + " Fecha Vto : " + .F1DetalleCbteFch + Chr(13) _
          + "Total: " + Str(.F1DetalleImpTotal) + Chr(13) _
          + "Si este comprobante no está ingresado en el sistema, UD. debe REINGRESAR FACTURA poniendo de modo manual el CAE y la Fecha que aquì se indica"
       Else
          MsgBox ("fallo consulta: " + .UltimoMensajeError)
       End If
       
End With


End Sub




Private Sub KlexDocumentos_DblClick()
On Error Resume Next

Dim vValor As Double
Dim i, vr As Long

i = Me.KlexDocumentos.Row

If Not Me.viene = "cobro" And Not Me.viene = "ie" Then

    vr = Me.KlexDocumentos.Row

    If Trim(Me.KlexDocumentos.TextMatrix(i, 0)) = "" Then
        pintar (i)
    Else
        despintar (i)
    End If
Exit Sub
End If

Me.KlexDocumentos.Col = 0

If (Me.KlexDocumentos.CellBackColor = vbGreen) Then
    Me.KlexDocumentos.CellBackColor = vbWhite
    Me.KlexDocumentos.Text = ""
    vdocmarcados = vdocmarcados - 1
Else
    Me.KlexDocumentos.CellBackColor = vbGreen
     
        vValor = Val(KlexDocumentos.TextMatrix(Me.KlexDocumentos.Row, 10)) - Val(KlexDocumentos.TextMatrix(Me.KlexDocumentos.Row, 20))
    
        If vValor < 0.01 Then
        
            Me.KlexDocumentos.Text = "-"
            vdocmarcados = vdocmarcados + 1
            
        Else
            Me.KlexDocumentos.Text = "X"
            vdocmarcados = vdocmarcados + 1
          '  Me.KlexDocumentos.TextMatrix(i, 20) = Me.KlexDocumentos.TextMatrix(i, 10)
          '  Me.KlexDocumentos.TextMatrix(i, 22) = 0
    End If
End If

Call KlexDocumentos_SelChange

Me.vtotal.Caption = Str(totalesSeleccionados)
If Err Then Exit Sub
End Sub

Function totalesSeleccionados() As Double
On Error Resume Next
Dim i, j, k, l As Long

Dim vtotal As Double
vvpagoPacial = False

With KlexDocumentos

l = 0
j = .Rows

k = 10

vtotal = 0
vdocseleccionados = ""

For i = 1 To j - 1


If .TextMatrix(i, 0) = "X" Then

    vtotal = vtotal + Val(.TextMatrix(i, k)) - Val(.TextMatrix(i, 20))
     
    vdocseleccionados = vdocseleccionados + " " + .TextMatrix(i, 2) + " " + .TextMatrix(i, 4) + " " + .TextMatrix(i, 5) + " - "
    l = l + 1
    vsqlpago(l) = ""
    vsqlpago(l) = "update " + cpFactura + " set estadodocumento='Pagado', pagoparcial=" + Str(.TextMatrix(i, 10)) + ",saldos=" + Str(.TextMatrix(i, 10)) + " where id" + cpFactura + "=" + .TextMatrix(i, 1)

    ' todoing para recibo pago2
        
    
        '------------ aca puedo preparar el arreglo para el recigo tabla pago2 --------------------- todoing
        vpap2 = "insert into pago2 (comprobante,fechaemision,importe,nroorden) values ("
        
        vvalorespap2 = "'" + Trim(.TextMatrix(i, 2)) + " - " + Trim(.TextMatrix(i, 3)) + " - " + Trim(.TextMatrix(i, 4)) + " - " + Trim(.TextMatrix(i, 5)) + "'," + _
                       "'" + Trim(.TextMatrix(i, 6)) + "'," + _
                       "" + Trim(.TextMatrix(i, 10)) + "," + _
                       "99"

      '  varreglo_pago2(l) = vpap2 + vvalorespap2 + ")"
        
        
        
       vvalorespap2_temp = "     > " + fc(Trim(.TextMatrix(i, 2)), 15) + "  " + fc(Trim(.TextMatrix(i, 4)), 4) + " - " + fc(Trim(.TextMatrix(i, 5)), 10) + "               " + fc(.TextMatrix(i, 6), 10)
        
        
      varreglo_pago2_temp(l) = "insert into temp2 (c02,c05) values ('" + vvalorespap2_temp + "', '" + Format(.TextMatrix(i, 10), "###,###,##0.00") + "')"
        
        
       
        
        
        '-------------------------------------------------------------------------------------------

    
    

End If

If .TextMatrix(i, 0) = "-" Then
   vvpagoPacial = True
End If

Next

End With

totalesSeleccionados = vtotal


Dim v2 As String

v2 = "'','-------'"
vsql = "insert into temp2 (c02,c05) values (" + v2 + ")"


'varreglo_pago2_temp(l + 1) = vsql


v2 = "'','" + Format(vtotal, "###,###,##0.00")
vsql = "insert into temp2 (c02,c05) values (" + v2 + "')"

'varreglo_pago2_temp(l + 2) = vsql


If Err Then Exit Function
End Function

Public Sub KlexDocumentos_DblClick2()
On Error Resume Next

'    If vieneCobro = True Then
'        CmdEjecutarCobro_Click
'    Else
'        cmdVerDetalle_Click
'    End If
    
    If Err Then GrabarLog "klexDocumentos_DblClick", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub KlexDocumentos_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If this is not row 0, do nothing.
    If KlexDocumentos.MouseRow <> 0 Then Exit Sub

    ' Sort by the clicked column.
    SortByColumn KlexDocumentos.MouseCol, KlexDocumentos
End Sub

Private Sub KlexDocumentos_SelChange()
On Error Resume Next

    Me.KlexDocumentos.RowSel = KlexDocumentos.Row

If Err Then Exit Sub
End Sub

Private Sub PusCerrar_Click(Index As Integer)
On Error Resume Next
    If Me.vieneCobro Then pasaDocCobros
    Unload Me
    
If Err Then GrabarLog "cmdSalir_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub pasaDocCobros()
'frmCobros.vsqlpago = vsqlpago
'frmCobros.vdocSeleccionados = vdocSeleccionados
'frmCobros.vtotalDocSeleccionado = Val(vtotal.Caption)
End Sub

Private Sub PusDetalles_Click(Index As Integer)
Dim vsql2, vsql, vwhere, fd, fh, vtipo As String



    vsql = " and (1>1 "

    If chkFacturaA.Value = 1 Then vsql = vsql + " OR (tipo = 'Fact A')"
    If chkMonotributo.Value = 1 Then vsql = vsql + " OR (tipo = 'Fact B')"
    If chkFacturaC.Value = 1 Then vsql = vsql + " OR (tipo = 'Fact C')"
    If chkNotasDe.Value = 1 Then vsql = vsql + " OR (Tipo = 'Nota D')"
    If chkRemito.Value = 1 Then vsql = vsql + " OR (Tipo = 'Remito')"
    If chkDocNo.Value = 1 Then vsql = vsql + " OR (Tipo = 'Documento')"
    If chkPresupuesto.Value = 1 Then vsql = vsql + " OR (Tipo = 'Presupuesto')"
    If chkNotaCA.Value = 1 Then vsql = vsql + " OR (tipo = 'Nota C' AND (Letra = 'A'))"
    If chkNotaCB.Value = 1 Then vsql = vsql + " OR (tipo = 'Nota C' AND (Letra = 'B'))"
    If chkNotaCC.Value = 1 Then vsql = vsql + " OR (tipo = 'Nota C' AND (Letra = 'C'))"
    
   
    
    
    vsql = vsql + ")"
    
    If Me.rbdc.Value Then
         vsql = vsql + " and (Cod_repartidor = '1') "
    End If
    
    If Me.rbdl.Value Then
        vsql = vsql + " and (Cod_repartidor = '2') "
    End If
    
    If Not txtCodigoCliente.Text = "" Then
        vsql = vsql + " and (codigo = '" + txtCodigoCliente.Text + "')"
    End If
    
    If Not Me.txtNDesde.Text = "" Then
    
        vsql = vsql + " and  (NComprobante >= " + txtNDesde.Text + ")"
    
    End If
    

    If Not Me.txtNHasta.Text = "" Then
    
        vsql = vsql + " and  (NComprobante <= " + txtNHasta.Text + ")"
    
    End If
    

fd = strfechaMySQL(Me.dtpFecha(0))

fh = strfechaMySQL(Me.dtpFecha(1))

vwhere = "(fecha >= '" + fd + "' and fecha <= '" + fh + "') " + vsql

vsql2 = " SHAPE {SELECT * FROM `factura` where " + vwhere + " order by NComprobante asc } AS FacturaFDetalle " + _
" APPEND ({SELECT * FROM `fdetalle`}  AS FDetalle " + _
" RELATE 'Remito' TO 'Remito') AS FDetalle "


With Mantenimiento

   If .rsFacturaFDetalle.State = 1 Then .rsFacturaFDetalle.Close
   .rsFacturaFDetalle.Source = vsql2
   If .rsFacturaFDetalle.State = 0 Then .rsFacturaFDetalle.Open
            .rsFacturaFDetalle.Close
            .rsFacturaFDetalle.Open
End With

drFacturaDetalle2.Show

End Sub

Private Sub PusExell_Click()
On Error Resume Next
    
  Call grillaToExcel(Me.KlexDocumentos)

If Err Then Exit Sub
End Sub

Private Sub PushButton1_Click()
Call fbuscarGrilla("empleados", "Nombre", "Codigo", Me.txtEmpleado.Name, Me) ' ema:
End Sub

Private Sub PushButton10_Click()
sacarMarcas

Call cmdFiltrar_Click

pagarDocumentos (Val(Me.vImporteSeleccionado.Tag))

Me.vtotal.Caption = Str(totalesSeleccionados)



Call formatGrillaDoc(10, "##########0.00")
Call formatGrillaDoc(20, "##########0.00")

End Sub

Private Sub pagarDocumentos(vimporte As Double)
On Error Resume Next
Dim i, j, k As Long

Dim vtotal As Double

With KlexDocumentos
    j = .Rows
    k = 0
    vtotal = 0
    vdocseleccionados = ""

For i = 1 To j
    .Row = i
    .Col = k
    
    vtotal = Val(.TextMatrix(i, 10)) - Val(.TextMatrix(i, 20))
    
    If Not vtotal = 0 And Not .TextMatrix(i, 0) = "X" Then
    
            If vimporte > vtotal Then
            
                        vimporte = vimporte - vtotal
                               
                        .TextMatrix(i, 15) = "Pagado"
                        '.TextMatrix(i, 20) = Val(.TextMatrix(i, 10))
                               
                        vsqlpagoAuto(i) = ""
                        vsqlpagoAuto(i) = "update " + cpFactura + " set estadodocumento='Pagado', pagoparcial=" + Str(.TextMatrix(i, 10)) + ",saldos=0  where id" + Trim(cpFactura) + "=" + .TextMatrix(i, 1)
                    
                        vdocseleccionados = vdocseleccionados + " " + Trim(.TextMatrix(i, 21)) + " " + .TextMatrix(i, 2) + " " + .TextMatrix(i, 3) + " " + .TextMatrix(i, 4) + " " + .TextMatrix(i, 5) + " - "
                
                
                        .Col = 0
                        .CellBackColor = vbBlue
                        Me.KlexDocumentos.Text = "X"
                
            Else
            
                            If Not vimporte <= 0 Then
                              
                                            .TextMatrix(i, 15) = "Parcial"
                                            .TextMatrix(i, 20) = vimporte + Val(.TextMatrix(i, 20))
                                            .TextMatrix(i, 22) = 0
                                            
                                            vsqlpagoAuto(i) = ""
                                            vsqlpagoAuto(i) = "update " + cpFactura + " set estadodocumento='Parcial', pagoparcial=" + Str(Val(.TextMatrix(i, 20))) + ",saldos=" + Str(vtotal - vimporte) + " where id" + Trim(cpFactura) + "=" + Str(.TextMatrix(i, 1))
                                    
                                     
                                            .Col = 0
                                            .Row = i
                                            .CellBackColor = vbBlue
                                             vimporte = vimporte - vtotal
                                            Me.KlexDocumentos.Text = "X"
                                            vdocseleccionados = vdocseleccionados + " " + Trim(.TextMatrix(i, 21)) + " " + .TextMatrix(i, 2) + " " + .TextMatrix(i, 3) + " " + .TextMatrix(i, 4) + " " + .TextMatrix(i, 5) + " - "
                    
                            End If
                    
            End If
    
    
    
    Else
    
        Me.KlexDocumentos.Text = "-"
        '.CellBackColor = vbBlue
        
    End If
   ' .TextMatrix(i, 0) = ""
    '.CellBackColor = vbWhite
Next

End With


If Err Then Exit Sub
End Sub


Private Sub sacarMarcas()
On Error Resume Next
Dim i, j, k As Long

Dim vtotal As Double

With KlexDocumentos
    j = .Rows
    k = 0
    vtotal = 0
    vdocseleccionados = ""

For i = 1 To j
    .Row = i
    .Col = k
    .TextMatrix(i, 0) = ""
    .CellBackColor = vbWhite
Next

End With


If Err Then Exit Sub
End Sub

Private Sub PushButton11_Click()


If MsgBox("Confirma la impresión de las facturas seleccionadas", vbYesNo) = vbYes Then

    vvmes = Month(Me.KlexDocumentos.TextMatrix(1, 6))
    vvano = Year((Me.KlexDocumentos.TextMatrix(1, 6)))

    finitImpresion
    'TODO: Enter task description here
    facturasImprimir
    
    If MsgBox("Presiones aceptar para que se descuente el stock de los documentos impresos", vbOKCancel) = vbOK Then
        Call ejecutar_actualizar_stock
    End If

    
    End If

End Sub

Private Sub ejecutar_actualizar_stock()
Dim i As Long
Dim vsql As String

Me.barra.Max = UBound(arr_as)
Me.barra.Value = 0

For i = 1 To UBound(arr_as)
    vsql = ""
    vsql = arr_as(i)
    If Not vsql = "" Then Call EjecutarScript(vsql, pathDBMySQL)
    arr_as(i) = ""
    Me.barra.Value = i
Next

    'ReDim arr_as(500)


End Sub

Private Sub finitImpresion()
    Printer.PaperSize = 9
End Sub


Private Sub facturasImprimir()
Dim i As Long
Dim vremito As Long
'------------------------------------------
If Not ValFacturasImprimir Then Exit Sub
If Not validariImpreFE Then Exit Sub
'------------------------------------------

barra2.Value = 0
barra2.Max = Me.KlexDocumentos.Rows - 1

cadena_errores = ""

For i = 1 To Me.KlexDocumentos.Rows - 1
    
    If Me.KlexDocumentos.TextMatrix(i, 0) = "I" Then
    
        vsql = "select remito from factura where idfactura =" + Me.KlexDocumentos.TextMatrix(i, 1)
        vremito = Val(traerDatos2(vsql, "remito", pathDBMySQL))
    
        vgrow = i
        Call imprimeUnaFactura(vremito, Me.KlexDocumentos.TextMatrix(i, 21), Me.KlexDocumentos.TextMatrix(i, 5))
        
        If errexit Then Exit For
        
        Wait (1000)
        
    End If
    
    barra2.Value = barra2.Value + 1
Next

MsgBox "Terminó la serie de impresión", vbInformation


If Not cadena_errores = "" Then
    MsgBox "Documentos que no se pudieron imprimir: " + cadena_errores
End If

End Sub

Function ValFacturasImprimir() As Boolean
On Error Resume Next
Dim i As Long
Dim vtipo As String

vtipo = Me.KlexDocumentos.TextMatrix(1, 21)
ValFacturasImprimir = True


For i = 1 To Me.KlexDocumentos.Rows - 1

    If Not vtipo = Me.KlexDocumentos.TextMatrix(i, 21) Then
    
        MsgBox "Hay distintos tipos de documentos para realizar esta operación. " + Chr(13) + "Debe filtrar solamente un tipo de documento"
        ValFacturasImprimir = False
        i = Me.KlexDocumentos.Rows - 1
    End If
    
Next


If Err Then
    Exit Function
    ValFacturasImprimir = False
End If

End Function

Private Sub guardarFdetalleTemp(vremito As Long)
Dim vValor, vcampos, vsql As String
Dim i, vrow As Long

Call BorrarBase("Documentos", PathDBListados)

bdetalle.ConnectionString = pathDBMySQL

bdetalle.RecordSource = "select * from fdetalle where remito=" + Str(vremito)
bdetalle.Refresh


vcampos = ""
vValor = ""

With bdetalle.Recordset

.MoveFirst
Do Until .EOF

        
        vValor = Str(.Fields("cantidad")) + ",'" + (.Fields("codigo")) + "','" + (.Fields("detalle")) + "'," + Str(.Fields("precio")) + ",0,0," + Str(.Fields("total"))
        'vValor = .TextMatrix(i, 5) + ",'" + Replace(Replace(.TextMatrix(i, 4), "[", ""), "]", "") + "','" + .TextMatrix(i, 6) + "'," + Str(Val(.TextMatrix(i, 7))) + "," + Str(Val(.TextMatrix(i, 8))) + "," + Str(Val(.TextMatrix(i, 9))) + "," + Str(Val(.TextMatrix(i, 11)))
        vcampos = "cantidad,codigo,descripcion,pventa,descuento,iva,total"
        
        vsql = "insert into Documentos (" + vcampos + ") values (" + vValor + ")"
        
        Call EjecutarScript(vsql, PathDBListados)

.MoveNext

Loop


End With


End Sub


Private Sub setMarcarImpresa(vremito As Long)
    Dim vsql As String
    
    vsql = "update factura set estado2= 'Impreso' where remito =" + Str(vremito)
    Call EjecutarScript(vsql, pathDBMySQL)
End Sub



Private Sub imprimeUnaFactura(vremito As Long, vTipoDocumento As String, vncomprobante As Long)
        
        Dim i, t As Long
       
        setFacturaDatos (vremito) ' lo pongo en un dataset
        
      If errexit Then
                cadena_errores = cadena_errores + " - " + Str(vremito) + " -" + vTipoDocumento + Str(vncomprobante) + Chr(13)
                Exit Sub
      End If
        
       Call llenarDetalles2(vremito, vTipoDocumento)
        
      If errexit Then Exit Sub
        
        setMarcarImpresa (vremito)
        
     If errexit Then Exit Sub
  
  'guardarFdetalleTemp (vremito)
        
        'Unload Mantenimiento3
        'Load Mantenimiento3
        
        
        'Me.WindowState = vbMinimized
                
        
        Call mostrar_Doc2(vTipoDocumento, vncomprobante)
        
        
'        Select Case vTipoDocumento
'
'            Case "Fact A"
'                'Call ifacta.Hide
'                'Call ifacta.PrintReport(False, rptRangeAllPages)
'
'                'mostrar_ifactaLabel
'
'                Call mostrar_Doc2(vTipoDocumento, vncomprobante)
'
'               ' ifacta.Show
'
'                'Call ImprimirTicket(Mantenimiento.rscfact.Fields("Remito").Value)
'
'
'            Case "Fact B"
'               ' imonotributo.Show
'                 mostrar_ifactb
'                 'ifacta.Show
'
'            Case "Presupuesto"
'                'ipresupuesto.Show
'
'                'mostrar_documentos
'                'idocumento.Show
'
'            Case "Remito"
'                 mostrar_remito
'
'            Case "Nota C"
'                'mostrar_documentos (iNotaCredito)
'
'                'mostrar_iNotaCredito
'                'inotac.Show
'                 mostrar_ifactaLabel
'
'            Case "Documento"
'
'                    'Call idocumento.PrintReport(False, rptRangeAllPages)
'                    mostrar_documentos (vncomprobante)
'                    'idocumento.Refresh
'
'            Case Else
'                mostrar_ifactaLabel
'
'        End Select
    
        MousePointer = vbDefault
    
        'Me.WindowState = 1

    
 
   ' NuevoCliente



End Sub

Private Sub setFacturaDatos(vremito As Long)
On Error Resume Next
Dim vsql, vsql2, vcodigo  As String
Dim vnrointerno As Long
Dim vcae, vcaeVto, vcuit  As String

vsql = "select * from factura inner join ivafacturaventa on factura.NroInterno = ivafacturaventa.nrointerno  where factura.remito = " + Str(vremito)

vsql = "select * from (select * from factura where year(fecha)= " + Str(vvano) + _
" and month(fecha) = " + Str(vvmes) + ") as factura inner join ivafacturaventa on factura.NroInterno = ivafacturaventa.nrointerno  where factura.remito = " + Str(vremito)

vcodigo = traerDatos2(vsql, "codigo", pathDBMySQL)
vsql2 = "select * from clientes  where codigo = " + Str(vcodigo)

vnrointerno = Val(traerDatos2(vsql, "nrointerno", pathDBMySQL))
vf.vcodigo = traerDatos2(vsql, "codigo", pathDBMySQL)
vf.vfecha = CDate(traerDatos2(vsql, "fecha", pathDBMySQL))
vf.vnombre = traerDatos2(vsql, "nombre", pathDBMySQL)
vf.vdireccion = traerDatos2(vsql2, "direccion", pathDBMySQL)
vf.vlocalidad = traerDatos2(vsql2, "localidad", pathDBMySQL)
'vf.vCuit = traerDatos2(vsql2, "cuit", pathDBMySQL)
vcuit = traerDatos2("select cuit from factura where remito = " + Str(vremito), "cuit", pathDBMySQL)
vf.vcuit = Replace(Replace(vcuit, "-", ""), " ", "")
vf.vtotal = Val(traerDatos2(vsql, "total", pathDBMySQL))
vf.vSubTotal = Val(traerDatos2(vsql, "subtotal", pathDBMySQL))
vf.vIva = traerDatos2(vsql, "Iva", pathDBMySQL)
vf.vsaldo = getSaldoProveedor3(vremito)

vf.vcae = "0"   ' le pongo cero por las dudas de que qde cargado con otro nro
vf.vcaeVto = ""

vf.vcae = traerDatos2(vsql, "cae", pathDBMySQL)
vf.vcaeVto = traerDatos2(vsql, "caevto", pathDBMySQL)

vf.viva150 = Val(traerDatos2(vsql, "iva105", pathDBMySQL))
vf.vIva210 = Val(traerDatos2(vsql, "iva210", pathDBMySQL))
vf.viva270 = Val(traerDatos2(vsql, "iva270", pathDBMySQL))

'Call fecae2(fe, vtipoFactura, Str(Val(Me.txtNroComprobante)), Me.txtSubtotal, Me.txtTotal, vCuit, _
Format(Me.dtpFecha, "yyyymmdd"), Val(Me.cboPuntoDeVenta.Text), Me.txtNroInterno, vc1, vc2, Val(Me.txtIva(0).Text), Val(Me.txtIva(1).Text), Val(Me.txtIva(2).Text), vmodotest, Val(txtNroInterno))

vTipoComprobante = getTipoComprobante(Me.KlexDocumentos.TextMatrix(vgrow, 21))
venro = Me.KlexDocumentos.TextMatrix(vgrow, 5)
veptovta = Str(Val(Me.KlexDocumentos.TextMatrix(vgrow, 4)))

vnroCodigoBarra = feCodigoBarra(vf.vcuit, vTipoComprobante, Format(veptovta, "0000"), vcae, vcaeVto)


errexit = False
If vTipoComprobante = 99 Then Exit Sub

    '' obtengo el cae si es que no está calculado porque estoy reimprimiendo

If LeerXml("ObtieneCAE") = "NO" Then vf.vcae = "1"

    If Not Val(vf.vcae) > 0 Then
                Call fecae2(vfe, vTipoComprobante, Me.KlexDocumentos.TextMatrix(vgrow, 5), _
                vf.vSubTotal, vf.vtotal, Trim(Replace(vf.vcuit, "-", "")), Trim(Format(vf.vfecha, "yyyymmdd")), Str(Val(Me.KlexDocumentos.TextMatrix(vgrow, 4))), _
                "1", vf.vcae, vf.vcaeVto, vf.viva150, vf.vIva210, vf.viva270, LeerXml("modoFiscal"), 1, Me.Name)
    End If


' vn_anterior  = getnroconafip()
'


If Val(vf.vcae) = 0 Then
    errexit = True
Else
    errexit = False
End If

vsql = "update factura set cae='" + vf.vcae + "', caevto='" + vf.vcaeVto + "' where remito = " + Str(vremito)
Call EjecutarScript(vsql, pathDBMySQL)

Call log(vf.vcae + " - " + vf.vcaeVto)
'Call log(vf.vcae + " - " + vf.vcaeVto + " " + (Format(vf.vfecha, "yyyymmdd")) + " " + f.vCuit)

If Err Then Exit Sub
End Sub


Function getSaldoProveedor3(ByVal vremito As String) As Double
On Error Resume Next
Dim vsql As String

getSaldoProveedor3 = 0

vsql = "select saldos from factura where remito = '" + vremito + "'"
getSaldoProveedor3 = traerDatos2(vsql, "saldos", pathDBMySQL)

If Err Then Exit Function
End Function

Function getTipoComprobante(vdato As String) As Long
getTipoComprobante = 99

Select Case vdato

    Case "Fact A"
        veDocumento = "Factura"
        getTipoComprobante = 1
        veLetra = "A"
    Case "Fact B"
         veDocumento = "Factura"
        getTipoComprobante = 6
        veLetra = "B"
    Case "Fact C"
        veDocumento = "Factura"
        getTipoComprobante = 11
        veLetra = "C"
End Select

End Function

Private Sub PushButton12_Click()
Dim i As Long

For i = 1 To Me.KlexDocumentos.Rows - 1

    pintar (i)

Next

End Sub

Private Sub pintar(i As Long)
Me.KlexDocumentos.TextMatrix(i, 0) = "I"
Me.KlexDocumentos.Row = i
Me.KlexDocumentos.Col = 0
Me.KlexDocumentos.CellBackColor = vbYellow + 35
Me.vdocmarcados = vdocmarcados + 1

End Sub

Private Sub despintar(i As Long)
Me.KlexDocumentos.TextMatrix(i, 0) = ""
Me.KlexDocumentos.Row = i
Me.KlexDocumentos.Col = 0
Me.KlexDocumentos.CellBackColor = vbWhite
Me.vdocmarcados = vdocmarcados - 1
End Sub

Private Sub PushButton13_Click()
Dim i As Long

For i = 1 To Me.KlexDocumentos.Rows - 1

    despintar (i)

Next
End Sub

Public Sub PushButton14_Click()
chkFacturaA.Value = False
chkNotasDe.Value = False
chkNotaCA.Value = False
chkNotaCC.Value = False
chkMonotributo.Value = False
chkPresupuesto.Value = False
chkNotaCB.Value = False
chkDocNo.Value = False
chkRemito.Value = False
chkOtros.Value = False
chkFacturaC.Value = False
chkFacturaX.Value = False

End Sub

Private Sub PushButton15_Click()
chkFacturaA.Value = 1
chkNotasDe.Value = 1
chkNotaCA.Value = 1
chkMonotributo.Value = 1
chkPresupuesto.Value = 1
chkNotaCB.Value = 1
chkDocNo.Value = 1
chkRemito.Value = 1
chkOtros.Value = 1
chkFacturaC.Value = 1
chkFacturaX.Value = 1
chkNotaCC.Value = 1
End Sub

Private Sub PushButton16_Click()
Call seleccionartodo(True)
End Sub

Private Sub PushButton17_Click()
Call seleccionartodo(False)
End Sub

Private Sub PushButton18_Click()
   frmFeStatus.Show
End Sub

Private Sub PushButton19_Click()

If log2.Visible = True Then
  log2.Visible = False
Else
  log2.Visible = True
End If

End Sub

Private Sub PushButton20_Click()
Call fbuscarGrilla("rubros", "Rubro", "idRubros", Me.vrubro.Name, Me) ' ema:
End Sub

Private Sub PushButton21_Click()
Call fbuscarGrilla("subrubros", "SubRubros", "idsubrubros", Me.vsubrubro.Name, Me)
End Sub

Private Sub PushButton23_Click()
Dim vsql, vc1, vc2 As String

vsql = "(Select * from proveedores where tipocliente  = 'Empresa') t"
vc1 = "Nombre"
vc2 = "Codigo"

Call fbuscarGrilla(vsql, vc1, vc2, Me.vdescEmpresa.Name, Me)
End Sub

Private Sub PushButton24_Click()
Dim vsql, vc1, vc2 As String

vsql = "(Select * from proveedores where tipocliente  = 'Vendedor') t"
vc1 = "Nombre"
vc2 = "Codigo"

Call fbuscarGrilla(vsql, vc1, vc2, Me.vDesRepartidor.Name, Me)
End Sub

Private Sub PushButton3_Click()

If Val(Me.KlexDocumentos.TextMatrix(Me.KlexDocumentos.Row, 1)) > 0 Then
    
    frmVolquetesAdmin.vtablaFactura = Me.cpFactura
    frmVolquetesAdmin.vIdFactura = "id" + Me.cpFactura
    
    If cpFactura = "pFactura" Then frmVolquetesAdmin.vtablaCtaCte = "pcuentascorrientes"
    If cpFactura = "Factura" Then frmVolquetesAdmin.vtablaCtaCte = "cuentascorrientes"
    
    frmVolquetesAdmin.Show

End If


End Sub

Private Sub PushButton5_Click()
fpintarGrilla
End Sub



Private Sub PushButton6_Click()
Call fbuscarGrilla("articulos", "Descrip", "Codigo", Me.vDarticulo.Name, Me)  ' ema:
End Sub

Private Sub PushButton7_Click()
On Error Resume Next

vwhere = ""
vgrupo = ""
vorden = "fdetalle.codigo"

If Me.rbSinAgruar Then
    vgrupo = ""
End If

If Me.rdArticulo Then
    vgrupo = "group by fdetalle.codigo"
End If

If Me.rdPersona Then
    vgrupo = "group by t.codigo"
End If


'vwhere = " 1=1 "

If chkFechaTodas.Value = 0 Then
    vwhere = vwhere + " and (t.Fecha >= '" & strfechaMySQL(dtpFecha(0).Value) + "' and t.fecha <= '" & strfechaMySQL(dtpFecha(1).Value) + "')"
End If
 
 
If Not Trim(Me.txtEmpleado.Tag) = "" Then vwhere = vwhere + " AND (factura.Repartidor = '" & Trim(Val(txtEmpleado.Tag)) & "')"
 

If Not Me.vCarticulo.Text = "" Then
    vwhere = vwhere + " and fdetalle.codigo = '" + Me.vCarticulo + "'"
End If


If Me.rbOArticulo.Value Then
    vorden = "fdetalle.codigo"
End If

If Me.rbOCantidad Then
    vorden = "Cantidad desc"
End If

If Me.rbOTotales Then
    vorden = "Deuda desc"
End If



If Me.rdPersona Then
    vsql = sqlAgrupa_persona(vwhere)
End If



If Me.rbMeses Then
    vsql = sqlAgrupa_mes(vwhere)
   'Call valoresGraficas(bfactura, Me.grafico, "")
End If

If Me.rbAnos Then
    vsql = sqlAgrupa_anos(vwhere)
End If

If Me.rdArticulo.Value And Me.tab.Item(1).Selected = True Then
    
    If Not Me.vrubro.Tag = "" Then
        vwhere = vwhere + "and idRubros = '" + Me.vrubro.Tag + "'"
    End If

    vsql = sqlArtFactVenta2(vwhere, vgrupo, vorden, cpFactura)
End If

   With bfactura
        .ConnectionString = pathDBMySQL
        .RecordSource = vsql
        .Refresh
    
    
         If Not .Recordset Is Nothing Then Set KlexDocumentos.Recordset = .Recordset
    
    End With
    
    Me.tab.Item(3).Selected = True  ' cambia a la solapa de datos
    
    cmdImprimir(0).Tag = "grupos"

If Err Then Exit Sub
End Sub
Function sqlAgrupa_mes(ByVal vwhere As String) As String


sqlAgrupa_mes = "select codigo, month(fecha) as Mes, year(fecha) as Ano," & _
 "cast(sum(total) as decimal(10,2)) as Deuda, cast(sum(t.`pagoparcial`) as decimal(10,2)) as Pago, cast((sum(t.`pagoparcial`) / sum(total)) as decimal(10,2)) as Por " & _
 "fro" & _
 "m " + cpFactura + " t where 1=1" + vwhere + "  group by year(t.fecha), month(t.fecha) order by year(t.fecha), month(t.fecha)  desc "

End Function
Function sqlAgrupa_anos(ByVal vwhere As String) As String

sqlAgrupa_anos = "select  codigo, year(fecha) as Ano,  '' as Mes,  " & _
 " cast(sum(total) as decimal(10,2)) as Deuda, cast(sum(t.`pagoparcial`) as decimal(10,2)) as Pago, cast((sum(t.`pagoparcial`) / sum(total)) as decimal(10,2)) as Por " & _
 "fro" & _
 "m " + cpFactura + " t where 1=1 " + vwhere + "  group by year(t.fecha) order by year(t.fecha) desc"

End Function


Function sqlAgrupa_persona(ByVal vwhere As String) As String

sqlAgrupa_persona = "select  t.codigo, t.nombre, " & _
 " cast(sum(total) as decimal(10,2)) as Deuda, cast(sum(t.`pagoparcial`) as decimal(10,2)) as Pago, cast((sum(t.`pagoparcial`) / sum(total)) as decimal(10,2)) as Por " & _
 "fro" & _
 "m " + cpFactura + " t where 1=1 " + vwhere + "  group by t.codigo order by " + vorden

End Function
Private Sub PushButton8_Click()
On Error Resume Next
  'Cargar los datos.
Dim i As Long
ReDim Values(1 To bfactura.Recordset.RecordCount, 1 To 4)
   
         bfactura.Recordset.MoveFirst
         For i = 1 To bfactura.Recordset.RecordCount
              Values(i, 1) = Str(bfactura.Recordset("ano")) + "-" + Str(bfactura.Recordset("mes"))
              Values(i, 2) = bfactura.Recordset("deuda")
              Values(i, 3) = bfactura.Recordset("pago")
              'Values(i, 4) = Val(EsNulo(bfactura.Recordset("por")))
              bfactura.Recordset.MoveNext
         Next i
        
         'Dibujar gráfica
        ' grafico.ChartType = VtChChartType2dXY
        ' grafico.Plot.Axis(VtChAxisIdX).AxisTitle.Text = "Meses"
        ' grafico.Plot.Axis(VtChAxisIdY).AxisTitle.Text = "Valores"
        ' grafico.TitleText = "Facturación  /  Cobro "
        ' grafico.RowCount = 4
        ' grafico.ColumnCount = bfactura.Recordset.RecordCount
      '   grafico.ChartData = Values
If Err Then Exit Sub
End Sub



Private Sub PushButton9_Click()


If Not Val(Me.vtotal.Caption) > 0 And Not Me.viene = "ie" Then
    MsgBox "Cuidado!. Usted No ha seleccionado ningún documento a pagar. " + Chr(13) + "Debe marcarlo haciendo doble clic para que quede marcado con una cruz ", vbInformation
    Exit Sub
End If


Select Case vieneCobro

    Case True
        pasarACobros
        Me.WindowState = 1

End Select

If Me.viene = "ie" Then
    pasarAIE
    Me.WindowState = 1

End If


Unload Me
End Sub

Private Sub PusImprimir_Click()
Call imprimirGrilla(Me.gconciliacion, 4)
End Sub

Private Sub PusIrA_Click()
    
    
    
    If Me.CP = "proveedores" Then
        frmCtaCteC.Tag = "Proveedores"
    End If
    
    If Me.CP = "clientes" Then
        frmCtaCteC.Tag = "Clientes"
    End If


frmCtaCteC.txtCliente = Me.txtCliente
frmCtaCteC.txtCliente.Tag = Me.txtCodigoCliente

Call frmCtaCteC.cmdFiltroMovimientos_Click
frmCtaCteC.Show
frmCtaCteC.SetFocus
End Sub

Private Sub ref_Click()
Call fbuscarGrilla("(select refexportpedidos from Factura group by refexportpedidos) c ", "refexportpedidos", "refexportpedidos", vreferenciaPedido.Name, Me) ' ema:
End Sub

Private Sub tab_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
 log2.Visible = False
End Sub

Private Sub txtCliente_Change()
Me.txtCodigoCliente.Text = Me.txtCliente.Tag
End Sub

Public Sub txtClienteImpagos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        BuscarCliente (txtClienteImpagos.Text)
    End If
End Sub
Public Sub txtEmpleado_KeyPress(KeyAscii As Integer)
On Error Resume Next
    
    If KeyAscii = 13 Then
        If Not Trim(txtEmpleado.Text) = "" Then
            
            Dim rsempleados As New ADODB.Recordset, sqlEmpleados As String
            
            sqlEmpleados = "SELECT * FROM Empleados WHERE (Codigo = '" & Trim(txtEmpleado.Text) & "') OR (nombre LIKE '%" & Trim(txtEmpleado.Text) + "%')"
            
            With rsempleados
                Call .Open(sqlEmpleados, ConnDDBB, adOpenStatic, adLockReadOnly)
                
                If .EOF = True Then Exit Sub
                .MoveFirst
                .Filter = ""
                If Not .EOF = True Then
                    txtEmpleado.Text = .Fields("Nombre").Value
                    txtEmpleado.Tag = .Fields("Codigo").Value
                Else
                    txtEmpleado.Text = ""
                    txtEmpleado.Tag = ""
                End If
            
            End With
            
        End If
    End If


If Err Then GrabarLog "txtEmpleado_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub txtCliente_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        If Not Trim(txtCliente.Text) = "" Then
        

            BuscarCliente (txtCliente.Text)
        End If
        txtEmpleado.SetFocus
    End If
    
If Err Then GrabarLog "txtCliente_Keypress", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub txtRepartoImpagos_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        Dim rsReparto As New ADODB.Recordset, sqlReparto As String
        
        sqlReparto = "SELECT * FROM clireparto WHERE (nreparto = '" & Trim(txtRepartoImpagos.Text) & "')"
        
        With rsReparto
            Call .Open(sqlReparto, ConnDDBB, adOpenStatic, adLockReadOnly)
            
            If Not .EOF = True Then
                txtRepartoImpagos.Tag = .Fields("nreparto").Value
                txtRepartoImpagos.Text = .Fields("Descrip").Value
            End If
        
        End With

    End If
    
    sqlReparto = ""
    
    If rsReparto.State = 1 Then
        rsReparto.Close
        Set rsReparto = Nothing
    End If
    
If Err Then GrabarLog "txtRepartoImpagos_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub ImprimirComentariosFacturaHasar()
On Error Resume Next

    Dim rsComentariosFactura As New ADODB.Recordset, sqlComentariosFactura As String, l As Long
    
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
    End If
    
    
If Err Then GrabarLog "ImprimirComentariosFacturaHasar", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub vcodRepartidor_Change()
vidVendedor = codigo2id(Me.vcodRepartidor.Text)
End Sub

Private Sub vDarticulo_Change()
    Me.vCarticulo.Text = Me.vDarticulo.Tag
End Sub

Private Sub vdescEmpresa_Change()
Me.vcodEmpresa.Text = Me.vdescEmpresa.Tag
End Sub

Private Sub vDesRepartidor_Change()
Me.vcodRepartidor.Text = Me.vDesRepartidor.Tag

vidVendedor = codigo2id(Me.vcodRepartidor.Text)

End Sub

Private Sub vImporteSeleccionado_Click()
vImporteSeleccionado.Caption = Format(vImporteSeleccionado.Caption, "###,###,##0.00")
End Sub

Private Sub vtotal_Change()
'vtotal.Caption = Format(vtotal.Caption, "###,###,##0.00")
End Sub


Public Sub formatGrillaDoc(Col As Long, vformato As String)
Dim i As Long

With Me.KlexDocumentos

    
For i = 1 To .Rows - 1

    .TextMatrix(i, Col) = Format(.TextMatrix(i, Col), vformato)

Next


End With
End Sub




Private Sub borrarTodosLosModulos(vnrointerno As Long)

On Error Resume Next

Dim vasientoNumero As Long



    Call BorrarBase("pcuentascorrientes WHERE nrointerno=" + Str(vnrointerno), pathDBMySQL)
    Call BorrarBase("cuentascorrientes WHERE nrointerno=" + Str(vnrointerno), pathDBMySQL)

    Call BorrarFDetalle(vnrointerno, "idPFDetalle")
    Call BorrarFDetalle(vnrointerno, "idFDetalle")
    
    Call BorrarBase("factura WHERE nrointerno=" + Str(vnrointerno), pathDBMySQL)
    Call BorrarBase("pfactura WHERE nrointerno=" + Str(vnrointerno), pathDBMySQL)
    

    Call BorrarBase("ivafacturaventa WHERE nrointerno=" + Str(vnrointerno), pathDBMySQL)
    Call BorrarBase("ivafacturacompra WHERE nrointerno=" + Str(vnrointerno), pathDBMySQL)


    Call BorrarBase("bancosmovimientos WHERE nrointerno=" + Str(vnrointerno), pathDBMySQL)
    
    
    vasientoNumero = traerDatos2("select * from asientos where nrointerno=" + Str(vnrointerno), "numero", pathDBMySQL)
    
    Call BorrarBase("bancosmovimientos WHERE nrointerno=" + Str(vnrointerno), pathDBMySQL)
    
    
    Call BorrarBase("asientos" + " WHERE numero=" + Trim(Str(vasientoNumero)), pathDBMySQL)
    Call BorrarBase("asientosdetalle" + " WHERE Numero=" + Trim(Str(vasientoNumero)), pathDBMySQL)
    
    Call BorrarBase("retencionesmovimientos" + " WHERE nrointerno=" + Str(vnrointerno), pathDBMySQL)
    
    ' verifica si hay cheques que deben ser borrados o hay que cambiar la custodia
    'verificarChequesBorrar (vnrointerno)
    'Call BorrarBase("cheques" + " WHERE nrointerno=" + Str(vnrointerno), pathDBMySQL)
    
    
If Err Then
    Exit Sub
End If

End Sub

Private Sub llenarDetalles2(vremito As Long, vtipo As String)
On Error Resume Next

bdetalle.ConnectionString = pathDBMySQL
bdetalle.RecordSource = "select * from fdetalle where remito = " + Str(vremito)
bdetalle.Refresh

Dim i As Long

i = 0
With bdetalle.Recordset
    .MoveFirst
    Do Until .EOF
        i = i + 1
        Call llenarlinea2(i, .Fields("cantidad"), .Fields("detalle"), .Fields("precio"), Val(EsNulo(.Fields("descuento"))), .Fields("total"))
        
        
        If vtipo = "Nota C" Then
        
            Call actualizar_stock(.Fields("codigo"), .Fields("cantidad"), "+")
        
        Else
        
            Call actualizar_stock(.Fields("codigo"), .Fields("cantidad"), "-")
        
        End If
        .MoveNext
    Loop
End With


If Err Then Exit Sub
End Sub

Private Sub actualizar_stock(vcodigo As String, vCantidad As Double, sumaresta As String)
    ii = ii + 1
    arr_as(ii) = " update articulos set stock = stock  " + Trim$(sumaresta) + " " + Str(Round(vCantidad, 2)) + " where codigo = '" + vcodigo + "'"
End Sub



Private Sub fllenarCabeDocON()
On Error Resume Next
        With drDoc2.Sections("titulos")
                .Controls("ecodigo2").Caption = ""
                '.Controls("enrocomprobante2").Caption = ""
                .Controls("eDocumento").Visible = True
                .Controls("eLetra").Visible = True
                .Controls("eCodigo").Visible = True
                .Controls("ePtoVta").Visible = True
                .Controls("logo").Visible = True
                .Controls("edcodigo").Visible = True
                .Controls("ednro").Visible = True
                .Controls("cuadrado1").Visible = True
                .Controls("cuadrado1").Visible = True
                .Controls("elinea").Visible = True
                .Controls("enro").Visible = True
                .Controls("eiva").Visible = True
                .Controls("elblCAE0").Visible = True
                .Controls("elblCAE").Visible = True
                .Controls("elblCAE2").Visible = True
                .Controls("enroCodigoBarra").Visible = True
                .Controls("elblVtoCAE0").Visible = True
                .Controls("enroCodigoBarra").Visible = True
                .Controls("elblVtoCAE").Visible = True
                .Controls("lineacae").Visible = True
        End With
End Sub

Private Sub fllenarCabeDoc()
On Error Resume Next

With drDoc2.Sections("titulos")
        .Controls("ecodigo2").Caption = ""
        '.Controls("enrocomprobante2").Caption = ""
        .Controls("eDocumento").Visible = False
        .Controls("eLetra").Visible = False
        .Controls("eCodigo").Visible = False
        .Controls("ePtoVta").Visible = False
        .Controls("logo").Visible = False
        .Controls("edcodigo").Visible = False
        .Controls("ednro").Visible = False
        .Controls("cuadrado1").Visible = False
        .Controls("cuadrado1").Visible = False
        .Controls("elinea").Visible = False
        .Controls("enro").Visible = False
        .Controls("eiva").Visible = False
        .Controls("elblCAE0").Visible = False
        .Controls("elblCAE").Visible = False
        .Controls("elblCAE2").Visible = False
        .Controls("enroCodigoBarra").Visible = False
        .Controls("elblVtoCAE0").Visible = False
        .Controls("elblVtoCAE").Visible = False
        .Controls("lineacae").Visible = False
End With


If Err Then Exit Sub
End Sub


Private Sub llenarlinea2(vi As Long, vCantidad As Double, vDetalle As String, vPrecio As Double, vdesc As Double, vtotal As Double)
Dim ve, vd, vp, vdes, vt As String

With drDoc2.Sections("titulos")

ve = "e" + Trim(Str(vi))
vd = "d" + Trim(Str(vi))
vp = "p" + Trim(Str(vi))
vdes = "des" + Trim(Str(vi))
vt = "t" + Trim(Str(vi))


.Controls(ve).Caption = Str(vCantidad)
.Controls(vd).Caption = vDetalle
.Controls(vp).Caption = Format(vPrecio, "###,###,##0.00")
.Controls(vdes).Caption = Format(vdesc, "###,###,##0.00")
.Controls(vt).Caption = Format(vtotal, "###,###,##0.00")

Exit Sub  ' para que no haga un duplicado en la misma hoja

ve = "ee" + Trim(Str(vi))
vd = "dd" + Trim(Str(vi))
vp = "pp" + Trim(Str(vi))
vdes = "ddes" + Trim(Str(vi))
vt = "tt" + Trim(Str(vi))


.Controls(ve).Caption = Str(vCantidad)
.Controls(vd).Caption = vDetalle
.Controls(vp).Caption = Format(vPrecio, "###,###,##0.00")
.Controls(vdes).Caption = Format(vdesc, "###,###,##0.00")
.Controls(vt).Caption = Format(vtotal, "###,###,##0.00")

End With

End Sub



'With drDoc2
'
''----------- titulos -------
'.Sections("titulos").Controls("enroremito").Caption = ""
'.Sections("titulos").Controls("ecventa").Caption = "Cuentas Corrientes"
'
'.Sections("titulos").Controls("enombre").Caption = vf.vnombre
'.Sections("titulos").Controls("edomicilio").Caption = vf.vdireccion
'.Sections("titulos").Controls("elocalidad").Caption = vf.vlocalidad
'.Sections("titulos").Controls("ecuit").Caption = vf.vCuit
'.Sections("titulos").Controls("efecha").Caption = vf.vfecha
''---------------------------
'
'End With


Private Sub llenarDocumentos2(vtipo As String, vncomprobante As Long)
Dim Form As DataReport


With drDoc2

'----------- titulos -------
.Sections("titulos").Controls("enroremito").Caption = ""
.Sections("titulos").Controls("ecventa").Caption = "Cuentas Corrientes"

.Sections("titulos").Controls("enombre").Caption = vf.vnombre
.Sections("titulos").Controls("edomicilio").Caption = vf.vdireccion
.Sections("titulos").Controls("elocalidad").Caption = vf.vlocalidad
.Sections("titulos").Controls("ecuit").Caption = vf.vcuit
.Sections("titulos").Controls("efecha").Caption = vf.vfecha
.Sections("titulos").Controls("eiva").Caption = vf.vIva
'---------------------------

.Sections("titulos").Controls("etotal").Caption = Format(vf.vtotal, "#,###,##0.00")
.Sections("titulos").Controls("esubtotal").Caption = Format(vf.vSubTotal, "#,###,##0.00")

.Sections("titulos").Controls("eiva105").Caption = Format(vf.viva150, "#,###,##0.00")
.Sections("titulos").Controls("eiva21").Caption = Format(vf.vIva210, "#,###,##0.00")
.Sections("titulos").Controls("eiva27").Caption = Format(vf.viva270, "#,###,##0.00")

.Sections("titulos").Controls("edescuento").Caption = ""
'.Sections("Totales").Controls("eimpuesto").Caption = Format(vgTimpuesto, "#,###,##0.00")

'.Sections("Totales").Controls("etotal").Caption = Format(vgTtotal, "#,###,##0.00")

If vtipo = "Documento" Then
    .Sections("titulos").Controls("etitulo").Caption = "Docuemento no válido como factura"
    .Sections("titulos").Controls("encomprobante").Caption = "Nro. Comprobante : " + Str(vncomprobante)
   ' .Sections("titulos").Controls("eetitulo").Caption = "Docuemento no válido como factura"
   ' .Sections("titulos").Controls("eencomprobante").Caption = "Nro. Comprobante : " + Str(vncomprobante)
    
    Call fllenarCabeDoc
    
    .Sections("titulos").Controls("ttiva").Caption = ""
    '.Sections("titulos").Controls("tttiva").Caption = ""
End If


If vtipo = "Exento" Then

Call fllenarCabeDoc

   ' .Sections("titulos").Controls("etitulo").Caption = ""
   ' .Sections("titulos").Controls("ncomprobante").Caption = ""
'    .Sections("titulos").Controls("eetitulo").Caption = ""
'    .Sections("titulos").Controls("encomprobante").Caption = ""
'
    .Sections("titulos").Controls("ttiva").Caption = ""
   ' .Sections("titulos").Controls("tttiva").Caption = ""
End If


If vtipo = "Fact A" Or vtipo = "Fact B" Then


        If .Sections("titulos").Controls("cuadrado1").Visible = False Then ' verifico si están apagado los encabe para los doc a y b
                Call fllenarCabeDocON
        End If


    .Sections("titulos").Controls("etitulo").Caption = ""
    .Sections("titulos").Controls("encomprobante").Caption = ""
    

       
' encabezado factura A
.Sections("titulos").Controls("eDocumento").Caption = veDocumento
.Sections("titulos").Controls("eCodigo").Caption = vTipoComprobante
.Sections("titulos").Controls("eLetra").Caption = UCase(veLetra)
.Sections("titulos").Controls("ePtoVta").Caption = Format(veptovta, "0000")
.Sections("titulos").Controls("eNro").Caption = Format(venro, "00000000")
    
    
  '  .Sections("titulos").Controls("eetitulo").Caption = ""
  '  .Sections("titulos").Controls("encomprobante").Caption = " "
End If


'
'
''----------- titulos -------
'.Sections("titulos").Controls("eenroremito").Caption = ""
'.Sections("titulos").Controls("eecventa").Caption = "Cuentas Corrientes"
'
'.Sections("titulos").Controls("eenombre").Caption = vf.vnombre
'.Sections("titulos").Controls("eedomicilio").Caption = vf.vdireccion
'.Sections("titulos").Controls("eelocalidad").Caption = vf.vlocalidad
'.Sections("titulos").Controls("eecuit").Caption = vf.vCuit
'.Sections("titulos").Controls("eefecha").Caption = vf.vfecha
''---------------------------
'
'.Sections("titulos").Controls("eetotal").Caption = Format(vf.vtotal, "#,###,##0.00")
'.Sections("titulos").Controls("eesubtotal").Caption = Format(vf.vSubTotal, "#,###,##0.00")
'
'.Sections("titulos").Controls("eeiva105").Caption = Format(vf.viva150, "#,###,##0.00")
'.Sections("titulos").Controls("eeiva21").Caption = Format(vf.vIva210, "#,###,##0.00")
'.Sections("titulos").Controls("eeiva27").Caption = Format(vf.viva270, "#,###,##0.00")
'
'.Sections("titulos").Controls("eedescuento").Caption = ""
''.Sections("Totales").Controls("eeimpuesto").Caption = Format(vgTimpuesto, "#,###,##0.00")
'
''.Sections("Totales").Controls("eetotal").Caption = Format(vgTtotal, "#,###,##0.00")
'
'If vTipo = "Documento" Then
'    .Sections("titulos").Controls("eetitulo").Caption = "Docuemento no válido como factura"
'    .Sections("titulos").Controls("encomprobante").Caption = "Nro. Comprobante : "
'Else
'    .Sections("titulos").Controls("eetitulo").Caption = ""
'    .Sections("titulos").Controls("eencomprobante").Caption = ""
'End If
'
'.Sections("titulos").Controls("esaldo").Caption = Format(vf.vsaldo, "#,###,##0.00")
'.Sections("titulos").Controls("eesaldo").Caption = Format(vf.vsaldo, "#,###,##0.00")
'


End With

End Sub


Private Sub llenarDocumentos()
Dim Form As DataReport
Dim vestructura As Long


vestructura = 0


If Not vConfigGral.vempresa = "wgestionPoli" Then

        If vestructura = 0 Then
                With ifacta
                
                '----------- titulos -------
                .Sections("titulos").Controls("enroremito").Caption = ""
                .Sections("titulos").Controls("ecventa").Caption = "Cuentas Corrientes"
                
                .Sections("titulos").Controls("enombre").Caption = vf.vnombre
                .Sections("titulos").Controls("edomicilio").Caption = vf.vdireccion
                .Sections("titulos").Controls("elocalidad").Caption = vf.vlocalidad
                .Sections("titulos").Controls("ecuit").Caption = vf.vcuit
                .Sections("titulos").Controls("efecha").Caption = vf.vfecha
                '---------------------------
                
                .Sections("Totales").Controls("etotal").Caption = Format(vf.vtotal, "#,###,##0.00")
                .Sections("Totales").Controls("esubtotal").Caption = Format(vf.vSubTotal, "#,###,##0.00")
                
                .Sections("Totales").Controls("eiva105").Caption = Format(vf.viva150, "#,###,##0.00")
                .Sections("Totales").Controls("eiva21").Caption = Format(vf.vIva210, "#,###,##0.00")
                .Sections("Totales").Controls("eiva27").Caption = Format(vf.viva270, "#,###,##0.00")
                
                .Sections("Totales").Controls("edescuento").Caption = ""
                '.Sections("Totales").Controls("eimpuesto").Caption = Format(vgTimpuesto, "#,###,##0.00")
                '.Sections("Totales").Controls("etotal").Caption = Format(vgTtotal, "#,###,##0.00")
                
                End With
        End If

Else

With ifactaPoli


'----------- titulos -------
.Sections("titulos").Controls("enroremito").Caption = ""
.Sections("titulos").Controls("ecventa").Caption = "Cuentas Corrientes"

.Sections("titulos").Controls("enombre").Caption = vf.vnombre
.Sections("titulos").Controls("edomicilio").Caption = vf.vdireccion
.Sections("titulos").Controls("elocalidad").Caption = vf.vlocalidad
.Sections("titulos").Controls("ecuit").Caption = vf.vcuit
.Sections("titulos").Controls("efecha").Caption = vf.vfecha
'---------------------------

.Sections("Totales").Controls("etotal").Caption = Format(vf.vtotal, "#,###,##0.00")
.Sections("Totales").Controls("esubtotal").Caption = Format(vf.vSubTotal, "#,###,##0.00")

.Sections("Totales").Controls("eiva105").Caption = Format(vf.viva150, "#,###,##0.00")
.Sections("Totales").Controls("eiva21").Caption = Format(vf.vIva210, "#,###,##0.00")
.Sections("Totales").Controls("eiva27").Caption = Format(vf.viva270, "#,###,##0.00")

.Sections("Totales").Controls("edescuento").Caption = ""
'.Sections("Totales").Controls("eimpuesto").Caption = Format(vgTimpuesto, "#,###,##0.00")

'.Sections("Totales").Controls("etotal").Caption = Format(vgTtotal, "#,###,##0.00")

End With


End If

End Sub


Private Sub mostrar_Doc2(vtipo As String, vncomprobante As Long)
Dim vtocae As Variant

Call llenarDocumentos2(vtipo, vncomprobante)


With drDoc2
vtocae = Trim(Right(Trim(vf.vcaeVto), 2) + "/" + Mid(Trim(vf.vcaeVto), 5, 2) + "/" + Left(Trim(vf.vcaeVto), 4))
.Sections("titulos").Controls("elblCAE").Caption = vf.vcae
.Sections("titulos").Controls("elblCAE2").Caption = vnroCodigoBarra

.Sections("titulos").Controls("enroCodigoBarra").Caption = vnroCodigoBarra
.Sections("titulos").Controls("elblVtoCAE").Caption = vtocae


'.Sections("titulos").Controls("eelblCAE").Caption = vf.vcae
'.Sections("titulos").Controls("eelblVtoCAE").Caption = vf.vcaeVto

'.Sections("titulos").Controls("eiva105").Visible = False

        .Sections("titulos").Controls("eiva105").Visible = True
        .Sections("titulos").Controls("etiva10").Visible = True

    If vf.viva150 = 0 Then
        .Sections("titulos").Controls("eiva105").Visible = False
        .Sections("titulos").Controls("etiva10").Visible = False
        
        '.Sections("titulos").Controls("eeiva105").Visible = False
        '.Sections("titulos").Controls("eetiva10").Visible = False
    End If

            .Sections("titulos").Controls("eiva21").Visible = True
            .Sections("titulos").Controls("etiva21").Visible = True
    If vf.vIva210 = 0 Then
            .Sections("titulos").Controls("eiva21").Visible = False
            .Sections("titulos").Controls("etiva21").Visible = False
            
            '.Sections("titulos").Controls("eeiva21").Visible = False
            '.Sections("titulos").Controls("eetiva21").Visible = False
    End If

        .Sections("titulos").Controls("eiva27").Visible = True
        .Sections("titulos").Controls("etiva27").Visible = True
    If vf.viva270 = 0 Then
        .Sections("titulos").Controls("eiva27").Visible = False
        .Sections("titulos").Controls("etiva27").Visible = False
        
        '.Sections("titulos").Controls("eetiva27").Visible = False
        '.Sections("titulos").Controls("eeiva27").Visible = False
    End If
    

.Sections("titulos").Controls("esaldo").Caption = Format(vf.vsaldo, "###,###,##0.00")
'.Sections("titulos").Controls("eesaldo").Visible = Format(vf.vsaldo, "###,###,##0.00")



'Set .DataSource = Nothing
'.DataMember = ""
'.Show
Call .PrintReport(False, rptRangeAllPages)

LimpiardrDoc2

'Unload .object

End With


End Sub


Private Sub limpiarDrDoc2Cuerpo()

Dim ve, vd, vp, vdes, vt As String

Dim vi As Long



With drDoc2.Sections("titulos")
        
        For vi = 1 To 18
                        ve = "e" + Trim(Str(vi))
                        vd = "d" + Trim(Str(vi))
                        vp = "p" + Trim(Str(vi))
                        vdes = "des" + Trim(Str(vi))
                        vt = "t" + Trim(Str(vi))
                        
                        
                        .Controls(ve).Caption = ""
                        .Controls(vd).Caption = ""
                        .Controls(vp).Caption = ""
                        .Controls(vdes).Caption = ""
                        .Controls(vt).Caption = ""
                        
'                        ve = "ee" + Trim(Str(vi))
'                        vd = "dd" + Trim(Str(vi))
'                        vp = "pp" + Trim(Str(vi))
'                        vdes = "ddes" + Trim(Str(vi))
'                        vt = "tt" + Trim(Str(vi))
'
'
'                        .Controls(ve).Caption = ""
'                        .Controls(vd).Caption = ""
'                        .Controls(vp).Caption = ""
'                        .Controls(vdes).Caption = ""
'                        .Controls(vt).Caption = ""
        Next

End With

End Sub



Private Sub LimpiardrDoc2()


limpiarDrDoc2Cuerpo


With drDoc2

'----------- titulos -------
.Sections("titulos").Controls("enroremito").Caption = ""
.Sections("titulos").Controls("ecventa").Caption = ""
.Sections("titulos").Controls("enombre").Caption = ""
.Sections("titulos").Controls("edomicilio").Caption = ""
.Sections("titulos").Controls("elocalidad").Caption = ""
.Sections("titulos").Controls("ecuit").Caption = ""
.Sections("titulos").Controls("efecha").Caption = ""
'---------------------------

.Sections("titulos").Controls("etotal").Caption = ""
.Sections("titulos").Controls("esubtotal").Caption = ""

.Sections("titulos").Controls("eiva105").Caption = ""
.Sections("titulos").Controls("eiva21").Caption = ""
.Sections("titulos").Controls("eiva27").Caption = ""

.Sections("titulos").Controls("edescuento").Caption = ""
'.Sections("Totales").Controls("eimpuesto").Caption = Format(vgTimpuesto, "#,###,##0.00")

'.Sections("Totales").Controls("etotal").Caption = Format(vgTtotal, "#,###,##0.00")

   ' .Sections("titulos").Controls("ttiva").Caption = ""
   ' .Sections("titulos").Controls("tttiva").Caption = ""
    .Sections("titulos").Controls("etitulo").Caption = ""
    '.Sections("titulos").Controls("ncomprobante").Caption = ""
   ' .Sections("titulos").Controls("eetitulo").Caption = ""
    .Sections("titulos").Controls("encomprobante").Caption = ""
   ' .Sections("titulos").Controls("ttiva").Caption = ""
   ' .Sections("titulos").Controls("tttiva").Caption = ""
    .Sections("titulos").Controls("etitulo").Caption = ""
    .Sections("titulos").Controls("encomprobante").Caption = ""
    '.Sections("titulos").Controls("eetitulo").Caption = ""
    .Sections("titulos").Controls("encomprobante").Caption = ""
     
     '----------- titulos -------
    
   ' .Sections("titulos").Controls("eenroremito").Caption = ""
   ' .Sections("titulos").Controls("eecventa").Caption = ""

   ' .Sections("titulos").Controls("eenombre").Caption = ""
   ' .Sections("titulos").Controls("eedomicilio").Caption = ""
   ' .Sections("titulos").Controls("eelocalidad").Caption = ""
   ' .Sections("titulos").Controls("eecuit").Caption = ""
   ' .Sections("titulos").Controls("eefecha").Caption = ""
    
    '---------------------------

   ' .Sections("titulos").Controls("eetotal").Caption = ""
  '  .Sections("titulos").Controls("eesubtotal").Caption = ""

  '  .Sections("titulos").Controls("eeiva105").Caption = ""
  '  .Sections("titulos").Controls("eeiva21").Caption = ""
  '  .Sections("titulos").Controls("eeiva27").Caption = ""

  '  .Sections("titulos").Controls("eedescuento").Caption = ""
   ' .Sections("titulos").Controls("eetitulo").Caption = ""
   ' .Sections("titulos").Controls("encomprobante").Caption = ""
  '  .Sections("titulos").Controls("eencomprobante").Caption = ""
   ' .Sections("titulos").Controls("esaldo").Caption = ""
  '  .Sections("titulos").Controls("eesaldo").Caption = ""


    .Sections("titulos").Controls("eiva105").Caption = ""
  '  .Sections("titulos").Controls("etiva10").Caption = ""
  '  .Sections("titulos").Controls("eiva21").Caption = ""
  '  .Sections("titulos").Controls("etiva21").Caption = ""
  '  .Sections("titulos").Controls("eeiva21").Caption = ""
 '   .Sections("titulos").Controls("eetiva21").Caption = ""
    .Sections("titulos").Controls("eiva27").Caption = ""
  '  .Sections("titulos").Controls("etiva27").Caption = ""
    '.Sections("titulos").Controls("eetiva27").Caption = ""
    '.Sections("titulos").Controls("eeiva27").Caption = ""
    .Sections("titulos").Controls("esaldo").Caption = ""
   ' .Sections("titulos").Controls("eesaldo").Caption = ""

End With

End Sub


Private Sub mostrar_ifactaLabel()

llenarDocumentos
 
If vConfigGral.vempresa = "wgestionPoli" Then
    With ifactaPoli
    If vf.viva150 = 0 Then
        .Sections("Totales").Controls("eiva105").Visible = False
        .Sections("Totales").Controls("etiva10").Visible = False
    End If


If vf.vIva210 = 0 Then
.Sections("Totales").Controls("eiva21").Visible = False
.Sections("Totales").Controls("etiva21").Visible = False
End If


If vf.viva270 = 0 Then
.Sections("Totales").Controls("eiva27").Visible = False
.Sections("Totales").Controls("etiva27").Visible = False
End If



Call .PrintReport(False, rptRangeAllPages)
Unload .object
End With

End If




If Not vConfigGral.vempresa = "wgestionPoli" Then

With ifacta
    
    If vf.viva150 = 0 Then
        .Sections("Totales").Controls("eiva105").Visible = False
        .Sections("Totales").Controls("etiva10").Visible = False
    End If


If vf.vIva210 = 0 Then
    .Sections("Totales").Controls("eiva21").Visible = False
    .Sections("Totales").Controls("etiva21").Visible = False
End If


If vf.viva270 = 0 Then
    .Sections("Totales").Controls("eiva27").Visible = False
    .Sections("Totales").Controls("etiva27").Visible = False
End If



Call .PrintReport(False, rptRangeAllPages)
'.Show
Unload .object
End With

End If

End Sub


Private Sub mostrar_ifactb()

llenarDocumentos

If vConfigGral.vempresa = "wgestionPoli" Then


            With ifactaPoli
            
            '---------- acomoda datos
            .Sections("Totales").Controls("eiva105").Caption = ""
            .Sections("Totales").Controls("eiva21").Caption = ""
            .Sections("Totales").Controls("eiva27").Caption = ""
            .Sections("Totales").Controls("ttiva").Caption = ""
            
           ' .Sections("Totales").Controls("etiva10").Caption = ""
          '  .Sections("Totales").Controls("etiva21").Caption = ""
           ' .Sections("Totales").Controls("etiva27").Caption = ""
          '
           ' .Sections("titulos").Controls("ettiva").Caption = "IVA exento"
            '--------------------------
             
            Call .PrintReport(False, rptRangeAllPages)
            Unload ifactaPoli
            
            
            End With


Else

            With ifacta
            
            '---------- acomoda datos
            .Sections("Totales").Controls("eiva105").Caption = ""
            .Sections("Totales").Controls("eiva21").Caption = ""
            .Sections("Totales").Controls("eiva27").Caption = ""
            .Sections("Totales").Controls("ttiva").Caption = ""
            
          '  .Sections("Totales").Controls("etiva10").Caption = ""
          '  .Sections("Totales").Controls("etiva21").Caption = ""
          '  .Sections("Totales").Controls("etiva27").Caption = ""
          
           ' .Sections("titulos").Controls("ettiva").Caption = "IVA exento"
            '--------------------------
             
            Call .PrintReport(False, rptRangeAllPages)
            Unload ifacta
            
            
            End With


End If

End Sub

Private Sub mostrar_remito()

'llearDocumentos

With drRemito

'----------- titulos -------
'.Sections("titulos").Controls("enroremito").Caption = Me.vnroremito2
'.Sections("titulos").Controls("ecventa").Caption = Me.vcventa

'.Sections("titulos").Controls("enombre").Caption = txtCliente(0).Text
'.Sections("titulos").Controls("edomicilio").Caption = txtCliente(1).Text
'.Sections("titulos").Controls("elocalidad").Caption = txtCliente(2).Text
'.Sections("titulos").Controls("ecuit").Caption = txtCliente(3).Text
'.Sections("titulos").Controls("efecha").Caption = Str(dtpFecha)
''---------------------------
'
'
''---------- acomoda datos
'.Sections("Totales").Controls("erecibio").Caption = Me.vRemitoRecibio
'.Sections("Totales").Controls("eTransportistaNombre").Caption = Me.vTransportistaNombre
'.Sections("Totales").Controls("eTransportistaCuit").Caption = Me.vTransportistaCuit
'.Sections("Totales").Controls("eTransportistaDomicilio").Caption = Me.vTransportistaDomicilio
'.Sections("Totales").Controls("elentrega").Caption = Me.vlentrega
'.Sections("Totales").Controls("eobservaciones").Caption = vobservacion
'
'.Show
'--------------------------
End With


'Me.vRemitoRecibio = ""
'Me.vTransportistaNombre = ""
'Me.vTransportistaCuit = ""
'Me.vTransportistaDomicilio = ""
'Me.vlentrega = ""
'Me.txtObservaciones = ""


End Sub



Private Sub mostrar_documentos(vnrocomprobante As Long)

With idocumento
        '----------- titulos -------
        .Sections("titulos").Controls("enroremito").Caption = Str(vnrocomprobante)
        '.Sections("titulos").Controls("ecventa").Caption = Me.vcventa
        
        .Sections("titulos").Controls("enombre").Caption = vf.vnombre
        .Sections("titulos").Controls("edomicilio").Caption = vf.vdireccion
        .Sections("titulos").Controls("elocalidad").Caption = vf.vlocalidad
        .Sections("titulos").Controls("ecuit").Caption = vf.vcuit
        .Sections("titulos").Controls("efecha").Caption = vf.vfecha
        
        '---------------------------
        
        .Sections("totales").Controls("etotal").Caption = Format(vf.vtotal, "#,###,##0.00")
        .Sections("totales").Controls("esubtotal").Caption = Format(vf.vSubTotal, "#,###,##0.00")
        .Sections("Totales").Controls("edescuento").Caption = ""
        
        Call .PrintReport(False, rptRangeAllPages)
        Unload idocumento
End With

End Sub

