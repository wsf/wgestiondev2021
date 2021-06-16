VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "Copia de KlexGrid.ocx"
Object = "{9746E3DA-06E1-4D26-9CE4-D9F6411A9C70}#1.0#0"; "SMGA_OcxTxt2008.ocx"
Begin VB.Form frmCompras 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Documentos de Compras"
   ClientHeight    =   8730
   ClientLeft      =   4980
   ClientTop       =   -12840
   ClientWidth     =   13635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   13635
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.GroupBox GroupBox5 
      Height          =   135
      Left            =   0
      TabIndex        =   187
      Top             =   420
      Width           =   13485
      _Version        =   851968
      _ExtentX        =   23786
      _ExtentY        =   238
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
   End
   Begin XtremeSuiteControls.GroupBox GBOtrosDocumentos 
      Height          =   3075
      Left            =   15240
      TabIndex        =   40
      Top             =   6420
      Visible         =   0   'False
      Width           =   10425
      _Version        =   851968
      _ExtentX        =   18389
      _ExtentY        =   5424
      _StockProps     =   79
      Caption         =   " Otros Documentos"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.FlatEdit txtIB 
         Height          =   315
         Index           =   5
         Left            =   6480
         TabIndex        =   41
         Top             =   630
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtTipoMovimiento 
         Height          =   315
         Index           =   0
         Left            =   1650
         TabIndex        =   42
         Top             =   270
         Width           =   735
         _Version        =   851968
         _ExtentX        =   1296
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.ComboBox cboBienesServicios 
         Height          =   315
         Left            =   6480
         TabIndex        =   43
         Top             =   1350
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         DropDownItemCount=   2
      End
      Begin XtremeSuiteControls.FlatEdit txtIB 
         Height          =   315
         Index           =   0
         Left            =   1650
         TabIndex        =   45
         Top             =   630
         Width           =   2655
         _Version        =   851968
         _ExtentX        =   4683
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtIB 
         Height          =   315
         Index           =   1
         Left            =   1650
         TabIndex        =   46
         Top             =   990
         Width           =   735
         _Version        =   851968
         _ExtentX        =   1296
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtIB 
         Height          =   315
         Index           =   2
         Left            =   2490
         TabIndex        =   47
         Top             =   990
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtIB 
         Height          =   315
         Index           =   3
         Left            =   1650
         TabIndex        =   48
         Top             =   1350
         Width           =   2655
         _Version        =   851968
         _ExtentX        =   4683
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtIB 
         Height          =   315
         Index           =   4
         Left            =   1650
         TabIndex        =   49
         Top             =   1710
         Width           =   2655
         _Version        =   851968
         _ExtentX        =   4683
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtIB 
         Height          =   315
         Index           =   6
         Left            =   6480
         TabIndex        =   50
         Top             =   990
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtIB 
         Height          =   315
         Index           =   10
         Left            =   6480
         TabIndex        =   51
         Top             =   1710
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
         MaxLength       =   254
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   1
         Left            =   2415
         TabIndex        =   52
         Tag             =   "TipoMovimientos"
         Top             =   270
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtTipoMovimiento 
         Height          =   315
         Index           =   1
         Left            =   2775
         TabIndex        =   53
         Top             =   270
         Width           =   1530
         _Version        =   851968
         _ExtentX        =   2708
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtNroInterno 
         Height          =   285
         Left            =   6480
         TabIndex        =   66
         Top             =   300
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtIB 
         Height          =   435
         Index           =   7
         Left            =   1650
         TabIndex        =   44
         Top             =   2520
         Width           =   7785
         _Version        =   851968
         _ExtentX        =   13732
         _ExtentY        =   767
         _StockProps     =   77
         BackColor       =   -2147483643
         ScrollBars      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtBancoCheque 
         Height          =   315
         Index           =   0
         Left            =   1620
         TabIndex        =   130
         Top             =   2100
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
         Left            =   2520
         TabIndex        =   131
         Tag             =   "compra-caja"
         Top             =   2100
         Visible         =   0   'False
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
         Left            =   2910
         TabIndex        =   132
         Top             =   2100
         Visible         =   0   'False
         Width           =   6495
         _Version        =   851968
         _ExtentX        =   11456
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Caja:"
         Height          =   195
         Index           =   3
         Left            =   1050
         TabIndex        =   133
         Top             =   2160
         Visible         =   0   'False
         Width           =   435
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   15
         Left            =   4710
         TabIndex        =   65
         Top             =   360
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3087
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Nro Interno :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   0
         Left            =   -270
         TabIndex        =   64
         Top             =   315
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3087
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Tipo Comprobantes:"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   1
         Left            =   -270
         TabIndex        =   63
         Top             =   675
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3087
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Importe Gravado :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   2
         Left            =   -270
         TabIndex        =   62
         Top             =   1035
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3087
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "IVA :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   3
         Left            =   -270
         TabIndex        =   61
         Top             =   1395
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3087
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Deducciones :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   4
         Left            =   -270
         TabIndex        =   60
         Top             =   1755
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3087
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Retenciones :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   10
         Left            =   4680
         TabIndex        =   59
         Top             =   1785
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3096
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Total :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   5
         Left            =   4680
         TabIndex        =   58
         Top             =   675
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3087
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Importe NO Gravado :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   6
         Left            =   4680
         TabIndex        =   57
         Top             =   1035
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3087
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Impuesto Exento :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   7
         Left            =   4680
         TabIndex        =   56
         Top             =   1395
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3087
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Bienes/Servicios :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   8
         Left            =   -240
         TabIndex        =   55
         Top             =   2595
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3087
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Leyenda :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblTipoDocumento 
         Height          =   255
         Left            =   7440
         TabIndex        =   54
         Top             =   120
         Visible         =   0   'False
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         ForeColor       =   -2147483635
         BackColor       =   65535
         Alignment       =   2
      End
   End
   Begin MSDataGridLib.DataGrid dgProveedores 
      Height          =   2535
      Left            =   1440
      TabIndex        =   68
      Top             =   1125
      Visible         =   0   'False
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4471
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   255
      HeadLines       =   1
      RowHeight       =   15
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
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   525
      Left            =   0
      TabIndex        =   135
      Top             =   -90
      Width           =   13515
      _Version        =   851968
      _ExtentX        =   23839
      _ExtentY        =   926
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton cmdAcciones 
         Height          =   375
         Index           =   0
         Left            =   4350
         TabIndex        =   136
         Top             =   120
         Width           =   1395
         _Version        =   851968
         _ExtentX        =   2461
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Borrar Detalle"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCompras.frx":0000
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton cmdAcciones 
         Height          =   375
         Index           =   1
         Left            =   5760
         TabIndex        =   137
         Top             =   120
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Vaciar Grilla"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCompras.frx":059A
         ImageAlignment  =   8
      End
      Begin XtremeSuiteControls.PushButton cmdAcciones 
         Height          =   375
         Index           =   3
         Left            =   30
         TabIndex        =   138
         Top             =   120
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Grabar <F2>"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCompras.frx":0B34
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton cmdAcciones 
         Height          =   375
         Index           =   2
         Left            =   7110
         TabIndex        =   139
         Top             =   120
         Width           =   1185
         _Version        =   851968
         _ExtentX        =   2090
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Nuevo"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCompras.frx":10CE
         ImageAlignment  =   4
         TextImageRelation=   0
      End
      Begin XtremeSuiteControls.PushButton cmdAcciones 
         Height          =   375
         Index           =   4
         Left            =   1410
         TabIndex        =   140
         Top             =   120
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCompras.frx":1668
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton cmdAcciones 
         Height          =   375
         Index           =   5
         Left            =   12210
         TabIndex        =   141
         Top             =   120
         Visible         =   0   'False
         Width           =   1275
         _Version        =   851968
         _ExtentX        =   2249
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Salir"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCompras.frx":1C02
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton cmdAcciones 
         Height          =   375
         Index           =   6
         Left            =   9180
         TabIndex        =   190
         Top             =   120
         Width           =   2625
         _Version        =   851968
         _ExtentX        =   4630
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Grabar como otro Documento"
         UseVisualStyle  =   -1  'True
         ImageAlignment  =   4
      End
   End
   Begin MSDataGridLib.DataGrid dgArticulos 
      Height          =   2655
      Left            =   75
      TabIndex        =   33
      Top             =   3690
      Visible         =   0   'False
      Width           =   13380
      _ExtentX        =   23601
      _ExtentY        =   4683
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   255
      HeadLines       =   1
      RowHeight       =   15
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
   Begin XtremeSuiteControls.GroupBox GBCaja 
      Height          =   3135
      Left            =   15120
      TabIndex        =   105
      Top             =   7170
      Visible         =   0   'False
      Width           =   4815
      _Version        =   851968
      _ExtentX        =   8493
      _ExtentY        =   5530
      _StockProps     =   79
      Caption         =   "Ingreso de Pago de Contado:"
      ForeColor       =   14737632
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   30
         ScaleHeight     =   465
         ScaleWidth      =   4785
         TabIndex        =   124
         TabStop         =   0   'False
         Top             =   2610
         Width           =   4785
         Begin XtremeSuiteControls.PushButton cmdCerrarPago 
            Height          =   375
            Left            =   3580
            TabIndex        =   125
            Top             =   90
            Width           =   1155
            _Version        =   851968
            _ExtentX        =   2028
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Cerrar"
            Appearance      =   6
         End
         Begin XtremeSuiteControls.PushButton cmdGuardarPago 
            Height          =   375
            Left            =   2450
            TabIndex        =   126
            Top             =   90
            Width           =   1155
            _Version        =   851968
            _ExtentX        =   2028
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Guardar"
            Appearance      =   6
            BorderGap       =   10
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
            ForeColor       =   &H00E0E0E0&
            Height          =   240
            Index           =   2
            Left            =   50
            TabIndex        =   128
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
            Index           =   3
            Left            =   75
            TabIndex        =   127
            Top             =   150
            Width           =   1770
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtCaja 
         Height          =   315
         Index           =   0
         Left            =   1440
         TabIndex        =   106
         Top             =   240
         Width           =   495
         _Version        =   851968
         _ExtentX        =   873
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   2
         Left            =   1995
         TabIndex        =   107
         Tag             =   "CajaBanco"
         Top             =   240
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtCaja 
         Height          =   315
         Index           =   1
         Left            =   2400
         TabIndex        =   108
         Top             =   240
         Width           =   2295
         _Version        =   851968
         _ExtentX        =   4048
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtCaja 
         Height          =   315
         Index           =   2
         Left            =   1440
         TabIndex        =   109
         Top             =   600
         Width           =   495
         _Version        =   851968
         _ExtentX        =   873
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   3
         Left            =   1995
         TabIndex        =   110
         Tag             =   "BancoCuenta"
         Top             =   600
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtCaja 
         Height          =   315
         Index           =   3
         Left            =   2400
         TabIndex        =   111
         Top             =   600
         Width           =   2295
         _Version        =   851968
         _ExtentX        =   4048
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtCaja 
         Height          =   435
         Index           =   8
         Left            =   1440
         TabIndex        =   113
         Top             =   2040
         Width           =   3255
         _Version        =   851968
         _ExtentX        =   5741
         _ExtentY        =   767
         _StockProps     =   77
         BackColor       =   -2147483643
         MaxLength       =   255
         ScrollBars      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtCaja 
         Height          =   315
         Index           =   7
         Left            =   1440
         TabIndex        =   114
         Top             =   1680
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtCaja 
         Height          =   315
         Index           =   4
         Left            =   1440
         TabIndex        =   115
         Top             =   960
         Width           =   495
         _Version        =   851968
         _ExtentX        =   873
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   4
         Left            =   1995
         TabIndex        =   116
         Tag             =   "TipoValor"
         Top             =   960
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtCaja 
         Height          =   315
         Index           =   5
         Left            =   2400
         TabIndex        =   117
         Top             =   960
         Width           =   2295
         _Version        =   851968
         _ExtentX        =   4048
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtCaja 
         Height          =   315
         Index           =   6
         Left            =   1440
         TabIndex        =   112
         Top             =   1320
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   9
         Left            =   -90
         TabIndex        =   123
         Top             =   315
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Caja/Banco:"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   11
         Left            =   30
         TabIndex        =   122
         Top             =   690
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Cuenta Banco:"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   12
         Left            =   -90
         TabIndex        =   121
         Top             =   1035
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Tipo de Valor:"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   13
         Left            =   -90
         TabIndex        =   120
         Top             =   1755
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Importe:"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   14
         Left            =   -90
         TabIndex        =   119
         Top             =   2070
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Observaciones:"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   16
         Left            =   -90
         TabIndex        =   118
         Top             =   1395
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Numero de Valor:"
         Alignment       =   1
         Transparent     =   -1  'True
      End
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "ver ..."
      Height          =   255
      Index           =   1
      Left            =   16440
      TabIndex        =   101
      Top             =   4860
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "ver ..."
      Height          =   255
      Index           =   0
      Left            =   11940
      TabIndex        =   100
      Top             =   5640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   -30
      TabIndex        =   93
      Top             =   2580
      Width           =   13485
      Begin VB.OptionButton opTipoDocumento 
         Caption         =   "Nota de Crédito"
         Height          =   345
         Index           =   8
         Left            =   8610
         Style           =   1  'Graphical
         TabIndex        =   129
         Top             =   120
         Width           =   1440
      End
      Begin VB.OptionButton opTipoDocumento 
         Caption         =   "Comprob."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   7
         Left            =   5730
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   120
         Width           =   1425
      End
      Begin VB.OptionButton opTipoDocumento 
         Caption         =   "Presupuesto"
         Height          =   345
         Index           =   5
         Left            =   2850
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   120
         Width           =   1425
      End
      Begin VB.OptionButton opTipoDocumento 
         Caption         =   "Nota de Débito"
         Height          =   345
         Index           =   3
         Left            =   7170
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   120
         Width           =   1425
      End
      Begin VB.OptionButton opTipoDocumento 
         Caption         =   "Factura"
         Height          =   345
         Index           =   0
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   120
         Value           =   -1  'True
         Width           =   1425
      End
      Begin VB.OptionButton opTipoDocumento 
         Caption         =   "Remito / O.Comp"
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   2
         Left            =   1410
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   120
         Width           =   1425
      End
      Begin VB.OptionButton opTipoDocumento 
         Caption         =   "Documento"
         Height          =   345
         Index           =   6
         Left            =   4290
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   120
         Width           =   1425
      End
   End
   Begin VB.OptionButton OptExento 
      Caption         =   "Exento"
      Height          =   215
      Index           =   8
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   3900
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.OptionButton opTipoDocumento 
      Caption         =   "Exento"
      Height          =   215
      Index           =   1
      Left            =   9150
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   6900
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.OptionButton opTipoDocumento 
      Caption         =   "Nota de Crédito"
      Height          =   195
      Index           =   4
      Left            =   5340
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   6900
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Frame Frame1 
      Height          =   1845
      Left            =   0
      TabIndex        =   69
      Top             =   690
      Width           =   9585
      Begin VB.Frame FraMenuCliente 
         BorderStyle     =   0  'None
         Height          =   1665
         Left            =   9060
         TabIndex        =   76
         Top             =   120
         Width           =   435
         Begin VB.CommandButton cmdMenuB 
            Height          =   315
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   79
            ToolTipText     =   "Buscar documento"
            Top             =   780
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton cmdMenuN 
            Height          =   315
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   78
            ToolTipText     =   "Nuevo"
            Top             =   450
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.CommandButton cmdMenuG 
            Height          =   315
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   77
            ToolTipText     =   "Grabar un Proveedor nuevo"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   315
         End
      End
      Begin VB.TextBox txtProveedor 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   75
         Top             =   150
         Width           =   7245
      End
      Begin VB.TextBox txtProveedor 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   3
         Left            =   1410
         TabIndex        =   74
         Top             =   1140
         Width           =   1575
      End
      Begin VB.TextBox txtProveedor 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   5
         Left            =   3990
         TabIndex        =   73
         Top             =   1170
         Width           =   5025
      End
      Begin VB.TextBox txtProveedor 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   4
         Left            =   1410
         TabIndex        =   72
         Top             =   1500
         Width           =   7605
      End
      Begin VB.TextBox txtProveedor 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   1
         Left            =   1410
         TabIndex        =   71
         Top             =   480
         Width           =   7605
      End
      Begin VB.TextBox txtProveedor 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00404040&
         Height          =   285
         Index           =   2
         Left            =   1410
         TabIndex        =   70
         Top             =   810
         Width           =   7605
      End
      Begin VB.Label Label3 
         Caption         =   "<F1>"
         Height          =   285
         Left            =   8700
         TabIndex        =   186
         Top             =   150
         Width           =   375
      End
      Begin VB.Label lblProveedor 
         Alignment       =   1  'Right Justify
         Caption         =   "> Tipo de I.V.A. :"
         Height          =   165
         Index           =   4
         Left            =   90
         TabIndex        =   85
         Top             =   1560
         Width           =   1275
      End
      Begin VB.Label lblProveedor 
         Alignment       =   1  'Right Justify
         Caption         =   "> Dirección:"
         Height          =   165
         Index           =   1
         Left            =   120
         TabIndex        =   84
         Top             =   510
         Width           =   1275
      End
      Begin VB.Label lblProveedor 
         Alignment       =   1  'Right Justify
         Caption         =   "> Teléfono :"
         Height          =   165
         Index           =   3
         Left            =   90
         TabIndex        =   83
         Top             =   1230
         Width           =   1275
      End
      Begin VB.Label lblProveedor 
         Alignment       =   1  'Right Justify
         Caption         =   "> Localidad:"
         Height          =   165
         Index           =   2
         Left            =   120
         TabIndex        =   82
         Top             =   855
         Width           =   1275
      End
      Begin VB.Label lblProveedor 
         AutoSize        =   -1  'True
         Caption         =   "> C.U.I.T :"
         Height          =   195
         Index           =   5
         Left            =   3120
         TabIndex        =   81
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblProveedor 
         Alignment       =   1  'Right Justify
         Caption         =   "> Proveedor:"
         Height          =   165
         Index           =   0
         Left            =   120
         TabIndex        =   80
         Top             =   210
         Width           =   1275
      End
   End
   Begin XtremeSuiteControls.TabControl TabTotales 
      Height          =   1845
      Left            =   60
      TabIndex        =   19
      Top             =   6870
      Width           =   11295
      _Version        =   851968
      _ExtentX        =   19923
      _ExtentY        =   3254
      _StockProps     =   68
      PaintManager.DisableLunaColors=   0   'False
      PaintManager.ShowIcons=   -1  'True
      ItemCount       =   4
      Item(0).Caption =   "Totales"
      Item(0).ControlCount=   26
      Item(0).Control(0)=   "lblTotales(0)"
      Item(0).Control(1)=   "chkIva(2)"
      Item(0).Control(2)=   "chkIva(1)"
      Item(0).Control(3)=   "chkIva(0)"
      Item(0).Control(4)=   "txtIva(0)"
      Item(0).Control(5)=   "txtIva(1)"
      Item(0).Control(6)=   "txtSubTotal"
      Item(0).Control(7)=   "txtIva(2)"
      Item(0).Control(8)=   "lblTotales(1)"
      Item(0).Control(9)=   "lblTotales(2)"
      Item(0).Control(10)=   "lblTotales(3)"
      Item(0).Control(11)=   "txtPorcentajeDescuento"
      Item(0).Control(12)=   "txtDescuento"
      Item(0).Control(13)=   "lblTotales(8)"
      Item(0).Control(14)=   "lblTotales(7)"
      Item(0).Control(15)=   "lblTotales(10)"
      Item(0).Control(16)=   "vNoGravado"
      Item(0).Control(17)=   "vExento"
      Item(0).Control(18)=   "lblTotales(15)"
      Item(0).Control(19)=   "lblTotales(16)"
      Item(0).Control(20)=   "FlatEdit6"
      Item(0).Control(21)=   "lblTotales(17)"
      Item(0).Control(22)=   "lblTotales(18)"
      Item(0).Control(23)=   "vflete"
      Item(0).Control(24)=   "vintereses"
      Item(0).Control(25)=   "vNetoGravado"
      Item(1).Caption =   "Impuestos"
      Item(1).ControlCount=   5
      Item(1).Control(0)=   "lblTotales(6)"
      Item(1).Control(1)=   "lblTotales(9)"
      Item(1).Control(2)=   "txtAuxiliares(1)"
      Item(1).Control(3)=   "txtAuxiliares(3)"
      Item(1).Control(4)=   "GroupBox3"
      Item(2).Caption =   "Comentarios"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "txtObservacion"
      Item(3).Caption =   "Órdenes de compras y trabajo"
      Item(3).ControlCount=   2
      Item(3).Control(0)=   "gridOrdenes"
      Item(3).Control(1)=   "PushButton2"
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   375
         Left            =   -69910
         TabIndex        =   183
         Top             =   450
         Visible         =   0   'False
         Width           =   2595
         _Version        =   851968
         _ExtentX        =   4577
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Seleccionar órdes de compras"
         Appearance      =   3
         Picture         =   "frmCompras.frx":219C
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gridOrdenes 
         Height          =   885
         Left            =   -69910
         TabIndex        =   182
         Top             =   870
         Visible         =   0   'False
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   1561
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.TextBox txtObservacion 
         Height          =   1005
         Left            =   -69880
         TabIndex        =   134
         Top             =   630
         Visible         =   0   'False
         Width           =   11025
      End
      Begin VB.CheckBox chkIva 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   9390
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   23
         Tag             =   "10.5"
         Top             =   675
         Width           =   250
      End
      Begin VB.CheckBox chkIva 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   9390
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   22
         Tag             =   "21"
         Top             =   975
         Width           =   250
      End
      Begin VB.CheckBox chkIva 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   9390
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   21
         Tag             =   "27"
         Top             =   1275
         Width           =   250
      End
      Begin XtremeSuiteControls.FlatEdit txtIva 
         Height          =   285
         Index           =   0
         Left            =   9720
         TabIndex        =   25
         Top             =   675
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2566
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtSubTotal 
         Height          =   285
         Left            =   2250
         TabIndex        =   24
         Top             =   480
         Width           =   1665
         _Version        =   851968
         _ExtentX        =   2937
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtIva 
         Height          =   285
         Index           =   1
         Left            =   9720
         TabIndex        =   26
         Top             =   960
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2566
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtIva 
         Height          =   285
         Index           =   2
         Left            =   9720
         TabIndex        =   27
         Top             =   1275
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2566
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtAuxiliares 
         Height          =   285
         Index           =   1
         Left            =   -61060
         TabIndex        =   31
         Top             =   630
         Visible         =   0   'False
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2558
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtAuxiliares 
         Height          =   285
         Index           =   3
         Left            =   -61090
         TabIndex        =   34
         Top             =   1305
         Visible         =   0   'False
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2558
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtPorcentajeDescuento 
         Height          =   285
         Left            =   2250
         TabIndex        =   142
         Top             =   1470
         Width           =   630
         _Version        =   851968
         _ExtentX        =   1111
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtDescuento 
         Height          =   285
         Left            =   2910
         TabIndex        =   143
         Top             =   1470
         Width           =   990
         _Version        =   851968
         _ExtentX        =   1746
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit vNoGravado 
         Height          =   285
         Left            =   2250
         TabIndex        =   145
         Top             =   780
         Width           =   1665
         _Version        =   851968
         _ExtentX        =   2937
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit vExento 
         Height          =   285
         Left            =   2250
         TabIndex        =   147
         Top             =   1110
         Width           =   1665
         _Version        =   851968
         _ExtentX        =   2937
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   1365
         Left            =   -69910
         TabIndex        =   154
         Top             =   360
         Visible         =   0   'False
         Width           =   11085
         _Version        =   851968
         _ExtentX        =   19553
         _ExtentY        =   2408
         _StockProps     =   79
         Caption         =   "Percepciones:"
         Appearance      =   2
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit vPerIngBrutoStaFe 
            Height          =   285
            Left            =   1650
            TabIndex        =   155
            Top             =   300
            Width           =   1455
            _Version        =   851968
            _ExtentX        =   2558
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            Alignment       =   1
         End
         Begin XtremeSuiteControls.FlatEdit vPerIva 
            Height          =   285
            Left            =   1650
            TabIndex        =   156
            Top             =   630
            Width           =   1455
            _Version        =   851968
            _ExtentX        =   2558
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            Alignment       =   1
         End
         Begin XtremeSuiteControls.FlatEdit vPerImpGanancia 
            Height          =   285
            Left            =   1650
            TabIndex        =   157
            Top             =   990
            Width           =   1455
            _Version        =   851968
            _ExtentX        =   2558
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            Alignment       =   1
         End
         Begin XtremeSuiteControls.FlatEdit vIBBsAs 
            Height          =   285
            Left            =   4740
            TabIndex        =   161
            Top             =   270
            Width           =   1455
            _Version        =   851968
            _ExtentX        =   2558
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            Alignment       =   1
         End
         Begin XtremeSuiteControls.FlatEdit vIBOtros 
            Height          =   285
            Left            =   7230
            TabIndex        =   163
            Top             =   270
            Width           =   1455
            _Version        =   851968
            _ExtentX        =   2558
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            Alignment       =   1
         End
         Begin VB.Label lblTotales 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "IB. Otros:"
            Height          =   195
            Index           =   14
            Left            =   6300
            TabIndex        =   164
            Top             =   330
            Width           =   825
         End
         Begin VB.Label lblTotales 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "IB. Bs.As."
            Height          =   195
            Index           =   13
            Left            =   3150
            TabIndex        =   162
            Top             =   330
            Width           =   1305
         End
         Begin VB.Label lblTotales 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "IB. Sta Fe"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   160
            Top             =   360
            Width           =   1365
         End
         Begin VB.Label lblTotales 
            BackStyle       =   0  'Transparent
            Caption         =   "I.V.A:"
            Height          =   195
            Index           =   11
            Left            =   1050
            TabIndex        =   159
            Top             =   690
            Width           =   495
         End
         Begin VB.Label lblTotales 
            BackStyle       =   0  'Transparent
            Caption         =   "Imp. a la ganancia:"
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   158
            Top             =   1020
            Width           =   1395
         End
      End
      Begin XtremeSuiteControls.FlatEdit vflete 
         Height          =   285
         Left            =   6390
         TabIndex        =   165
         Top             =   450
         Width           =   1665
         _Version        =   851968
         _ExtentX        =   2937
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit vintereses 
         Height          =   285
         Left            =   6390
         TabIndex        =   167
         Top             =   750
         Width           =   1665
         _Version        =   851968
         _ExtentX        =   2937
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit vNetoGravado 
         Height          =   285
         Left            =   6360
         TabIndex        =   169
         Top             =   1140
         Width           =   1665
         _Version        =   851968
         _ExtentX        =   2937
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit6 
         Height          =   285
         Left            =   6360
         TabIndex        =   170
         Top             =   1470
         Width           =   1665
         _Version        =   851968
         _ExtentX        =   2937
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin VB.Label lblTotales 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Neto Gravado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   4260
         TabIndex        =   172
         Top             =   1170
         Width           =   2025
      End
      Begin VB.Label lblTotales 
         BackStyle       =   0  'Transparent
         Caption         =   "Neto  No Gravado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   4680
         TabIndex        =   171
         Top             =   1500
         Width           =   1665
      End
      Begin VB.Label lblTotales 
         BackStyle       =   0  'Transparent
         Caption         =   "Intereses:"
         Height          =   195
         Index           =   16
         Left            =   5310
         TabIndex        =   168
         Top             =   780
         Width           =   705
      End
      Begin VB.Label lblTotales 
         BackStyle       =   0  'Transparent
         Caption         =   "Flete:"
         Height          =   195
         Index           =   15
         Left            =   5580
         TabIndex        =   166
         Top             =   510
         Width           =   435
      End
      Begin VB.Label lblTotales 
         BackStyle       =   0  'Transparent
         Caption         =   " Exento:"
         Height          =   195
         Index           =   10
         Left            =   1590
         TabIndex        =   148
         Top             =   1110
         Width           =   675
      End
      Begin VB.Label lblTotales 
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Total  No Gravado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   150
         TabIndex        =   146
         Top             =   810
         Width           =   2055
      End
      Begin VB.Label lblTotales 
         BackStyle       =   0  'Transparent
         Caption         =   "% Descuento:"
         Height          =   195
         Index           =   8
         Left            =   1200
         TabIndex        =   144
         Top             =   1485
         Width           =   1065
      End
      Begin VB.Label lblTotales 
         BackStyle       =   0  'Transparent
         Caption         =   "I.T.C. :"
         Height          =   165
         Index           =   9
         Left            =   -62290
         TabIndex        =   35
         Top             =   1320
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label lblTotales 
         BackStyle       =   0  'Transparent
         Caption         =   "Retenciones:"
         Height          =   195
         Index           =   6
         Left            =   -62620
         TabIndex        =   32
         Top             =   660
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblTotales 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Iva 27 %:"
         Height          =   195
         Index           =   3
         Left            =   8400
         TabIndex        =   30
         Top             =   1320
         Width           =   945
      End
      Begin VB.Label lblTotales 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Iva 21 %:"
         Height          =   195
         Index           =   2
         Left            =   8370
         TabIndex        =   29
         Top             =   1020
         Width           =   945
      End
      Begin VB.Label lblTotales 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Iva 10.5 %:"
         Height          =   195
         Index           =   1
         Left            =   8400
         TabIndex        =   28
         Top             =   720
         Width           =   945
      End
      Begin VB.Label lblTotales 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Total Gravado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   20
         Top             =   480
         Width           =   2025
      End
   End
   Begin TabDlg.SSTab TabTipoDoc 
      Height          =   2130
      Left            =   14340
      TabIndex        =   0
      Top             =   4110
      Visible         =   0   'False
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   3757
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      Enabled         =   0   'False
      TabCaption(0)   =   "Tipo de Documentos"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Configuración"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture1"
      Tab(1).Control(1)=   "bienes"
      Tab(1).Control(2)=   "tprecio"
      Tab(1).Control(3)=   "lblPFdetalle"
      Tab(1).Control(4)=   "Label11"
      Tab(1).ControlCount=   5
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   288
         Left            =   -73560
         ScaleHeight     =   285
         ScaleWidth      =   2505
         TabIndex        =   16
         Top             =   1360
         Width           =   2500
         Begin VB.OptionButton opDetalle 
            Caption         =   "Con Detalle"
            Height          =   345
            Index           =   0
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   1155
         End
         Begin VB.OptionButton opDetalle 
            Caption         =   "Sin Detalle"
            Height          =   345
            Index           =   1
            Left            =   1230
            TabIndex        =   17
            Top             =   0
            Width           =   1065
         End
      End
      Begin VB.CheckBox bienes 
         Alignment       =   1  'Right Justify
         Caption         =   "> Identificar los productos como Bienes de Capital :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74940
         TabIndex        =   2
         Top             =   1020
         Width           =   3945
      End
      Begin VB.ComboBox tprecio 
         Height          =   315
         ItemData        =   "frmCompras.frx":21B8
         Left            =   -73290
         List            =   "frmCompras.frx":21BA
         TabIndex        =   1
         Text            =   "Pesos ($)"
         Top             =   600
         Width           =   2325
      End
      Begin VB.Label lblPFdetalle 
         Caption         =   "> Cargar Detalle :"
         Height          =   195
         Left            =   -74880
         TabIndex        =   15
         Top             =   1410
         Width           =   1230
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H007EE9FC&
         BackStyle       =   0  'Transparent
         Caption         =   "> Tomar precio en:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   -74895
         TabIndex        =   3
         Top             =   660
         Width           =   1335
      End
   End
   Begin MSAdodcLib.Adodc bfactura 
      Height          =   330
      Left            =   12810
      Top             =   3390
      Visible         =   0   'False
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
   Begin MSAdodcLib.Adodc barticulo 
      Height          =   330
      Left            =   7350
      Top             =   1740
      Visible         =   0   'False
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
      Caption         =   "barticulo"
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
   Begin VB.PictureBox PicDetalle 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   -30
      ScaleHeight     =   3735
      ScaleWidth      =   13545
      TabIndex        =   4
      Top             =   3060
      Width           =   13545
      Begin VB.Frame Frame2 
         Height          =   3375
         Left            =   0
         TabIndex        =   90
         Top             =   -60
         Width           =   13695
         Begin XtremeSuiteControls.GroupBox GroupBox4 
            Height          =   465
            Left            =   90
            TabIndex        =   175
            Top             =   180
            Width           =   13395
            _Version        =   851968
            _ExtentX        =   23627
            _ExtentY        =   820
            _StockProps     =   79
            Appearance      =   2
            BorderStyle     =   2
            Begin XtremeSuiteControls.PushButton b4 
               Height          =   345
               Left            =   2550
               TabIndex        =   176
               ToolTipText     =   "Depura la Grilla de Detalles"
               Top             =   60
               Width           =   1635
               _Version        =   851968
               _ExtentX        =   2884
               _ExtentY        =   609
               _StockProps     =   79
               Caption         =   "Vaciar Detalle"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   0
               TextImageRelation=   4
            End
            Begin XtremeSuiteControls.PushButton b2 
               Height          =   345
               Left            =   30
               TabIndex        =   177
               ToolTipText     =   "Borra el Detalle Seleccionado de la Grilla"
               Top             =   60
               Width           =   2505
               _Version        =   851968
               _ExtentX        =   4419
               _ExtentY        =   609
               _StockProps     =   79
               Caption         =   "Borrar la linea seleccionada"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   0
               TextImageRelation=   4
            End
            Begin XtremeSuiteControls.FlatEdit txtEmpleados 
               Height          =   285
               Index           =   0
               Left            =   6660
               TabIndex        =   178
               Top             =   90
               Width           =   1425
               _Version        =   851968
               _ExtentX        =   2514
               _ExtentY        =   503
               _StockProps     =   77
               BackColor       =   -2147483643
               Locked          =   -1  'True
               MaxLength       =   3
            End
            Begin XtremeSuiteControls.PushButton pbCarga 
               Height          =   285
               Index           =   5
               Left            =   8190
               TabIndex        =   179
               Tag             =   "Vendedor"
               Top             =   90
               Width           =   315
               _Version        =   851968
               _ExtentX        =   556
               _ExtentY        =   503
               _StockProps     =   79
               Caption         =   "..."
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit txtEmpleados 
               Height          =   285
               Index           =   1
               Left            =   8580
               TabIndex        =   180
               Top             =   90
               Width           =   4740
               _Version        =   851968
               _ExtentX        =   8361
               _ExtentY        =   503
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin VB.Label Label1 
               Caption         =   "Empleado: "
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   7
               Left            =   5820
               TabIndex        =   181
               Top             =   120
               Width           =   795
            End
         End
         Begin Grid.KlexGrid KlexDetalle 
            Height          =   2715
            Left            =   90
            TabIndex        =   150
            Top             =   630
            Width           =   13395
            _ExtentX        =   23627
            _ExtentY        =   4789
            GridLinesFixed  =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmCompras.frx":21BC
         End
         Begin Grid.KlexGrid KlexDetalle3 
            Height          =   3195
            Left            =   60
            TabIndex        =   91
            Top             =   150
            Visible         =   0   'False
            Width           =   13455
            _ExtentX        =   23733
            _ExtentY        =   5636
            GridLinesFixed  =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MouseIcon       =   "frmCompras.frx":21D8
         End
      End
      Begin VB.Frame fraCargaDetalle 
         Height          =   525
         Left            =   0
         TabIndex        =   5
         Top             =   3210
         Width           =   13545
         Begin XtremeSuiteControls.CheckBox chkfijo 
            Height          =   315
            Left            =   870
            TabIndex        =   149
            Top             =   150
            Width           =   525
            _Version        =   851968
            _ExtentX        =   926
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Fijo"
            Appearance      =   3
         End
         Begin VB.TextBox f 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   4
            Left            =   10320
            TabIndex        =   10
            Top             =   120
            Width           =   850
         End
         Begin VB.CommandButton cmdArticuloG 
            BackColor       =   &H80000016&
            Height          =   285
            Left            =   8040
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Grabar Documento"
            Top             =   180
            UseMaskColor    =   -1  'True
            Width           =   315
         End
         Begin VB.TextBox f 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   6
            Left            =   12300
            TabIndex        =   12
            Top             =   120
            Width           =   1140
         End
         Begin VB.TextBox f 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   5
            Left            =   11310
            TabIndex        =   11
            Top             =   120
            Width           =   850
         End
         Begin VB.TextBox f 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   3
            Left            =   9390
            TabIndex        =   9
            Top             =   120
            Width           =   850
         End
         Begin VB.TextBox f 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            Left            =   8430
            TabIndex        =   8
            Top             =   120
            Width           =   850
         End
         Begin VB.TextBox f 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   1380
            TabIndex        =   7
            Top             =   90
            Width           =   7005
         End
         Begin VB.TextBox f 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   30
            TabIndex        =   6
            Top             =   120
            Width           =   795
         End
      End
   End
   Begin XtremeSuiteControls.GroupBox GBTipoComprobante 
      Height          =   585
      Left            =   9600
      TabIndex        =   36
      Top             =   2010
      Width           =   3900
      _Version        =   851968
      _ExtentX        =   6879
      _ExtentY        =   1032
      _StockProps     =   79
      Caption         =   "Comprobante"
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.ComboBox cboPuntoDeVenta 
         Height          =   315
         Left            =   1080
         TabIndex        =   37
         Top             =   210
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "1"
         DropDownItemCount=   5
      End
      Begin XtremeSuiteControls.ComboBox cboLetra 
         Height          =   315
         Left            =   180
         TabIndex        =   38
         Top             =   210
         Width           =   765
         _Version        =   851968
         _ExtentX        =   1349
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "A"
         DropDownItemCount=   5
      End
      Begin XtremeSuiteControls.FlatEdit txtNroComprobante 
         Height          =   315
         Left            =   1980
         TabIndex        =   39
         Top             =   210
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3096
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
         MaxLength       =   8
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1395
      Left            =   9570
      TabIndex        =   67
      Top             =   690
      Width           =   3915
      _Version        =   851968
      _ExtentX        =   6906
      _ExtentY        =   2461
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin Aplisoft_CajasDeTexto.TxF vFechaIva 
         Height          =   285
         Left            =   2040
         TabIndex        =   173
         Top             =   780
         Width           =   1785
         _ExtentX        =   3149
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
      Begin Aplisoft_CajasDeTexto.TxF vfechaPago 
         Height          =   285
         Left            =   2040
         TabIndex        =   184
         Top             =   1080
         Width           =   1785
         _ExtentX        =   3149
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
         Index           =   0
         Left            =   2040
         TabIndex        =   188
         Top             =   120
         Width           =   1785
         _ExtentX        =   3149
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
         BackStyle       =   0
         MinValor        =   36526
      End
      Begin Aplisoft_CajasDeTexto.TxF dtpFecha 
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   189
         Top             =   450
         Width           =   1785
         _ExtentX        =   3149
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
         BackStyle       =   0
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de pago:"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   210
         TabIndex        =   185
         Top             =   1110
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Imp. IVA :"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   630
         TabIndex        =   174
         Top             =   840
         Width           =   1200
      End
      Begin VB.Label lblProveedor 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Emision :"
         Height          =   195
         Index           =   6
         Left            =   480
         TabIndex        =   89
         Top             =   180
         Width           =   1350
      End
      Begin VB.Label lblProveedor 
         Caption         =   "Fecha de Vencimiento :"
         Height          =   195
         Index           =   7
         Left            =   150
         TabIndex        =   88
         Top             =   510
         Width           =   2115
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtTotal 
      Height          =   315
      Left            =   11520
      TabIndex        =   151
      Top             =   7590
      Width           =   1995
      _Version        =   851968
      _ExtentX        =   3519
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   255
      BackColor       =   -2147483643
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
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   255
      Left            =   11520
      TabIndex        =   152
      Top             =   8250
      Width           =   2085
      _Version        =   851968
      _ExtentX        =   3678
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Volver a calcular Totales"
      UseVisualStyle  =   -1  'True
   End
   Begin MSAdodcLib.Adodc bdetalle 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
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
   Begin VB.Label lblTotales 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total del Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   11310
      TabIndex        =   153
      Top             =   7260
      Width           =   2085
   End
   Begin VB.Label lblCantidadRemito 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Left            =   15510
      TabIndex        =   104
      Top             =   4860
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label lblCantidadPresupuesto 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0080FF80&
      Height          =   195
      Left            =   11040
      TabIndex        =   103
      Top             =   5280
      Width           =   345
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "> Presupuesto: "
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   9510
      TabIndex        =   102
      Top             =   5280
      Width           =   1110
   End
   Begin VB.Label lblSinDetalle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Asiento de factura sin detalle."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Left            =   3075
      TabIndex        =   14
      Top             =   4905
      Width           =   4695
   End
End
Attribute VB_Name = "frmCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vOpenGrilla() As Boolean
Dim vIdPFactura As Long
Public vvdescrip, vvcodigo, vvvdescrip, vvvcodigo, vcheque As String
Public vvpdolar, vvpventa, vpcosto  As Double
Public vGrabaModo, vvvnrointerno As Integer
Dim vRemitoCompras As Long
Dim rsArticulosCompra As ADODB.Recordset
Dim rsProveedores As ADODB.Recordset, rsArticulos As ADODB.Recordset
Dim vLeyendaAsiento As String, vTotalAsiento As Double
Dim vfacturaDuplicadaMensaje As String
Dim vidArticulos As Long
Dim articuloNuevo As Boolean
Dim vnrocomprobante, vcomentario As String
Dim validarGuardado As Boolean
Dim vgTsubtotal, vgTiva105, vgTiva21, vgTiva27, vgTPdescuento, vgTimpuesto, vgTtotal As Double


Private Sub AcreditarCheque()
    On Error Resume Next

    Dim rsCCP As New ADODB.Recordset, sqlCCP As String
    
    sqlCCP = "SELECT * FROM pcuentascorrientes"
    
    With rsCCP
        Call .Open(sqlCCP, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If vGrabaModo = 1 Then
            .Find ("Remito = " & vRemitoCompras)

            If .EOF Then .AddNew
        Else
            .AddNew
        End If

        .Fields("Fecha").Value = Date
        .Fields("Codigo").Value = txtProveedor(0).Tag
        .Fields("Nombre").Value = txtProveedor(0).Text
        .Fields("Debito").Value = Val(txtTotal.Text)
        .Fields("Credito").Value = 0
        .Fields("Comentario").Value = "Acreditación cheque Nº: " & Trim(vcheque)
    
        .Update
    
    End With
    
    sqlCCP = ""

    If rsCCP.State = 1 Then
        rsCCP.Close
        Set rsCCP = Nothing
    End If

    If Err Then GrabarLog "AcreditarCheque", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Function BuscarArticulos() As Boolean
On Error Resume Next

    Set rsArticulosCompra = New ADODB.Recordset
    Dim sqlArticulos As String
    
    sqlArticulos = "SELECT * FROM articulos WHERE (CodigoBarra = '" & Trim(f(1).Text) & "')"
    
    With rsArticulosCompra
        .CursorLocation = adUseClient
        Call .Open(sqlArticulos, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then
            f(1).Tag = .Fields("Codigo").Value
            f(2).Tag = .Fields("PCosto").Value
            If vidArticulos = 0 Then vidArticulos = .Fields("idArticulos").Value  ' guardo el id el articulo en el caso que no lo haya cargado desde la búsqueda del artículo
            
            MostrarDatos
            BuscarArticulos = True
        Else
            If .State = 1 Then .Close
                
            sqlArticulos = ""
            
            If vOpenGrilla(1) = True Then
                With rsArticulos
                    If Not (.EOF = True) And Not (.BOF = True) Then
                        sqlArticulos = "SELECT * FROM Articulos WHERE (codigo =  '" & Trim(.Fields("Codigo").Value) & "')"
                    Else
                        sqlArticulos = "SELECT * FROM Articulos WHERE (codigo =  '" & f(1).Text & "')"
                    End If
                End With
            Else
                sqlArticulos = "SELECT * FROM Articulos WHERE (codigo = '" & Trim(f(1).Text) & "')"
            End If

            Call .Open(sqlArticulos, ConnDDBB, adOpenStatic, adLockReadOnly)
            
            If Not .EOF = True Then
                f(1).Tag = .Fields("codigo").Value
                f(2).Tag = .Fields("PCosto").Value
                If vidArticulos = 0 Then vidArticulos = .Fields("idArticulos").Value  ' guardo el id el articulo en el caso que no lo haya cargado desde la búsqueda del artículo
                MostrarDatos
                BuscarArticulos = True
            Else
                If .State = 1 Then .Close
                
                sqlArticulos = ""
                sqlArticulos = "SELECT * FROM articulos WHERE (descrip LIKE '%" & Trim(f(1).Text) & "%')"
                
                Call .Open(sqlArticulos, ConnDDBB, adOpenStatic, adLockReadOnly)
            
                If .RecordCount >= 1 Then
                    f(1).Tag = .Fields("codigo").Value
                    'f(2).Tag = .Fields("PCosto").Value
                    If vidArticulos = 0 Then vidArticulos = .Fields("idArticulos").Value  ' guardo el id el articulo en el caso que no lo haya cargado desde la búsqueda del artículo

                    MostrarDatos
                    BuscarArticulos = True
                    If .RecordCount > 1 Then

                    End If
                
                ElseIf .RecordCount = 0 Then
                    'MsgBox "El articulo NO existe carguelo y vuelva al Remito", vbInformation, "Mensaje ..."
                    'f(1).Text = ""
                    'f(2).SetFocus
                    BuscarArticulos = False
                    articuloNuevo = True
                Else
                
                End If
            End If
        End If

    End With
    
    sqlArticulos = ""
    
If Err Then GrabarLog "BuscarArticulos", Err.Number & " " & Err.Description, Me.Name
End Function
Private Function BuscarProveedor() As Boolean
    On Error Resume Next
    
    Dim rsProveedores As New ADODB.Recordset, sqlProveedores As String
    
    sqlProveedores = "SELECT * FROM proveedores WHERE (Nombre = '" & Trim(txtProveedor(0).Text) + "') OR (Codigo = '" & Trim(txtProveedor(0).Text) & "')"
    
    With rsProveedores
        Call .Open(sqlProveedores, ConnDDBB, adOpenStatic, adLockReadOnly)

        If .EOF = True Then
            'frmBuscarProveedor.o = 1
            'frmBuscarProveedor.txtProveedor = v(0).Text
            'frmBuscarProveedor.Show
            'frmBuscarProveedor.txtProveedor.SetFocus
            BuscarProveedor = Not True
        
        Else
            BuscarProveedor = True
            txtProveedor(0).Tag = .Fields("Codigo").Value
            txtProveedor(0).Text = EsNulo(.Fields("Nombre").Value)
            txtProveedor(1).Text = EsNulo(.Fields("Direccion").Value)
            txtProveedor(2).Text = EsNulo(.Fields("Localidad").Value)
            txtProveedor(3).Text = EsNulo(.Fields("Telefono").Value)
            txtProveedor(4).Text = TraerDato("TipoIva", "idTipoIva =  '" & EsNulo(.Fields("idTipoIva").Value) & "'", "TipoIva")
            txtProveedor(5).Text = EsNulo(.Fields("Cuit").Value)
            
            
            TabTipoDoc.Enabled = True
            TabTotales.Enabled = True
            
            
           ' opTipoDocumento(7).Value = True
           ' opTipoDocumento_Click (7)
            txtTipoMovimiento(0).SetFocus
            
        End If
    
        'Set .Recordset = Nothing
    End With
    
    sqlProveedores = ""
    
    If rsProveedores.State = 1 Then
        rsProveedores.Close
        Set rsProveedores = Nothing
    End If
    
    If Err Then GrabarLog "BuscarProveedor", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub CalcularTotales()
    On Error Resume Next
    Dim vvsubtotal As Double
    
    Dim vSubTotal As Double
    Dim vTnoGravado As Double, vIva() As Double, vDescuentoTotal As Double, vImpuesto As Double, vRetenciones As Double, vPercepciones As Double, vITC As Double
    ReDim vIva(2) As Double
    
    vSubTotal = KlexDetalle.Aggregate(klexSTSum, 1, 11, KlexDetalle.Rows - 1, 11) ' calcula el total de la columna 11
    
    vIva(0) = 0
    vIva(1) = 0
    vIva(2) = 0
    vTnoGravado = 0
    


      
    
    Dim i As Integer
    
    For i = 1 To KlexDetalle.Rows - 1
    
        Select Case KlexDetalle.TextMatrix(i, 9)
                
            Case "10.50"
                vIva(0) = Val(vIva(0)) + Val(KlexDetalle.TextMatrix(i, 11)) * Val(KlexDetalle.TextMatrix(i, 9)) / 100
            
            Case "21.00"
                vIva(1) = Val(vIva(1)) + Val(KlexDetalle.TextMatrix(i, 11)) * Val(KlexDetalle.TextMatrix(i, 9)) / 100
            
            Case "27.00"
                vIva(2) = Val(vIva(2)) + Val(KlexDetalle.TextMatrix(i, 11)) * Val(KlexDetalle.TextMatrix(i, 9)) / 100
            Case Else
                vTnoGravado = vTnoGravado + CDbl(KlexDetalle.TextMatrix(i, 11))
        End Select
    
    Next
    

    
    vDescuentoTotal = Val(txtPorcentajeDescuento.Text) * vSubTotal / 100
    
    
    
    txtSubtotal.Text = vSubTotal
    
    
    vSubTotal = vSubTotal - vDescuentoTotal
   
   
   
   txtSubtotal.Text = vSubTotal - vTnoGravado '
    'vvsubtotal = vSubTotal - vTnoGravado
    
    
    Me.vNoGravado = vTnoGravado  ' pongo que el form el valor de total no gravado
    
    txtIva(0).Text = vIva(0)
    If txtIva(0).Text > 0 Then
        chkIva(0).Value = 1
        txtIva(0).Enabled = True
    Else
        chkIva(0).Value = 0
        txtIva(0).Enabled = Not True
    End If
    
    txtIva(1).Text = vIva(1)
    If txtIva(1).Text > 0 Then
        chkIva(1).Value = 1
        txtIva(1).Enabled = True
    Else
        chkIva(1).Value = 0
        txtIva(1).Enabled = Not True
    End If
    
    txtIva(2).Text = vIva(2)
    If txtIva(2).Text > 0 Then
        chkIva(2).Value = 1
        txtIva(2).Enabled = True
    Else
        chkIva(2).Value = 0
        txtIva(2).Enabled = Not True
    End If
    
    txtPorcentajeDescuento.Text = Val(txtPorcentajeDescuento.Text)
    
    
   ' -------- agrupo todas las percepciones --------------
    vPercepciones = Val(Me.vPerImpGanancia) + Val(Me.vPerIngBrutoStaFe) + Val(Me.vPerIva) + Val(Me.vIBBsAs) + Val(Me.vIBOtros)
    
    '---------- Subtotal-----------------
    vSubTotal = Val(Me.txtSubtotal) + CDbl(Me.vNoGravado) + Val(Me.vExento) + Val(Me.vflete)
    'vSubTotal = vSubTotal + Val(Me.vNoGravado) + Val(Me.vExento) + Val(Me.vflete)
    
    txtTotal.Text = vSubTotal + Val(vIva(0)) + Val(vIva(1)) + Val(vIva(2)) - vDescuentoTotal + vPercepciones
    
    
    
    vgTtotal = txtTotal.Text
    
    vgTsubtotal = vSubTotal
    
    vgTPdescuento = 0
    
   
    
    
    If Err < 0 Then
        Exit Sub
        MsgBox "Error Calculando los TOTALES", vbExclamation, "Mensaje ..."
        GrabarLog "CalcularTotales", Err.Number & " " & Err.Description, Me.Name
    End If
End Sub
Public Sub CargarBien()
On Error Resume Next
    
    f(1).Text = vvdescrip
    
    If tprecio = "Preguntar" Then
        If MsgBox("¿ Desea seleccionar el precio en Pesos ($) ?", vbYesNo) = vbYes Then
            f(2).Text = Val(vvpventa)
            tprecio.Text = "Pesos ($)"
        Else
            f(2).Text = inulo(vvpdolar) * Val(gdolar)
            tprecio.Text = "Dolar (u$s)"
        End If
    End If
    
    If Trim(tprecio.Text) = "Dolar (u$s)" Then f(2).Text = Val(vvpdolar)
    If Trim(tprecio.Text) = "Pesos ($)" Then f(2).Text = Val(vvpventa)
    
    f(1).Tag = vvcodigo
    
If Err Then GrabarLog "CargarBien", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CargarPFDetalle(vnroremito As Long)
    On Error Resume Next
    
    Dim rsDetalleCompra As New ADODB.Recordset, sqlDetalleCompra As String, j As Integer
    
    'sqlDetalleCompra = "SELECT * FROM pfdetalle WHERE (Fecha = '" & strfechaMySQL(dtpFecha(0).Value) + "') AND (remito = " & Val(vnroremito) & ")"
        
     sqlDetalleCompra = "SELECT * FROM pfdetalle WHERE (remito = " & Val(vnroremito) & ")"
        
    With rsDetalleCompra
        If Not .State = 0 Then .Close
        Call .Open(sqlDetalleCompra, ConnDDBB, adOpenKeyset, adLockOptimistic)
        
        FormatoGrillaDetalle (.RecordCount)
    
        If Not .EOF = True Then
            .MoveFirst
            
            j = 1
            
            Do Until .EOF = True
                
                KlexDetalle.TextMatrix(j, 1) = EsNulo(.Fields("idPFDetalle").Value)
                KlexDetalle.TextMatrix(j, 2) = EsNulo(.Fields("Fecha").Value)
                KlexDetalle.TextMatrix(j, 3) = EsNulo(.Fields("Remito").Value)
                KlexDetalle.TextMatrix(j, 4) = "[" & EsNulo(.Fields("Codigo").Value) & "]"
                KlexDetalle.TextMatrix(j, 5) = EsNulo(.Fields("cantidad").Value)
                KlexDetalle.TextMatrix(j, 6) = EsNulo(.Fields("Descripcion").Value)
                KlexDetalle.TextMatrix(j, 7) = EsNulo(.Fields("Precio").Value)
                KlexDetalle.TextMatrix(j, 8) = EsNulo(.Fields("Descuento").Value)
                KlexDetalle.TextMatrix(j, 9) = EsNulo(.Fields("TipoIva").Value)
                KlexDetalle.TextMatrix(j, 10) = EsNulo(.Fields("Impuesto").Value)
                KlexDetalle.TextMatrix(j, 11) = EsNulo(.Fields("Total").Value)
                KlexDetalle.TextMatrix(j, 21) = EsNulo(.Fields("Confirmado").Value)
                
                .MoveNext
                j = j + 1
            Loop

        End If
            
    End With
    
    sqlDetalleCompra = ""
    
    If rsDetalleCompra.State = 1 Then
        rsDetalleCompra.Close
        Set rsDetalleCompra = Nothing
    End If
    
If Err Then GrabarLog "CargarPFDetalle", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CargarComprobante(vremito)
    On Error Resume Next

    With bfactura
    
        dtpFecha(0).Value = EsNulo(.Recordset("Fecha").Value)
        vFechaIva.Value = EsNulo(.Recordset("FechaIva").Value)
        dtpFecha(1).Value = EsNulo(.Recordset("FechaVencimiento").Value)
        
        
        
        
        vfechaPago.Value = EsNulo(.Recordset("fechapago").Value)
        
        txtProveedor(0).Tag = EsNulo(.Recordset("Codigo").Value)
        txtProveedor(0).Text = EsNulo(.Recordset("nombre").Value)
        txtProveedor(1).Text = EsNulo(.Recordset("domicilio").Value)
        txtProveedor(2).Text = EsNulo(.Recordset("localidad").Value)
        txtProveedor(3).Text = EsNulo(.Recordset("telefono").Value)
        txtProveedor(4).Text = EsNulo(.Recordset("iva").Value)
        txtProveedor(5).Text = EsNulo(.Recordset("cuit").Value)
        
        vRemitoCompras = Val(.Recordset("remito").Value)
        txtSubtotal.Text = Val(.Recordset("SubTotal").Value)
        
        txtTotal.Text = Val(.Recordset("total").Value)
        txtDescuento.Text = Val(.Recordset("descuento").Value)
        
        txtNroInterno.Text = EsNulo(.Recordset("NroInterno").Value)
        
        cboLetra.Text = EsNulo(.Recordset("Letra").Value)
        cboPuntoDeVenta.Text = EsNulo(.Recordset("PuntoDeVenta").Value)
        txtNroComprobante.Text = String(8 - Len(.Recordset("NComprobante").Value), "0") & (.Recordset("NComprobante").Value)
        
        CargarTipoComprobante (.Recordset("Tipo").Value)
        CargarIva (.Recordset("Remito").Value)
        
        
        'Me.txtAuxiliares(0) = EsNulo(.Recordset("NoGravado").Value)
         Me.txtPorcentajeDescuento = EsNulo(.Recordset("PorcentajeDescuento").Value)
        'Me.txtAuxiliares(1) = EsNulo(.Recordset("Retenciones").Value)
        'Me.txtAuxiliares(2) = EsNulo(.Recordset("Percepciones").Value)
        'Me.txtAuxiliares(3) = EsNulo(.Recordset("ITC").Value)
        
         ' --------------- Todo lo referido a persepciones y retenciones ----------------
        Me.vNoGravado.Text = CDbl(.Recordset("NoGravado").Value)
        Me.vflete = .Recordset("Flete")
        
        
        Me.txtPorcentajeDescuento = .Recordset("PorcentajeDescuento").Value
                
        ' ----------- percepciones --------------------------
        Me.vPerIngBrutoStaFe = .Recordset("PerIngBrutoStaFe").Value
        
        Me.vIBBsAs = .Recordset("IBBsAs").Value
        Me.vIBOtros = .Recordset("IBOtros").Value
      
        Me.vPerIva = .Recordset("PerIva").Value
        Me.vPerImpGanancia = .Recordset("PerImpGanancia").Value
        '-------------------------------------------------------------------------------
        
    End With
    
    If Err < 0 Then
        MsgBox "Los datos del remito no son compatibles", vbCritical
        GrabarLog "CargarComprobante", Err.Number & " " & Err.Description, Me.Name
    End If

End Sub
Private Sub CargarIva(vremito As Long)
On Error Resume Next

    Dim rsIvaFacturaCompra As New ADODB.Recordset, sqlIvaFacturaCompra As String
    
    sqlIvaFacturaCompra = "SELECT * FROM IvaFacturaCompra WHERE (Remito = " & Val(vremito) & ")"
    
    With rsIvaFacturaCompra
        Call .Open(sqlIvaFacturaCompra, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then
            
            If Not Val(.Fields("Iva105").Value) = 0 Then
                txtIva(0).Text = EsNulo(.Fields("Iva105").Value)
                chkIva(0).Value = 1
            Else
                chkIva(0).Value = 0
                txtIva(0).Text = ""
            End If
            
            If Not Val(.Fields("Iva210").Value) = 0 Then
                chkIva(1).Value = 1
                txtIva(1).Text = EsNulo(.Fields("Iva210").Value)
            Else
                chkIva(1).Value = 0
                txtIva(1).Text = ""
            End If
            
            If Not Val(.Fields("Iva270").Value) = 0 Then
                txtIva(0).Text = EsNulo(.Fields("Iva270").Value)
                chkIva(0).Value = 1
            Else
                chkIva(2).Value = 0
                txtIva(2).Text = ""
            End If
            
            txtAuxiliares(0).Text = EsNulo(.Fields("Retenciones").Value)
        End If
        
    
    End With

    sqlIvaFacturaCompra = ""

    If rsIvaFacturaCompra.State = 1 Then
        rsIvaFacturaCompra.Close
        Set rsIvaFacturaCompra = Nothing
    End If
    
If Err < 0 Then
    GrabarLog "CargarIva", Err.Number & " " & Err.Description, Me.Name
    MsgBox "Cuidado!. La factura puso no haberse grabado adecuadamente"
End If
End Sub
Public Sub CargarRemito(vnroremito As Long)   'Carga un remito confeccionado
    On Error Resume Next
    
    With bfactura
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM pfactura WHERE (remito = " & vnroremito & ")"
        .Refresh

        If Not .Recordset.EOF = True Then
            
            CargarPFDetalle (vnroremito)
            CargarComprobante (vnroremito)
        End If
    End With
    
If Err Then GrabarLog "CargarRemito", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CargarTipoComprobante(vtipo)
On Error Resume Next

    Select Case vtipo

        Case "Fact A"
            opTipoDocumento(0).Value = True

        Case "Fact B"
            opTipoDocumento(0).Value = True

        Case "Fact C"
            opTipoDocumento(0).Value = True
        
        Case "Presupuesto"
            opTipoDocumento(1).Value = True

        Case "Nota C"
            opTipoDocumento(2).Value = True

        Case "Documento"
            opTipoDocumento(6).Value = True

        Case "Remito"
            opTipoDocumento(4).Value = True
    End Select

If Err Then GrabarLog "CargarTipoComprobante", Err.Number & " " & Err.Description, Me.Name
End Sub



Private Sub a_Click()

End Sub

Private Sub b_Click()

End Sub

Private Sub b2_Click()
On Error Resume Next
            KlexDetalle.RemoveItem KlexDetalle.RowSel
            
            
            CalcularTotales
If Err Then Exit Sub
End Sub

Private Sub b4_Click()
Dim i As Integer, j As Integer
            
            Call FormatoGrillaDetalle(1)
            
            Me.KlexDetalle.Tag = 0  ' para que arranque en el primero de la fila
            
        
End Sub

Private Sub cboLetra_GotFocus()
On Error Resume Next

    With cboLetra
        .Clear
        .AddItem ("A")
        .AddItem ("B")
        .AddItem ("C")
        .AddItem ("X")
    End With

If Me.opTipoDocumento(6).Value Then Me.cboLetra = "X"


If Err Then GrabarLog "cboLetra_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cboLetra_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then cboPuntoDeVenta.SetFocus

If Err Then GrabarLog "cboLetra_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub cboPuntoDeVenta_GotFocus()
On Error Resume Next
Exit Sub ' saco las curusales para que nose confundan con las que ponen

    With cboPuntoDeVenta
        .Clear
        .AddItem ("0001")
        .AddItem ("0002")
        .AddItem ("0003")
        .AddItem ("0004")
        .AddItem ("0005")
        .AddItem ("0006")
        .AddItem ("0007")
        .AddItem ("0008")
        .AddItem ("0009")
        .AddItem ("0010")
        .AddItem ("0011")
        .AddItem ("0012")
        .AddItem ("0013")
        .AddItem ("0014")
        .AddItem ("0015")
    End With

If Err Then GrabarLog "cboPuntoDeVenta_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub cboPuntoDeVenta_KeyPress(KeyAscii As Integer)
On Error Resume Next

    
    If KeyAscii = 13 Then txtNroComprobante.SetFocus
    
If Err Then GrabarLog "cboPuntoDeVenta_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cboPuntoDeVenta_LostFocus()
'On Error Resume Next

 '   cboPuntoDeVenta.Text = String(4 - Len(cboPuntoDeVenta.Text), "0") & Val(cboPuntoDeVenta.Text)

'If Err Then GrabarLog "cboPuntoDeVenta_LostFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub cmdCerrarPago_Click()
GBCaja.Visible = False
End Sub

Private Sub chkIva_Click(Index As Integer)
On Error Resume Next

    If chkIva(Index).Value = 1 Then
        
        If Not Val(txtSubtotal.Text) = 0 Then
            txtIva(Index).Enabled = True
            'txtIva(Index).Text = chkIva(Index).Tag * Val(txtSubtotal.Text) / 100
            
        Else
            txtIva(Index).Text = ""
            txtIva(Index).Enabled = False
        End If
    Else
        txtIva(Index).Enabled = False
        txtIva(Index).Text = ""
    End If
    
    txtSubtotal_change

If Err Then GrabarLog "chkIva_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdAcciones_Click(Index As Integer)
On Error Resume Next

vnrocomprobante = Me.txtNroComprobante

    Select Case Index
    
        Case 0
            BorrarDetalle
        
        Case 1
            FormatoGrillaDetalle (1)

        Case 2
            LimpiarCampos
        
        Case 3
            
            fGuardarDoc
            
          '  Call Asiento
                     
        Case 4
            
            Call cmdAcciones_Click(3)
            
            
            If Not validarGuardado Then Exit Sub
            
            Call init
            
            Call Imprimir(vRemitoCompras, TipoDocumento)
            
           ' Call Asiento
            
            
            Case 5
           ' Unload Me
            Case 6
            
                If controlNroPFactura Then Exit Sub
                
                vGrabaModo = 0
                
                fGuardarDoc
    
    End Select





If Err Then GrabarLog "cmdAcciones_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub fGuardarDoc()
            
            If Not ValirParaGuardar Then Exit Sub
            
            If Me.vGrabaModo = 1 Then
                'MsgBox "No se puede modificar los documentos de compras", vbInformation, "Mensaje."
                'Exit Sub
            End If
            
            
            Dim vnrointerno2
            
            'Me.txtNroInterno = UltimoNroInterno2 + 1
            
            Me.txtNroInterno = UltimoNroInterno2 ' panic .. sacarlo
            
            vnrointerno2 = txtNroInterno
            
            'If controlNroFactura And vGrabaModo = 0 Then
            '    MsgBox "Este documento ya fue grabado anteriormente." + Chr(13) + vfacturaDuplicadaMensaje, vbCritical, "Documento duplicado..."
            '    Exit Sub
            ' End If

            
            If Not opTipoDocumento(7).Value = True Then
                GuardarComprobante ' guardar 1
            Else
                If Trim(txtTipoMovimiento(0).Text) = "CD" Then
                    GBCaja.Visible = True
                    txtCaja(8).Text = Trim(txtIB(7).Text)
                    txtCaja(7).Text = Me.txtIB(10)
                    txtCaja(0).SetFocus
                Else
                    GuardarComprobante
                    'Me.txtNroInterno = vnrointerno2 + 1
                
                End If
            End If
                        
                        
           Call Asiento
End Sub

Private Sub imprimirDocumento()

' impresión para comuna


End Sub

Function controlNroPFactura() As Boolean ' true es duplicado
 On Error Resume Next

Dim c1 As String
Dim sql As String
'arreglar: sucursal por punto de venta
sql = ""

sql = " codigo = " + Me.txtProveedor(0).Tag + " and NComprobante=" + Str(Val(Me.txtNroComprobante.Text)) + " and PuntoDeVenta='" + Trim(Me.cboPuntoDeVenta.Text) + "' and letra='" + Trim(Me.cboLetra) + "'"    ' and TipoMovimiento='" + Trim(Me.txtTipoMovimiento(0).Text) + "'"


c1 = TraerDato("pfactura", sql, "ncomprobante")
If (Not Trim(c1) = "") Or Val(c1) > 0 Then

'-----------------------------------------------------------------------------------------------------------------------------
    vfacturaDuplicadaMensaje = ""
    vfacturaDuplicadaMensaje = vfacturaDuplicadaMensaje + Chr(13) + "> Cli/Provee :" + TraerDato("pfactura", sql, "Codigo")
    vfacturaDuplicadaMensaje = vfacturaDuplicadaMensaje + Chr(13) + "> Fecha :" + TraerDato("pfactura", sql, "Fecha")
    vfacturaDuplicadaMensaje = vfacturaDuplicadaMensaje + Chr(13) + "> Tipo :" + TraerDato("pfactura", sql, "TipoMovimiento")
    Me.txtNroInterno = Val(txtNroInterno) - 1
'-----------------------------------------------------------------------------------------------------------------------------

    controlNroPFactura = True
    
    MsgBox "Nro de factura duplicado", vbInformation

Else
    controlNroPFactura = False
    
End If

If Err Then
MsgBox "No se puede controlar si este documento fue grabado anteriormente." + Chr(13) + "Consulte con el servicio técnico.", vbCritical
Exit Function
End If
End Function



Function controlNroFactura() As Boolean ' true es duplicado
 On Error Resume Next

Dim c1 As String
Dim sql As String
'arreglar: sucursal por punto de venta
sql = ""

sql = " NComprobante=" + Str(Val(Me.txtNroComprobante.Text)) + " and PuntoDeVenta='" + Trim(Me.cboPuntoDeVenta.Text) + "' and letra='" + Trim(Me.cboLetra) + "'"  ' and TipoMovimiento='" + Trim(Me.txtTipoMovimiento(0).Text) + "'"


c1 = TraerDato("pfactura", sql, "ncomprobante")
If (Not Trim(c1) = "") Or Val(c1) > 0 Then

'-----------------------------------------------------------------------------------------------------------------------------
    vfacturaDuplicadaMensaje = ""
    vfacturaDuplicadaMensaje = vfacturaDuplicadaMensaje + Chr(13) + "> Cli/Provee :" + TraerDato("pfactura", sql, "Codigo")
    vfacturaDuplicadaMensaje = vfacturaDuplicadaMensaje + Chr(13) + "> Fecha :" + TraerDato("pfactura", sql, "Fecha")
    vfacturaDuplicadaMensaje = vfacturaDuplicadaMensaje + Chr(13) + "> Tipo :" + TraerDato("pfactura", sql, "TipoMovimiento")
    Me.txtNroInterno = Val(txtNroInterno) - 1
'-----------------------------------------------------------------------------------------------------------------------------

    controlNroFactura = True
    Me.txtNroInterno = Val(txtNroInterno) - 1
Else
    controlNroFactura = False
    
End If
If Err Then
MsgBox "No se puede controlar si este documento fue grabado anteriormente." + Chr(13) + "Consulte con el servicio técnico.", vbCritical
Exit Function
End If
End Function

Private Sub BorrarDetalle()
    On Error Resume Next

    If vPFDetalle = True Then
        If vGrabaModo = 0 Then
            
            Select Case KlexDetalle.Rows
            
                Case 2
                    FormatoGrillaDetalle (1)
                    'ReFormatear toda la Grilla
                    
                Case Else
                    KlexDetalle.RemoveItem (KlexDetalle.Row)
             
             End Select
             
             CalcularTotales
        
        Else
            MsgBox "No Puede borrar un detalle de la Factura si esta MODIFICANDO", vbExclamation, "Mensaje ..."
        End If
    Else
        MsgBox "Debe tener activado el Modo de FDETALLES en Facturas de Compras", vbExclamation, "Mensaje ..."
    End If
    
    If Err Then GrabarLog "cmdBorrar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub GuardarComprobante()
    On Error Resume Next

    Dim vPeriodoFactura As String
    
    
    validarGuardado = False
    
   ' If (Me.txtTipoMovimiento(0).Text = "CD" Or Me.txtTipoMovimiento(0).Text = "CC") And (Trim(Me.txtBancoCheque(0).Text) = "") Then
   '     MsgBox "Debe ingresar una cuenta de Caja válida.", vbInformation, "Mensaje ..."
   '     Exit Sub
   ' End If
    
    
   ' If Not TraerDato("NroInterno", "NroInterno = " & Val(txtNroInterno.Text) & "", "TablaUsada") = "" Then
   '     MsgBox "El Numero Interno ha sido usado en otro Registro", vbInformation, "Mensaje ..."
   '     Exit Sub
   ' End If
    
    If Val(vRemitoCompras) = 0 Then
        MsgBox "No se puede guardar el comprobante ya que no existe el Nro de Remito Interno", vbExclamation, "Mensaje ..."
        Exit Sub
    End If
    
    If Val(txtTotal.Text) = 0 Then
        MsgBox "El total del comprobante es cero."
        Exit Sub
    End If
    
    vPeriodoFactura = Year(dtpFecha(0).Value) & Mid(dtpFecha(0).Value, 4, 2)
    
    If vPeriodoFactura = TraerDato("Cerrado", "Periodo = '" & vPeriodoFactura & "'", "Periodo") Then
        MsgBox "La Factura pertenece a un periodo de Iva Compra Ya Cerrado!!!", vbInformation, "Mensaje ..."
        Exit Sub
    End If
    
    If Not Val(txtNroComprobante.Text) > 0 Then
        MsgBox "Debe ingresar un número de Comprobante Correcto !", vbCritical, "Mensaje ..."
        txtNroComprobante.SetFocus
        Exit Sub
    End If
    
      
    If (Me.cboLetra = "" Or Me.cboPuntoDeVenta.Text = "") Then
        MsgBox "Debe ingresar punto de venta y letra  del documento !", vbCritical, "Mensaje ..."
        txtNroComprobante.SetFocus
        Exit Sub
    End If
    
    
    
    Dim vGraba As Boolean, i As Integer
    
    For i = 0 To 8
        If opTipoDocumento(i).Value = True Then
            vGraba = True
            Exit For
        Else
            vGraba = Not True
        End If
    Next
    
    If Not vGraba = True Then
        MsgBox "No se Pudo elegir un Tipo de Comprobante, Seleccionelo y vuelva a grabar", vbExclamation, "Mensaje ..."
        Exit Sub
    End If

    MousePointer = vbHourglass
    
    If opTipoDocumento(7).Value = True Then
        vLeyendaAsiento = Trim(txtIB(7).Text)
        If Val(txtIB(10).Text) < 0 Then
            vTotalAsiento = Val(txtIB(10).Text) * (-1)
        Else
            vTotalAsiento = Val(txtIB(10).Text)
        End If
    End If
    
    
    
'--------------- fin validacion --------------------------

    validarGuardado = True
    Guardar ' guarda en fdetalle y factura ' guarfar 1.1
        
   ' WCtaCte vRemitoCompras  ' graba el movimiento en la ctacte
     
    'cobro ' llama al mòdulo de cobro
    
'    Asiento ' llama al módulo de asiento
    
    'WCaja ' guarda el movimiento en caja en el caso que sea contado
        
    RecargarForm
    
    MousePointer = vbDefault
    
    Call init

  '  If Not LeerXml("puesto") = "comuna" And LeerXml("IncluyeContabilidad") = "true" Then
  '      Call Asiento
  '  End If
    
    If Err < 0 Then
        MsgBox "Error! Revisar las operaciones ", vbCritical, "Mensaje ..."
        GrabarLog "GuardarComprobante", Err.Number & " " & Err.Description, Me.Name
    End If

End Sub
Private Sub WCaja()
' sale si no es contado
If (Me.txtTipoMovimiento(0).Text = "CD") Then
    Call GuardarBancosMovimientos(Me.txtBancoCheque(0), Val(Me.txtIB(10)), 0, 0, Me.txtIB(7), 0, Val(Me.txtNroInterno.Text), Me.dtpFecha(0).Value)
End If

If (Me.txtTipoMovimiento(0).Text = "CC") Then
    Call GuardarBancosMovimientos(Me.txtBancoCheque(0), 0, Val(Me.txtIB(10)), 0, Me.txtIB(7), 0, Val(Me.txtNroInterno.Text), Me.dtpFecha(0).Value)
End If

End Sub


Private Sub Asiento()
If vConfigGral.vIncluyeContabilidad = True Then
        With frmAsientosAlta
            .txtCuentaVieneDe.Text = Me.Caption
            .txtCuentaVieneDe.Tag = txtProveedor(0).Tag
            .txtLeyenda.Text = vLeyendaAsiento
            .dtpFecha.Value = dtpFecha(0).Value
            .chkControlar.Value = xtpChecked
            
            If Not opTipoDocumento(7).Value = True Then
                .txtImporteVieneDe.Text = Trim(txtTotal.Text)
            Else
                .txtImporteVieneDe.Text = Trim(txtIB(10).Text)
            End If
            
            .cboTipoMovimiento.Tag = txtTipoMovimiento(0).Text
            .cboTipoMovimiento.Text = txtTipoMovimiento(1).Text
            
            .lblNroInterno.Caption = EsNulo(txtNroInterno.Text)
        
            .vVieneTabla = "PFactura"
            .vVieneIdNombre = "idPfactura"
            .vVieneIdValor = vIdPFactura
            
            
            ' ---------------- mas datos del asiento -----------
            .vCodigoCliente = Me.txtProveedor(0).Text
            .vCodigoProveedor = txtProveedor(0).Tag
            '----------------------------------------------------
            
            
        
            .Show
            .ZOrder (0)
            .SetFocus
        End With
    End If

End Sub



Private Sub cmdGuardarPago_Click()
On Error Resume Next

 If controlNroFactura Then
                MsgBox "Este documento ya fue grabado anteriormente.", vbCritical, "Documento duplicado..."
                Exit Sub
End If
             
             
    Dim vnrointerno2 As Long
    vnrointerno2 = txtNroInterno
    GuardarComprobante
    Me.txtNroInterno = vnrointerno2 + 1
    
If Err Then GrabarLog "cmdGuardarPago_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cmdMenuB_Click()
On Error Resume Next
    
    frmBuscarCompra.Show

If Err Then GrabarLog "cmdMenuB_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CtaCte()
On Error Resume Next

    With frmCtaCteP
        .Show
        .txtProveedor.Text = txtProveedor(0).Tag
        .txtProveedor_KeyPress (13)
    End With

If Err Then GrabarLog "cmdCtaCte_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdMenuN_Click()
On Error Resume Next

    RecargarForm
    vGrabaModo = 0

If Err Then GrabarLog "cmdMenuN_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdArticuloG_Click()
On Error Resume Next
    
Dim vvalores, vcampos, vcodigo, vsql As String

    With barticulo
      '  .Refresh
      '  .Recordset.Find ("descrip LIKE '%" + Trim(f(1).Text) + "%'")

       ' If .Recordset.EOF = True Then .Recordset.AddNew
        
        
       vcodigo = InputBox("Ingresar Código del nuevo artículo:", "Alta de Artículo...")
        
        vcampos = "codigo,descrip,pcosto"
        vvalores = "'" + vcodigo + "','" + Trim(f(1).Text) + "'," + Str(Val(f(2).Text))
        
        vsql = "insert into articulos (" + vcampos + ") values (" + vvalores + ")"
        
        Call EjecutarScript(vsql, pathDBMySQL)
        
        '.Recordset("codigo").Value = InputBox("Ingresar Código del nuevo artículo:", "Alta de Artículo...")
        '.Recordset("descrip").Value = Val(f(1).Text)
        '.Recordset("pcosto").Value = Val(f(2).Text)
        '.Recordset.Update
    
    
    
        MsgBox "Los datos del NUEVO! artículo fueron guardados correctamente", vbInformation, "Grabando ..."
        f(2).SetFocus
    
    
    End With
    
If Err Then GrabarLog "cmdArticuloG_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdMenuG_Click()
    On Error Resume Next
    
    Dim rsProveedores As New ADODB.Recordset, sqlProveedores As String
    
    sqlProveedores = "SELECT * FROM proveedores"
    
    With rsProveedores
        Call .Open(sqlProveedores, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        .AddNew
        
        .Fields("Codigo").Value = GenerarDato("SELECT MAX(codigo_num) as UltimoCodigo FROM Proveedores;", "UltimoCodigo")
        .Fields("Codigo_Num").Value = Val(.Fields("Codigo").Value)
        .Fields("Nombre").Value = txtProveedor(0).Text
        .Fields("Direccion").Value = txtProveedor(1).Text
        .Fields("Localidad").Value = txtProveedor(2).Text
        .Fields("Telefono").Value = txtProveedor(3).Text
        .Fields("Iva").Value = txtProveedor(4).Text
        .Fields("Cuit").Value = txtProveedor(5).Text
        
        .Update
    End With
    
    sqlProveedores = ""
    
    If rsProveedores.State = 1 Then
        rsProveedores.Close
        Set rsProveedores = Nothing
    End If
    
    If Err Then
        MsgBox "Hubo un problema al intentar guardar los datos, verifique que los datos sean válidos"
        GrabarLog "cmdMenuG_Click", Err.Number & " " & Err.Description, Me.Name
    Else
        MsgBox "Los datos del cliente fueron guardados"
    End If

End Sub
Private Sub ConfirmarDetalle()
    On Error Resume Next
    
    Dim rsPFDetalle As New ADODB.Recordset, sqlPFDetalle As String, i As Integer
    
    sqlPFDetalle = "SELECT * FROM PFDetalle WHERE (Remito = " & Val(vRemitoCompras) & ") ORDER BY idPFDetalle ASC"
    
    With rsPFDetalle
        .CursorLocation = adUseClient
        Call .Open(sqlPFDetalle, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If Not .EOF = True Then .MoveFirst

        For i = 1 To KlexDetalle.Rows - 1
        
            If Not Trim(KlexDetalle.TextMatrix(i, 1)) = "" Then
                .Filter = "idPFDetalle = " & Trim(KlexDetalle.TextMatrix(i, 1)) & ""
                    'Se Borro, Algo malo Paso
                    If .EOF = True Then .AddNew
            Else
                .AddNew
            End If
        
            .Fields("Remito").Value = Val(vRemitoCompras)
            .Fields("Fecha").Value = strfechaMySQL(dtpFecha(0).Value)
            .Fields("Codigo").Value = Replace(Replace(EsNulo(KlexDetalle.TextMatrix(i, 4)), "[", ""), "]", "")
            .Fields("Cantidad").Value = KlexDetalle.TextMatrix(i, 5)
            .Fields("Descripcion").Value = EsNulo(KlexDetalle.TextMatrix(i, 6))
            .Fields("Precio").Value = KlexDetalle.TextMatrix(i, 7)
            .Fields("Descuento").Value = Val(KlexDetalle.TextMatrix(i, 8))
            .Fields("TipoIva").Value = Val(KlexDetalle.TextMatrix(i, 9))
            .Fields("Total").Value = Val(KlexDetalle.TextMatrix(i, 11))
        
            'If venta.Text = "Contado" Then
                .Fields("Total_Cdo").Value = Val(KlexDetalle.TextMatrix(i, 11))
            'Else
                .Fields("Total_CtaCte").Value = Val(KlexDetalle.TextMatrix(i, 11))
            'End If
        
            .Fields("Confirmado").Value = "S"
            'rsPFDetalle.Fields("Ganancia").Value
            'rsPFDetalle.Fields("PVenta").Value
            'rsPFDetalle.Fields("PCosto").Value
            
            If Not KlexDetalle.TextMatrix(i, 21) = "S" Then
                .Update
                KlexDetalle.TextMatrix(i, 1) = .Fields("idPFDetalle").Value

                If Me.opTipoDocumento(8).Value Then
                    Call GuardarEnStock("Compras-Devolucion", EsNulo(.Fields("Codigo").Value), strfechaMySQL(dtpFecha(0).Value), Val(KlexDetalle.TextMatrix(i, 5)), "Devolucion de Mercaderia", KlexDetalle.TextMatrix(i, 1), 0)
                Else
                    If Not TipoDocumento = "Presupuesto" Then
                        If Not KlexDetalle.TextMatrix(i, 24) = "S" Then
                            Call GuardarEnStock("Compras-Nuevo", EsNulo(.Fields("Codigo").Value), strfechaMySQL(dtpFecha(0).Value), Val(KlexDetalle.TextMatrix(i, 5)), "Entrada de Mercaderia", 0, KlexDetalle.TextMatrix(i, 1))
                        Else
                            Call GuardarEnStock("Compras-Nuevo", EsNulo(KlexDetalle.TextMatrix(i, 4)), strfechaMySQL(dtpFecha(0).Value), Val(KlexDetalle.TextMatrix(i, 5)), "Actualizacion de Mercaderia", 0, KlexDetalle.TextMatrix(i, 1))
                        End If
                    End If
                End If
            Else
                Call GuardarEnStock("Compras-Modificar", EsNulo(.Fields("Codigo").Value), strfechaMySQL(dtpFecha(0).Value), Val(KlexDetalle.TextMatrix(i, 5)), "Entrada de Mercaderia", 0, KlexDetalle.TextMatrix(i, 1))
                .MoveNext
            End If
    
    
            ' actualizo el precio de costo y de venta - ale cambiar
          
          If UCase(LeerXml("ActualizaPrecio")) = "SI" Then
            Call actualizarPreciosArticulo(Val(KlexDetalle.TextMatrix(i, 12)), Val(KlexDetalle.TextMatrix(i, 13))) ' vidarticulos
          End If
    
        Next
    
    End With
    
    If Err Then
        MsgBox "Verifique si el documendo fue guardado correctamente.", vbCritical, "Mensaje..."
        GrabarLog "ConfirmarDetalle", Err.Number & " " & Err.Description + "Panic!!!", Me.Name
    End If

End Sub
Private Sub CondicionVenta(vCondicion As String)
    On Error Resume Next

    Select Case Trim(vCondicion)

        Case "Cuenta Corriente", ""
            WVenta 0

        Case "Cheques"
            WVenta (0)
            WVenta (2)

        Case "Contado"
            If (opTipoDocumento(3).Value = True) Or (opTipoDocumento(0).Value = True) Or (opTipoDocumento(2).Value = True) Then  ' fac. o n.d.
                'PCtaCte
                WVenta (0)
                'Caja
                WVenta (1)
            End If

    End Select
    
    If Err Then GrabarLog "CondicionVenta", Err.Number & " " & Err.Description, Me.Name
End Sub


Private Sub d_Click()

End Sub

Private Sub dgArticulos_DblClick()
On Error Resume Next

    With rsArticulos
        If Not .EOF = True And Not .BOF = True Then
            f(1).Text = .Fields("Codigo").Value
            
            vidArticulos = .Fields("idArticulos").Value
            
            Call f_KeyPress(1, 13)
            
        End If
    End With
    
If Err Then GrabarLog "DgArticulos_DblClick", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub dgArticulos_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        dgArticulos_DblClick
    End If

If Err Then GrabarLog "dgArticulos_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub dgProveedores_DblClick()
On Error Resume Next

    txtProveedor(0).Text = EsNulo(rsProveedores.Fields("Codigo").Value)
    Call txtProveedor_KeyPress(0, 13)
    dgProveedores.Visible = False

If Err Then GrabarLog "dgClientes_DblClick", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub dgProveedores_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        dgProveedores_DblClick
    End If

If Err Then GrabarLog "dgProveedores_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub dtpFecha_Click(Index As Integer)
On Error Resume Next



If Err Then GrabarLog "dtpFecha_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub dtpFecha_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        cboLetra.SetFocus
    End If

If Err Then GrabarLog "dtpFecha_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub dtpFecha_LostFocus(Index As Integer)

If dtpFecha(Index).Text = "" Then
    MsgBox "Cuidado. Debe ingresar una fecha"
    dtpFecha(Index).SetFocus
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, _
                       Shift As Integer)
    On Error Resume Next

    If KeyCode = vbKeyF1 Then
        Me.txtProveedor(0).Text = ""
        Me.txtProveedor(0).SetFocus
        'venta.SetFocus
    End If
    
    If KeyCode = vbKeyF2 Then
      Call cmdAcciones_Click(3)
    End If

    If KeyCode = vbKeyF4 Then
        
    End If

    If KeyCode = vbKeyF5 Then
        'Imprimir
        'NuevoCliente
    End If
    
    If KeyCode = vbKeyF11 Then
        txtProveedor(0).SetFocus
    End If
    
     If KeyCode = vbKeyF12 Then
        
        If Me.chkfijo.Value = xtpChecked Then
            Me.chkfijo.Value = xtpUnchecked
        Else
        
             Me.chkfijo.Value = xtpChecked
        End If
    
    End If

    'If KeyCode = vbKeyF10 Then
    '    If Not venta.ListCount - 1 = venta.ListIndex Then
    '        venta.ListIndex = venta.ListIndex + 1
    '    Else
    '        venta.ListIndex = 0
    '    End If
    'End If
    
    If Err Then GrabarLog "Form_KeyUp", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub MarkupLabel1_GotFocus()

End Sub



Private Sub nuevo_Click()

End Sub

Private Sub pbCarga_Click(Index As Integer)
On Error Resume Next

    vVuelveBusqueda = Me.Name
    vVieneBusqueda = pbCarga(Index).Tag

    Select Case Index
        
        Case 0 To 10
            frmBusqueda.Show
            
            Me.ZOrder (1)
            
    End Select
            
If Err Then GrabarLog "pbCarga_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub pbCarga_GotFocus(Index As Integer)
If Me.txtTipoMovimiento(0) = "CC" Or Me.txtTipoMovimiento(0) = "CD" Then
    pbCarga(0).Enabled = True
Else
    pbCarga(0).Enabled = False
End If
End Sub

Private Sub PushButton1_Click()
CalcularTotales
End Sub

Private Sub PushButton2_Click()


With frmBuscarFactura

    .viene = "grillaOrdenes"
    .txtCliente.Tag = Me.txtProveedor(0).Tag
    .txtCliente.Text = Me.txtProveedor(0).Tag
    Call .cmdFiltrar_Click
    .Show

End With

End Sub

Private Sub txtAuxiliares_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        
        Select Case Index
            
            Case 0, 1, 2
                txtAuxiliares(Index + 1).SetFocus
                    
            Case 3
            TabTotales.SelectedItem = 0
            txtSubtotal.SetFocus
        End Select
    End If
    
If Err Then GrabarLog "txtITC_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub txtIB_Change(Index As Integer)
On Error Resume Next
    
    Dim vCalculoIva As Double
    Select Case Index
    
        Case 0
            txtIB(2).Text = Val(Format(txtIB(0).Text * Val(txtIB(1).Text / 100), "#######0.00"))
            
        
        Case 2, 3, 4, 5, 6
            txtIB(10).Text = Val(txtIB(0).Text) + Val(txtIB(2).Text) - Val(txtIB(3).Text) + Val(txtIB(4).Text) + Val(txtIB(5).Text) + Val(txtIB(6).Text)
        
        Case 1
            If (Val(txtIB(1).Text) = 21) Or (Val(txtIB(1).Text) = 10.5) Or (Val(txtIB(1).Text) = 27) Then
                txtIB(2).Text = Val(Format(txtIB(0).Text * Val(txtIB(1).Text / 100), "#######0.00"))
            End If
    End Select
    

If Err Then GrabarLog "txtIB_Change", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtIB_GotFocus(Index As Integer)
On Error Resume Next

    With txtIB(Index)
        .SelStart = 0
        .SelLength = Len(txtIB(Index).Text)
    End With

If Err Then GrabarLog "txtIB_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtIB_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        txtIB(Index + 1).SetFocus
    End If

If Err Then GrabarLog "txtIB_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtIva_Change(Index As Integer)
On Error Resume Next
'    CalcularTotales
    'txtTotal.Text = Val(txtSubtotal.Text) + Val(txtIva(0).Text) + Val(txtIva(1).Text) + Val(txtIva(2).Text) - Val(txtDescuento.Text) + Val(txtAuxiliares(0).Text) + Val(txtAuxiliares(1).Text) + Val(txtAuxiliares(2).Text) + Val(txtAuxiliares(3).Text)

If Err Then GrabarLog "txtIva_Change", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtDescuento_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 13 Then
    
    Dim vauxi As Double
    
    vauxi = Val(txtSubtotal.Text) '+ Val(txtIva(0).Text) + Val(txtIva(1).Text) + Val(txtIva(2).Text)
    vauxi = (Val(txtDescuento.Text) * 100) / vauxi
    txtPorcentajeDescuento.Text = vauxi
    
    txtTotal.Text = Val(txtSubtotal.Text) + Val(txtIva(0).Text) + Val(txtIva(1).Text) + Val(txtIva(2).Text) - Val(txtDescuento.Text) + Val(txtAuxiliares(0).Text) + Val(txtAuxiliares(1).Text) + Val(txtAuxiliares(2).Text) + Val(txtAuxiliares(3).Text)

    Call CalcularTotales

    txtAuxiliares(0).SetFocus


End If



If Err Then
    If Err Then GrabarLog "txtPorcentajeDescuento_Change", Err.Number & " " & Err.Description, Me.Name
    Exit Sub
End If
End Sub
Private Sub txtNroComprobante_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        'txtNComprobante.Text = String(4 - Len(Mid(txtNComprobante.Text, 1, 4)), 0) & Mid(txtNComprobante.Text, 1, 4) & "-" & String(8 - Len(Mid(txtNComprobante.Text, 6, 8)), 0) & Mid(txtNComprobante.Text, 6, 8)
        'txtTipoMovimiento(0).SetFocus
       If UCase(LeerXml("Puesto")) = "PONS" Then
            f(1).SetFocus
        Else
            f(0).SetFocus
      End If
    End If
    
If Err Then GrabarLog "txtNComprobante_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub txtNroComprobante_LostFocus()
On Error Resume Next

    txtNroComprobante.Text = String(8 - Len(txtNroComprobante.Text), "0") & Val(txtNroComprobante.Text)

If Err Then GrabarLog "txtNroComprobante_LostFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtNroInterno_Change()
'MsgBox "cambia interno"
End Sub

Private Sub txtNroInterno_GotFocus()
On Error Resume Next

    With txtNroInterno
        .SelStart = 0
        .SelLength = Len(txtNroInterno.Text)
    End With

If Err Then GrabarLog "txtNroInterno_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtNroInterno_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        If vPFDetalle = True Then
            f(0).SetFocus
        Else
            txtIB(0).SetFocus
        End If
    
        txtIB(0).SetFocus
    End If


    
If Err Then GrabarLog "txtNroInterno_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub txtPorcentajeDescuento_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        If vPFDetalle = True Then
            'CalcularTotales
            txtDescuento.SetFocus
        Else
            txtDescuento.SetFocus
        End If
    
    
    
    
    
    Dim vauxi As Double
    
    vauxi = Val(txtSubtotal.Text) + Val(txtIva(0).Text) + Val(txtIva(1).Text) + Val(txtIva(2).Text)
    vauxi = (vauxi * Val(txtPorcentajeDescuento.Text) / 100)
    txtDescuento.Text = vauxi
    
    txtTotal.Text = Val(txtSubtotal.Text) + Val(txtIva(0).Text) + Val(txtIva(1).Text) + Val(txtIva(2).Text) - Val(txtDescuento.Text) + Val(txtAuxiliares(0).Text) + Val(txtAuxiliares(1).Text) + Val(txtAuxiliares(2).Text) + Val(txtAuxiliares(3).Text)

    Call CalcularTotales


    End If

If Err Then
    If Err Then GrabarLog "txtPorcentajeDescuento_Change", Err.Number & " " & Err.Description, Me.Name
    Exit Sub
End If

End Sub
Public Sub f_Change(Index As Integer)
    On Error Resume Next

    
    Dim descuento, impuesto As Double

    If Index = 1 Then
        
        
        If chkfijo Then Exit Sub

       ' Call fbuscarGrilla("Articulos", "Descrip", "codigo", Me.txt_Texbox2.Name, Me) ' ema:
    
        
        Call MostrarCoincidencias("Articulos", f(Index).Text)
        vOpenGrilla(1) = True
        If Not Me.opTipoDocumento(6) And Me.cboLetra = "A" Then f(4).Text = "21"
    Else
        
        descuento = Val(f(3)) * Val(f(2)) * Val(f(0)) / 100
        impuesto = Val(f(5)) * Val(f(2)) * Val(f(0)) / 100

        'If (Val(f(0)) * Val(f(2))) > 0 Then
            f(6) = Val(f(0)) * Val(f(2)) - descuento + impuesto
       ' End If

    End If
    
If Err Then GrabarLog "f_Change", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub f_GotFocus(Index As Integer)
    On Error Resume Next
    
    'With bdetalle
    '    If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
    '    .RecordSource = "SELECT * FROM pfdetalle WHERE (remito = " & Val(vRemitoCompras) & ") ORDER BY idPFDetalle ASC"
    '    .Refresh
    'End With
    
    'AlFinal

    If Err Then GrabarLog "f_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub f_KeyUp(Index As Integer, _
                      KeyCode As Integer, _
                      Shift As Integer)
    On Error Resume Next


    If Index = 1 Then
        If KeyCode = 38 Then
            With rsArticulos
                If Not .EOF = True And Not .BOF = True Then
                    .MovePrevious
                Else
                    .MoveLast
                End If
            End With
        End If

        If KeyCode = 40 Then
            With rsArticulos
                If Not .EOF = True And Not .BOF = True Then
                    .MoveNext
                Else
                    .MoveFirst
                End If
            End With
        End If
    
        If KeyCode = 13 And Not Trim(f(Index).Text) = "" Then
            dgArticulos_DblClick
        End If
    End If
    
If Err Then GrabarLog "f_KeyUp", Err.Number & " " & Err.Description, Me.Caption
End Sub
Public Sub f_KeyPress(Index As Integer, _
                      KeyAscii As Integer)
    On Error Resume Next

    If KeyAscii = 13 Then

        Select Case Index

            Case 1
                If Not vOpenGrilla(1) = True Then Pasar (Index)
                If Not Trim(f(1).Text) = "" Then
                    If BuscarArticulos = True Then
                        dgArticulos.Visible = Not True
                    End If
                    Pasar (Index)
                End If
            Case Else
                Pasar (Index)
        End Select

    End If

    
    If KeyAscii = 10 Then
        
        Select Case Index

            Case 1
                Pasar (Index)
                frmBuscarArticulo.busca = 1
                barticulo.Refresh
                'barticulo.Recordset.Sort = "descrip"
                barticulo.Recordset.Find ("descrip LIKE '" + Trim(f(1).Text) + "%'")
                
                vvvdescrip = f(1).Text  ' para ver q codigo usa
                MostrarDatos
                CargarBien
                
                f(1).Tag = barticulo.Recordset("codigo").Value
                
                'buscaart
                Pasar (Index)

            Case Else
                Pasar (Index)
        End Select

    End If


    If Err Then GrabarLog "f_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Load()
    On Error Resume Next

    LimpiarCampos

    With Me
        .Show
        .Top = 50
        .Left = 100
        .Width = 13725
        .Height = 9135
        .KeyPreview = True
       ' .TabTotales.Enabled = Not True
        .TabTotales.SelectedItem = 0
    End With
        
    opTipoDocumento(0).Value = True
    opTipoDocumento_Click (0)
    Me.txtProveedor(9).SetFocus
    
    dtpFecha(0).Value = Date
    dtpFecha(1).Value = Date
    
    PicDetalle.Visible = vPFDetalle
    
    ReDim vOpenGrilla(1) As Boolean
    
    If vPFDetalle = True Then
        opDetalle(0).Value = True
    Else
        opDetalle(1).Value = True
    End If
    
   
    'If Not vIdUsuarioNivel = 1 Then ControlarPermisos
   
    'Me.txtNroInterno = UltimoNroInterno2 + 1
    Call init
 
   
    CentrarFormulario Me
   
    If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub init()
    Me.txtProveedor(0).SetFocus
    Me.WindowState = 0
    articuloNuevo = False
    Me.vfechaPago = Date + 90
    
    If Me.opTipoDocumento(2) Then
       Call opTipoDocumento_Click(2)
    End If
End Sub
Private Sub GuardarCondicion()
    On Error Resume Next

    Select Case vGrabaModo

        Case 1 'Modificando una Factura
            
            Call BorrarBase("IvaFacturaCompra WHERE (remito = " & Val(vRemitoCompras) & ")", pathDBMySQL)

        Case Else 'Creo una Factura Nueva
            
            With bfactura

                If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
                .RecordSource = "SELECT * FROM pfactura WHERE 1=2"
                .Refresh
                .Recordset.AddNew
            End With
                
    End Select

    If Err Then
        MsgBox "Error! No se puede grabar este documento en el libro ", vbCritical, "Mensaje ..."
        GrabarLog "GuardarCondicion", Err.Number & " " & Err.Description, Me.Name
    End If

End Sub
Private Sub GuardarPFactura()
    On Error Resume Next
    Dim vcampo, vvalores As String
    Dim vPercepciones As Double

    GuardarCondicion 'Selecciona si es una actualización o ingreso
    
    With bfactura
    
        .Recordset("Fecha").Value = dtpFecha(0).Value
        .Recordset("FechaIVA").Value = vFechaIva.Value
        .Recordset("Hora").Value = Time
        .Recordset("Codigo").Value = EsNulo(txtProveedor(0).Tag)
        .Recordset("Nombre").Value = EsNulo(Trim(txtProveedor(0).Text))
        .Recordset("Domicilio").Value = EsNulo(Trim(txtProveedor(1).Text))
        .Recordset("Localidad").Value = EsNulo(txtProveedor(2).Text)
        .Recordset("Telefono").Value = EsNulo(txtProveedor(3).Text)
        .Recordset("Iva").Value = EsNulo(txtProveedor(4).Text)
        .Recordset("cuit").Value = EsNulo(txtProveedor(5).Text)
        
        
        .Recordset("remito").Value = Val(vRemitoCompras)
        
        If Not opTipoDocumento(7).Value = True Then
            .Recordset("subtotal").Value = Val(txtSubtotal.Text)
            .Recordset("descuento").Value = Val(txtDescuento.Text)
            .Recordset("Total").Value = Val(txtTotal.Text)
        Else
            If Not opTipoDocumento(7).Value = True Then
                .Recordset("subtotal").Value = Val(txtSubtotal.Text)
                .Recordset("descuento").Value = Val(txtDescuento.Text)
                .Recordset("Total").Value = Val(txtTotal.Text)
            Else
                .Recordset("subtotal").Value = Val(txtIB(0).Text)
                .Recordset("descuento").Value = 0
                .Recordset("Total").Value = Val(txtIB(10).Text)
            
            End If
        End If
        
        .Recordset("NroInterno").Value = Val(txtNroInterno.Text)
        
        

        
        .Recordset("Impreso").Value = 0
        .Recordset("Letra").Value = Trim(cboLetra.Text)
        .Recordset("PuntoDeVenta").Value = Trim(cboPuntoDeVenta.Text)
        '.Recordset("Sucursal").Value = Trim(cboPuntoDeVenta.Text)
        .Recordset("Ncomprobante").Value = Trim(txtNroComprobante.Text)
        
        .Recordset("TipoMovimiento").Value = txtTipoMovimiento(0).Text
        
        .Recordset("Tipo").Value = TipoDocumento
        .Recordset("FechaVencimiento").Value = dtpFecha(1).Value
        
        .Recordset("FechaVencimiento").Value = dtpFecha(1).Value
        
        .Recordset("fechaPago").Value = Me.vfechaPago
        
        ' --------------- Todo lo referido a persepciones y retenciones ----------------
        .Recordset("NoGravado").Value = CDbl(Me.vNoGravado)
        .Recordset("Exento").Value = Val(Me.vExento)
        .Recordset("Flete").Value = Val(Me.vflete)
        
        .Recordset("PorcentajeDescuento").Value = Me.txtPorcentajeDescuento
        
        
        '.Recordset("Retenciones").Value = Val(Me.txtAuxiliares(1)) ' no hay retenciones
        
        ' ----------- percepciones --------------------------
        .Recordset("PerIngBrutoStaFe").Value = Val(Me.vPerIngBrutoStaFe)
        .Recordset("IBBsAs").Value = Val(Me.vIBBsAs)
        .Recordset("IBOtros").Value = Val(Me.vIBOtros)
        .Recordset("PerIva").Value = Val(Me.vPerIva)
        .Recordset("PerImpGanancia").Value = Val(Me.vPerImpGanancia)
        '-------------------------------------------------------------------------------
        
        vPercepciones = Val(Me.vPerIngBrutoStaFe) + Val(Me.vPerIva) + Val(Me.vPerImpGanancia) + Val(Me.vIBBsAs) + Val(Me.vIBOtros)
        
        .Recordset.Update

        vIdPFactura = Val(.Recordset("idPFactura").Value)
    End With
    
    
    ' ------- variables para el insert en iva compra -----------
    vcampo = "(nrointerno, remito, iva105, iva210, iva270, NoGravado,Exento,PerIngBrutoStaFe,IBBsAs,IBOtros,PerIva,PerImpGanancia,Percepciones)"
    
    
    
    vvalores = "(" & txtNroInterno & "," & vRemitoCompras & ", " & Val(Me.txtIva(0)) & ", " & Val(Me.txtIva(1)) & ", " & Val(Me.txtIva(2)) & "," & Val(vNoGravado) & "," & Val(vExento) & "," & Val(vPerIngBrutoStaFe) & "," & Val(Me.vIBBsAs) & "," & Val(Me.vIBOtros) & "," & Val(vPerIva) & "," & Val(vPerImpGanancia) & "," + Str(vPercepciones) + ");"
    '----------------------------------------------------------
    
    If Me.opTipoDocumento(0) Or Me.opTipoDocumento(3) Or Me.opTipoDocumento(0) Then ' sin son documentos que llevan iva
        Call EjecutarScript("INSERT INTO IvaFacturaCompra " + vcampo + " values " + vvalores)
    End If
    
    If Me.opTipoDocumento(8).Value = True Then
       ' vvalores = "(" & txtNroInterno & "," & vRemitoCompras & ", " & Val(-1 * Me.txtIva(0)) & ", " & Val(-1 * Me.txtIva(1)) & ", " & Val(-1 * Me.txtIva(2)) & "," & Val(-1 * vNoGravado) & "," & Val(-1 * vExento) & "," & Val(-1 * vPerIngBrutoStaFe) & "," & Val(-1 * vPerIva) & "," & Val(-1 * vPerIva) & "," + Str(-1 * vPercepciones) + ");"
        Call EjecutarScript("INSERT INTO IvaFacturaCompra " + vcampo + " values " + vvalores)
    End If
    
    
    If Err < 0 Then
        MsgBox "Los datos del remito no son compatibles", vbCritical, "Mensaje ..."
        GrabarLog "GuardarPFactura", Err.Number & " " & Err.Description + "Panic!!!", Me.Name
    End If

End Sub
Private Sub Guardar() ' guarda en fdetalle y factura
    On Error Resume Next
        
    If vPFDetalle = True Then
        If Not opTipoDocumento(7).Value = True Then
            ConfirmarDetalle
            
         If LeerXml("ActualizaPrecio") Then
            Call actualizarPreciosArticulo(vidArticulos, Val(f(2).Text))
          End If
        
        End If
    End If
    

    Call GuardarPFactura
    Call CondicionVenta("")
    
    
     MsgBox "El Documento fue guadado.", vbInformation, "Mensaje"
     
    
    If Err Then GrabarLog "Guardar", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtIva_KeyPress(Index As Integer, _
                         KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Index = 0 Then
            If chkIva(1).Value = 1 Then
                txtIva(1).SetFocus
            Else
                If chkIva(2).Value = 1 Then
                    txtIva(2).SetFocus
                Else
                    TabTotales.SelectedItem = 1
                    txtPorcentajeDescuento.SetFocus
                End If
            End If
        End If
        
        If Index = 1 Then
            If chkIva(2).Value = 1 Then
                txtIva(2).SetFocus
            Else
                TabTotales.SelectedItem = 1
                txtPorcentajeDescuento.SetFocus
            End If
        End If
        
        If Index = 2 Then
            TabTotales.SelectedItem = 1
            txtPorcentajeDescuento.SetFocus
        End If
    End If

End Sub
Private Sub txtIva_LostFocus(Index As Integer)
On Error Resume Next
    
    txtIva(Index).Text = Format(txtIva(Index).Text, "#####0.00")

If Err Then GrabarLog "iva_LostFocus", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub LimpiarBase()
    On Error Resume Next
    
    KlexDetalle.Enabled = False
    
    limpiardetalle

    KlexDetalle.Enabled = True

    If Err Then GrabarLog "LimpiarBase", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Limpiar()
    Dim i As Integer
    
    On Error Resume Next
    
    For i = 0 To 6
        f(i).Text = ""
        f(i).Tag = ""
    Next

    f(0).SetFocus
    
    If Err Then GrabarLog "Limpiar", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub LimpiarCampos()
    Dim i As Integer
On Error Resume Next

    If Not CBool(LeerConfig(26)) = True Then
        For i = 0 To 6
            txtProveedor(i).Text = ""
            txtProveedor(i).Tag = ""
        Next
    End If
    
    UltimoRemito
    
    'txtNroInterno.Text = ""
    'UltimoNroInterno
    
    Me.vfechaPago = Date + 90
    
    txtSubtotal.Text = ""
    txtIva(0).Text = ""
    txtIva(1).Text = ""
    txtIva(2).Text = ""
    txtPorcentajeDescuento.Text = ""
    txtDescuento = ""
    txtAuxiliares(0).Text = ""
    txtAuxiliares(1).Text = ""
    txtAuxiliares(2).Text = ""
    txtAuxiliares(3).Text = ""
    txtTotal = ""
    
    txtNroComprobante.Text = ""
    cboPuntoDeVenta.Text = ""
    cboLetra.Text = ""
    
    vLeyendaAsiento = ""
    vTotalAsiento = 0
    cboBienesServicios.Text = ""
    txtTipoMovimiento(0).Text = ""
    txtTipoMovimiento(1).Text = ""
    Me.txtBancoCheque(0) = ""
    Me.txtBancoCheque(1) = ""
    
    vcomentario = txtObservacion
    Me.txtObservacion = ""
    
    
    ' --------------- Todo lo referido a persepciones y retenciones ----------------
    Me.vNoGravado = 0
    Me.vExento = 0
    Me.vflete = 0
        
        
    Me.txtPorcentajeDescuento = 0
                
 ' ----------- percepciones --------------------------
    Me.vPerIngBrutoStaFe = 0
    Me.vPerIva = 0
    Me.vPerImpGanancia = 0
    Me.vIBBsAs = 0
    Me.vIBOtros = 0
    
    
    vidArticulos = 0
    

    For i = 0 To 10
        If Not i = 8 And Not i = 9 Then
            txtIB(i).Text = ""
        End If
    Next
    

    For i = 0 To Val(txtCaja.Count - 2)
        txtCaja(i).Text = ""
    Next

    GBCaja.Visible = False
    
    FormatoGrillaDetalle (1)
    limpiardetalle
    
    If txtProveedor(0).Text = "" Then
        txtProveedor(0).SetFocus
    Else
        dtpFecha(0).SetFocus
    End If
    
    
   
    
If Err Then GrabarLog "LimpiarCampos", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub limpiardetalle()
    On Error Resume Next
    
    SinConfirmar
    
If Err Then GrabarLog "LimpiarDetalle", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub MostrarDatos()
    On Error Resume Next

    If txtProveedor(4).Text = "Responsable Inscripto" Or txtProveedor(4).Text = "Resp.Inscripto" Then
        vvdescrip = rsArticulosCompra.Fields("Descrip").Value
        vvcodigo = rsArticulosCompra.Fields("Codigo").Value
        vvpventa = rsArticulosCompra.Fields("PCosto").Value
        vvpdolar = Val(TraerDato("TipoMoneda", "idTipoMoneda = '" & rsArticulosCompra.Fields("idTipoMoneda").Value & "'", "Cotizacion"))
        CargarBien
    Else
        
        vvdescrip = TraerDato("Articulos", "codigo = '" & f(1).Tag & "'", "Descrip")
        vvcodigo = TraerDato("Articulos", "codigo = '" & f(1).Tag & "'", "Codigo")
        vvpventa = Val(TraerDato("Articulos", "codigo = '" & f(1).Tag & "'", "PCosto"))
        vvpdolar = Val(TraerDato("TipoMoneda", "idTipoMoneda = '" & f(1).Tag & "'", "Cotizacion"))
        CargarBien

    End If
    
    If Err Then GrabarLog "MostrarDatos", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub SinConfirmar()
    On Error Resume Next
    
    Call BorrarBase("PFDetalle WHERE (confirmado = 'N')", pathDBMySQL)

    If Err Then GrabarLog "SinConfirmar", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Pasar(Index As Integer)
On Error Resume Next

       If UCase(LeerXml("Puesto")) = "PONS" Then
                'If Trim(vConfigGral.vEmpresa) = Trim("wgestionpons") Then
               Select Case Index
                    Case 0
                        Index = 1
                    Case 1
                        Index = -1
                    Case Is > 5
                        GuardarRenglon
                        f(1).SetFocus
                
                End Select

                    f(Index + 1).SetFocus
    End If
    
    Call Pasar2(Index)
    
        'f(Index + 1).SetFocus
End Sub

Private Sub Pasar2(Index As Integer)
On Error Resume Next

    If Index >= 5 Then
       ' If Val(f(6).Text) <= 0 Then
       '     MsgBox "La cantidad y el precio deben ser valores positivos !", vbCritical, "Error..."
       '     Exit Sub
       ' End If

        GuardarRenglon

    Else
        
        f(Index + 1).SetFocus
    
    End If

    If Err Then GrabarLog "Pasar", Err.Number & " " & Err.Description, Me.Name
End Sub


Public Sub RecargarForm()
On Error Resume Next
    
    txtProveedor(0).SetFocus
    LimpiarCampos
    vGrabaModo = 0
    ZOrder (1)

If Err Then GrabarLog "RecargarForm", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub FormatoGrillaDetalle(vCantidadRenglones As Integer)
On Error Resume Next

    Dim i As Integer

    With KlexDetalle
        .FixedRows = 1
        .FixedCols = 1
    
        .Cols = 26
        .Rows = vCantidadRenglones + 1
        
        If vCantidadRenglones = 1 Then
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
                .ColWidth(i) = 0
            Next
        End If
        
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 250
        
        .TextMatrix(0, 1) = "idFDetalle"
        .ColWidth(1) = 0
        
        .TextMatrix(0, 2) = "Fecha"
        .ColWidth(2) = 0
        
        .TextMatrix(0, 3) = "Remito"
        .ColWidth(3) = 0
        
        .TextMatrix(0, 4) = "Codigo"
        .ColWidth(4) = 0
        .ColDisplayFormat(4) = ""
                
        .TextMatrix(0, 5) = "Cant."
        .ColWidth(5) = 800
        .ColDisplayFormat(5) = "#0.00"
        
        .TextMatrix(0, 6) = "Detalle"
        .ColWidth(6) = 7500
        
        .TextMatrix(0, 7) = "P. Costo"
        .ColWidth(7) = 900
        .ColDisplayFormat(7) = "#0.00"
        
        .TextMatrix(0, 8) = "% Desc."
        .ColWidth(8) = 900
        .ColDisplayFormat(8) = "#0.00"
                
        .TextMatrix(0, 9) = "% Iva"
        .ColWidth(9) = 900
        .ColDisplayFormat(9) = "#0.00"
        
        .TextMatrix(0, 10) = "% Imp."
        .ColWidth(10) = 900
        .ColDisplayFormat(10) = "#0.00"

        .TextMatrix(0, 11) = "$ Total"
        .ColWidth(11) = 1100
        .ColDisplayFormat(11) = "#0.00"

        .Col = 25
        .Row = .Rows - 1
        .CellBackColor = &HFFFCCC

        .Editable = True

        '.EnterKeyBehaviour = klexEKMoveDown
        .EnterKeyBehaviour = klexEKNone
        .BackColorAlternate = &HE0E0E0

    End With
    
If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub GuardarRenglon()
On Error Resume Next

    Dim i As Integer
    
    If Trim(vConfigGral.vempresa) = Trim("wgestionpons") Then
        If articuloNuevo Then Call cmdArticuloG_Click
    End If


    If f(4).Text = "" Then MsgBox "No ingresó el % de iva para este items"

  With KlexDetalle

        If .Rows <= 2 And .TextMatrix(.Rows - 1, 2) = "" Then
            FormatoGrillaDetalle (1)
        Else
            .Rows = .Rows + 1
        End If
        
        i = .Rows - 1
        
        .TextMatrix(i, 1) = ""
        .TextMatrix(i, 2) = EsNulo(vRemitoCompras)
        .TextMatrix(i, 3) = dtpFecha(0).Value
        .TextMatrix(i, 4) = "[" & EsNulo(vvcodigo) & "]"
        .TextMatrix(i, 5) = EsNulo(f(0).Text)
        .TextMatrix(i, 6) = EsNulo(f(1).Text)
        .TextMatrix(i, 7) = EsNulo(f(2).Text)
        .TextMatrix(i, 8) = EsNulo(f(3).Text)
        .TextMatrix(i, 9) = EsNulo(f(4).Text)
        .TextMatrix(i, 10) = EsNulo(f(5).Text)
        .TextMatrix(i, 11) = EsNulo(f(6).Text)
        
        .TextMatrix(i, 12) = vidArticulos 'idArticulos
        .TextMatrix(i, 13) = EsNulo(f(2).Text) 'pcosto aprobado por el ususario
    
    
    
        If Trim(vLeyendaAsiento) = "" Then
            vLeyendaAsiento = f(1).Text
        Else
            vLeyendaAsiento = vLeyendaAsiento & " - " & EsNulo(f(1).Text)
        End If
    
      '  Call LastKlexRow(Me.KlexDetalle)
    
        CalcularTotales
        
        chkfijo.Value = False
        
    End With
    
    
    vidArticulos = 0
    Limpiar

If Err Then GrabarLog "GuardarRenglon", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub opDetalle_Click(Index As Integer)
On Error Resume Next

    vPFDetalle = Not CBool(Index)
        
    'RecargarForm
    
If Err Then GrabarLog "opDetalle_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtAuxiliares_Change(Index As Integer)
On Error Resume Next

    txtTotal.Text = Val(txtSubtotal.Text) + Val(txtIva(0).Text) + Val(txtIva(1).Text) + Val(txtIva(2).Text) - Val(txtDescuento.Text) + Val(txtAuxiliares(0).Text) + Val(txtAuxiliares(1).Text) + Val(txtAuxiliares(2).Text) + Val(txtAuxiliares(3).Text)

If Err Then GrabarLog "txtAuxiliares_Change", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtAuxiliares_LostFocus(Index As Integer)
On Error Resume Next

    txtAuxiliares(Index).Text = Format(txtAuxiliares(Index).Text, "######0.00")

If Err Then GrabarLog "txtAuxiliares_LostFocus", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtSubtotal_change()
On Error Resume Next



    If chkIva(0).Value = 1 Then
        'txtIva(0).Text = Format(Val(txtSubtotal.Text) * chkIva(0).Tag / 100, "######0.00")
    Else
        txtIva(0).Text = ""
    End If
    
    If chkIva(1).Value = 1 Then
        'txtIva(1).Text = Format(Val(txtSubtotal.Text) * chkIva(1).Tag / 100, "######0.00")
    Else
        txtIva(1).Text = ""
    End If
    
    If chkIva(2).Value = 1 Then
        'txtIva(2).Text = Format(Val(txtSubtotal.Text) * chkIva(2).Tag / 100, "######0.00")
    Else
        txtIva(2).Text = ""
    End If
    
If Err Then GrabarLog "txtSubtotal_change", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtSubtotal_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        If chkIva(0).Value = 1 Then
            TabTotales.SelectedItem = 0
            txtIva(0).SetFocus
        Else
            If chkIva(1).Value = 1 Then
                txtIva(1).SetFocus
            Else
                If chkIva(2).Value = 1 Then
                    txtIva(2).SetFocus
                Else
                    TabTotales.SelectedItem = 1
                    txtPorcentajeDescuento.SetFocus
                End If
            End If
        End If
    End If
If Err Then Exit Sub
End Sub
Private Sub txtSubtotal_LostFocus()
On Error Resume Next
    
    txtSubtotal.Text = Format(txtSubtotal.Text, "#####0.00")

If Err Then GrabarLog "txtSubtotal_LostFocus", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub opTipoDocumento_Click(Index As Integer)
On Error Resume Next
    
    If Index = 0 Then KlexDetalle.BackColorFixed = &H8080FF
    If Index = 1 Then KlexDetalle.BackColorFixed = &HC0C000
    If Index = 5 Then KlexDetalle.BackColorFixed = &HC0C000
    If Index = 6 Then KlexDetalle.BackColorFixed = &HFF00&
    If Index = 7 Then KlexDetalle.BackColorFixed = &HFFFF&
    If Index = 3 Then KlexDetalle.BackColorFixed = &HC0F000
    If Index = 2 Then KlexDetalle.BackColorFixed = &HFF00FF
    
    If Index = 7 Then
        GBOtrosDocumentos.Visible = True
        GBOtrosDocumentos.Left = 0
        GBOtrosDocumentos.Top = 2280
        TabTotales.Visible = False
        fraCargaDetalle.Visible = False
    Else
        GBOtrosDocumentos.Visible = False
        TabTotales.Visible = True
        fraCargaDetalle.Visible = True

    End If
    
    If Index = 2 Then   ' documentos
        
        Me.cboLetra.Text = "X"
        
        Me.cboPuntoDeVenta = "1"
        
        Me.txtNroComprobante.Text = ultimoDocumento + 1
        
        Me.txtNroComprobante.SetFocus
    End If

If Err Then GrabarLog "opTipoDocumento_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Function ultimoDocumento() As Integer
On Error Resume Next

Dim vsql As String

vsql = " select   max(t.ncomprobante) As maxcomprobate from pfactura t Where t.tipo = 'Remito'"

ultimoDocumento = traerDatos2(vsql, "maxcomprobate", pathDBMySQL)

If Err Then

    ultimoDocumento = 0
    Exit Function

End If

End Function

Private Sub AlFinal()
    On Error Resume Next
    
    'With bdetalle
    
    '    If Not .Recordset.RecordCount = 0 Then .Recordset.MoveLast
    
    'End With
    
    If Err Then GrabarLog "AlFinal", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtTipoMovimiento_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        Select Case Index
        
            Case 0
                txtTipoMovimiento(0).Text = UCase(txtTipoMovimiento(0).Text)
                txtTipoMovimiento(1).Text = TraerDato("TipoMovimientos", "Codigo = '" & Trim(txtTipoMovimiento(0).Text) & "'", "TipoMovimiento")
                
                
            Case 1
                txtTipoMovimiento(0).Text = UCase(TraerDato("TipoMovimientos", "TipoMovimiento = '" & Trim(txtTipoMovimiento(1).Text) & "'", "Codigo"))
        
        End Select
    
    
        txtNroInterno.SetFocus
    

    End If


If Err Then GrabarLog "txtTipoMovimiento_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtTipoMovimiento_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
    If KeyCode = vbKeyF3 Then
        pbCarga_Click (0)
    End If

If Err Then GrabarLog "txtTipoMovimiento_KeyUp", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtCaja_GotFocus(Index As Integer)
On Error Resume Next

    txtCaja(Index).SelStart = 0
    txtCaja(Index).SelLength = Len(txtCaja(Index).Text)

If Err Then GrabarLog "txtCaja_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtCaja_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
    
        txtCaja(Index).Text = UCase(txtCaja(Index).Text)
        
        Select Case Index
        
            Case 0
                txtCaja(Index + 1).Text = TraerDato("Bancos", "idBancos = '" & Trim(txtCaja(Index).Text) & "'", "Descripcion")
                txtCaja(Index + 2).SetFocus
            
            Case 2
                txtCaja(Index + 1).Text = TraerDato("BancosCuentas", "idBancosCuentas = " & Trim(txtCaja(Index).Text) & "", "Cuenta")
                txtCaja(Index + 2).SetFocus
                
            
            Case 4
                txtCaja(Index + 1).Text = TraerDato("TipoValor", "idTipoValor = '" & Trim(txtCaja(Index).Text) & "'", "TipoValor")
                
                If Not Trim(txtCaja(Index + 1).Text) = "" Then
                    If Not UCase(Trim(txtCaja(Index + 1).Text)) = "CH" Then
                        txtCaja(7).Text = ""
                        txtCaja(7).SetFocus
                    Else
                        txtCaja(8).SetFocus
                    End If
                Else
                    txtCaja(Index).Text = ""
                    txtCaja(Index + 1).Text = ""
                End If
                
            Case 6, 7
                txtCaja(Index + 1).SetFocus
            
            Case 8
                cmdGuardarPago.SetFocus
        
        End Select
    
    
        'If txtCaja(Index).Text = "" Then txtIB(0).SetFocus
    
    End If

If Err Then GrabarLog "txtCaja_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub txtCaja_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next

    If KeyCode = vbKeyF3 Then
        Select Case Index
        
            Case 0
                pbCarga_Click (2)
            Case 1
                
            Case 2
                pbCarga_Click (3)
            
            Case 3
                
            
            Case 4
                pbCarga_Click (4)
            Case 5
                
            
            Case 6
        
        End Select
        
        
        
    End If

If Err Then GrabarLog "txtCaja_KeyUp", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtTotal_Change()
On Error Resume Next

    'txtTotal.Text = Format(txtTotal.Text, "#######0.00")
    
If Err Then GrabarLog "txtTotal_Change", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtTotal_GotFocus()
On Error Resume Next
    
'    txtTotal.Text = Val(txtSubTotal.Text) + Val(txtIva(0).Text) + Val(txtIva(1).Text) + Val(txtIva(2).Text) - Val(txtDescuento.Text) + Val(txtAuxiliares(0).Text) + Val(txtAuxiliares.Text)

If Err Then GrabarLog "txtTotal_Change", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtTotal_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        'cmdAcciones(1).SetFocus
    End If

End Sub
Private Sub tprecio_Click()
    f(0).SetFocus
End Sub
Private Sub UltimoRemito()
    On Error Resume Next
    
    vRemitoCompras = NroRemitoNuevo

Exit Sub
    Dim rsRemito As New ADODB.Recordset, sqlRemito As String
    
    sqlRemito = "SELECT MAX(Remito) as UltimoRemito FROM PFactura"
    
    With rsRemito
        Call .Open(sqlRemito, ConnDDBB, adOpenStatic, adLockReadOnly)

        If Not .EOF = True Then
            vRemitoCompras = Val(EsNulo(.Fields("UltimoRemito").Value)) + 1
        Else
            vRemitoCompras = 1
        End If
        
        'Set .Recordset = Nothing
    End With
    
    sqlRemito = ""
    
    If rsRemito.State = 1 Then
        rsRemito.Close
        Set rsRemito = Nothing
    End If
    
    If Err Then GrabarLog "UltimoRemito", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub txtTotal_LostFocus()
On Error Resume Next
    
    txtTotal.Text = Format(txtTotal.Text, "######0.00")

If Err Then GrabarLog "txtTotal_LostFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtProveedor_Change(Index As Integer)
On Error Resume Next

    If Index = 0 And Not vGrabaModo = 1 Then
        If Not vOpenGrilla(0) = True Then
            Call MostrarCoincidencias("Proveedores", Trim(txtProveedor(0).Text))
        End If
    End If
    
If Err Then GrabarLog "txtProveedor_Change", Err.Number & " " & Err.Description, Me.Name
End Sub

Public Sub txtProveedor_KeyPress(Index As Integer, _
                      KeyAscii As Integer)
    On Error Resume Next
    
    If KeyAscii = 13 Then
        If Index = 6 Then
            If vPFDetalle = True Then
                f(0).SetFocus
            Else
                txtSubtotal.SetFocus
            End If
        
            Exit Sub
        End If
       
        If Index = 0 Then
            
            txtProveedor(0).Text = EsNulo(rsProveedores.Fields("Codigo").Value)
            
            If BuscarProveedor = True Then
                dtpFecha(0).SetFocus
            Else
            
            frmProveedoresAlta.viente = Me.Name
            frmProveedoresAlta.txtAlta(1) = Me.txtProveedor(0).Text
            frmProveedoresAlta.Show
            
            End If
            
        Else
        
            If Index >= 5 Then
                f(0).SetFocus
            Else

                If Index > 0 Then Index = 6
                txtProveedor(Index + 1).SetFocus
            End If
        End If

    End If

    If Err Then GrabarLog "txtProveedor_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Function TipoDocumento() As String
    On Error Resume Next
    
    Dim i As Integer, j As Integer

    For i = 0 To opTipoDocumento.Count - 1

        If opTipoDocumento(i).Value = True Then

            Select Case i

                Case 0
                    TipoDocumento = "Fact " & Trim(cboLetra.Text)
                    Exit For
                Case 1
                    TipoDocumento = "Exento"
                    Exit For
                
                Case 2
                    TipoDocumento = "Remito"
                    Exit For
                
                Case 3
                    TipoDocumento = "Nota D"
                    Exit For
                Case 4
                    TipoDocumento = "Nota C"
                    Exit For
                Case 5
                    TipoDocumento = "Presupuesto"
                    Exit For
            
                Case 6
                    TipoDocumento = "Documento"
                    Exit For
                Case 8
                    TipoDocumento = "Nota C"
                Case 7
                    Select Case UCase(txtTipoMovimiento(0).Text)
            
                        Case "RI", "RV", "CD", "CC", "FC", "RG", "SU", "SI", "AD", "AC"
                            TipoDocumento = "Fact " & Trim(cboLetra.Text)
                
                        Case "NC"
                            TipoDocumento = "Nota C"
                
                        Case "ND"
                            TipoDocumento = "Nota D"
                
                        Case "RC"
                            TipoDocumento = "Recibo"
                
                    End Select
            End Select

        End If

    Next
    
    If Err Then GrabarLog "TipoDocumento", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub WVenta(vtipo As Byte)
    On Error Resume Next
    
    Dim rsVenta As New ADODB.Recordset, sqlVenta As String, i As Integer

    With rsVenta
        Select Case vtipo
        
            Case 0
                sqlVenta = "SELECT * FROM PCuentascorrientes WHERE 1=1"

        
        End Select
        
        Call .Open(sqlVenta, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not vGrabaModo = 1 Then
            .AddNew
        Else
            .Find ("remito=" & vRemitoCompras)

            If .EOF Then Exit Sub
        End If
    
        If Trim(txtTipoMovimiento(0).Text) = "" Then txtTipoMovimiento(0).Text = "FC" ' para que no tome el credito
    
    
        Select Case vtipo
    
            Case 0
                .Fields("TipoMovimiento").Value = UCase(Trim(txtTipoMovimiento(0).Text))
                .Fields("Fecha").Value = strfechaMySQL(dtpFecha(0).Value)
                .Fields("Codigo").Value = EsNulo(txtProveedor(0).Tag)
                .Fields("Nombre").Value = EsNulo(txtProveedor(0).Text)
                .Fields("Remito").Value = vRemitoCompras
                
                
                ' --- comentarios --------------------
                If opTipoDocumento(0).Value = True Then
                    .Fields("comentario").Value = Left("Fact. " & Trim(cboLetra.Text) & " Nº " & txtNroComprobante.Text, 100)
                Else
                    .Fields("comentario").Value = Left("Docume " & TipoDocumento & " Nº " & txtNroComprobante.Text, 100)
                End If
                
                
                
                '----
                If opTipoDocumento(0) = True Or opTipoDocumento(6) = True Or opTipoDocumento(3) = True Then  ' factura y documentos
                    .Fields("Credito").Value = 0
                    .Fields("Debito").Value = Val(txtTotal.Text)
                End If
                
                
                If opTipoDocumento(8) = True Then  ' nota de credito
                    .Fields("Credito").Value = Val(txtTotal.Text)
                    .Fields("Debito").Value = 0
                End If
                
                

                .Fields("NroInterno").Value = Val(txtNroInterno.Text)
                .Fields("comentario").Value = Left(.Fields("comentario").Value + " - " + Me.txtIB(7).Text, 100)
                
                '.Fields("NroAsiento").Value = Val(vnroasiento)
                
                
                .Update

        End Select
        

    Select Case txtTipoMovimiento(0).Text
        
        Case "CD"
            Call EjecutarScript("INSERT INTO PCuentasCorrientes (Fecha,Codigo,Nombre,Debito,Comentario,Remito,NroInterno,TipoMovimiento, idMedioPago) VALUES ('" & strfechaMySQL(dtpFecha(0).Value) & "','" & txtProveedor(0).Tag & "','" & txtProveedor(0).Text & "'," & Val(txtIB(10).Text) & ",'" & Trim(.Fields("Comentario").Value) & "'," & vRemitoCompras & "," & Val(txtNroInterno.Text) & ",'CC', 11)")
            Call EjecutarScript("INSERT INTO BancosMovimientos (idBancos,idBancosCuentas,Fecha,Credito,Comentario,NroCheque,TipoMovimiento,NroInterno,idTipoValor) VALUES ('" & Trim(txtCaja(0).Text) & "'," & Val(txtCaja(2).Text) & ",'" & strfechaMySQL(dtpFecha(0).Value) & "'," & Val(txtCaja(7).Text) & ",'" & Trim(txtCaja(8).Text) & "'," & Val(txtCaja(6).Text) & ",'CC'," & Val(txtNroInterno.Text) & ",'" & Trim(txtCaja(4).Text) & "')")
        
       ' Case "AD"
           
        '    Call EjecutarScript("INSERT INTO pCuentasCorrientes (Fecha,Codigo,Nombre,Debito,Comentario,Remito,NroInterno,TipoMovimiento) VALUES ('" & strfechaMySQL(dtpFecha(0).Value) & "','" & txtProveedor(0).Tag & "','" & txtProveedor(0).Text & "'," & Val(.Fields("Debito").Value) & ",'" & Trim(.Fields("Comentario").Value) & "'," & vRemitoCompras & "," & Val(txtNroInterno.Text) & ",'AD')")
            
        Case "SU", "SI", "RG", "IB", "RV"
           'Call EjecutarScript("INSERT INTO CuentasCorrientes (Fecha,Codigo,Nombre,Credito,Comentario,Remito,NroInterno,TipoMovimiento) VALUES ('" & strfechaMySQL(dtpFecha.value) & "','" & txtProveedor(0).Tag & "','" & txtProveedor(0).Text & "'," & Val(.Fields("Debito").value) & ",'" & Trim(.Fields("Comentario").value) & "'," & vNroRemito & "," & Val(txtNroInterno.Text) & ",'CC')")
        
        Case "FC"
            'No pasa NADA
        Case Else
               'MsgBox "OJO"
       ' Case "CC"
            ' Panic: ver que es lo que tengo qu hacer aca
    End Select
    
        End With
    
    sqlVenta = ""

    If rsVenta.State = 1 Then
        rsVenta.Close
        Set rsVenta = Nothing
    End If
    
    If Err < 0 Then
        MsgBox "Revise si este movimiento fue realizado correctamente", vbCritical, "Cuidado ..."
        GrabarLog "WVenta", Err.Number & "" & Err.Description + "Panic !!!", Me.Name
    End If
End Sub
Private Sub ControlarPermisos()
On Error Resume Next

    Dim rsUsuariosPermisos As New ADODB.Recordset, sqlUsuariosPermisos As String
    
    'sqlUsuariosPermisos = "SELECT * FROM UsuariosPermisos WHERE (idUsuarios = " & vIdUsuario & ") AND (Formulario = 'frmCompras')  AND (NOT Accion IS NULL OR NOT Accion = '')"
   
    With rsUsuariosPermisos
        Call .Open(sqlUsuariosPermisos, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then .MoveFirst
        
        Do Until .EOF = True
    
            Select Case .Fields("Accion").Value
                
                Case "Nuevo"
                    cmdMenuN.Enabled = CBool(.Fields("Habilitado").Value)
                    
                
                Case "Borrar"
                    'cmdAcciones(2).Enabled = CBool(.Fields("Habilitado").Value)
                    
                    
                Case "Guardar"
                    'cmdAcciones(1).Enabled = CBool(.Fields("Habilitado").Value)
                    cmdMenuG.Enabled = CBool(.Fields("Habilitado").Value)
                    
                Case "Imprimir"
                    'cmdImprimir(0).Enabled = CBool(.Fields("Habilitado").Value)
                    'cmdImprimir(1).Enabled = CBool(.Fields("Habilitado").Value)
                
                Case "Buscar"
                    cmdMenuB.Enabled = CBool(.Fields("Habilitado").Value)
                    
                Case "Modificar"
                    'cmdAcciones.Enabled = CBool(.Fields("Habilitado").Value)
                
                    
            End Select

            .MoveNext
        Loop
    
    End With
    
    sqlUsuariosPermisos = ""
    
    If rsUsuariosPermisos.State = 1 Then
        rsUsuariosPermisos.Close
        Set rsUsuariosPermisos = Nothing
    End If

If Err Then GrabarLog "ControlarPermisos", Left(Err.Number & " " & Err.Description, 99), Me.Name
End Sub

Private Sub MostrarCoincidencias(vTipoBusqueda As String, vBusqueda As String)
On Error Resume Next

    If vTipoBusqueda = "Articulos" Then
        
        Dim sqlArticulos As String, sqlTipoDetalle As String
    
        
        Set rsArticulos = New ADODB.Recordset
    
        If Trim(f(1).Text) = "" Then
            sqlArticulos = "SELECT * FROM Articulos WHERE 1=2"
        Else
        
            If Val(f(1).Text) > 0 Then
                sqlArticulos = "SELECT * FROM Articulos WHERE (Codigo like '%" & Trim(vBusqueda) & "') "
            
            Else
        
                sqlArticulos = "SELECT * FROM Articulos WHERE (Codigo LIKE '%" & Trim(vBusqueda) & "%') OR (Descrip LIKE '%" & Trim(vBusqueda) & "%')"
        
            End If
        End If
    
        With rsArticulos
            If .State = 1 Then .Close
    
            .CursorLocation = adUseClient
        
            Call .Open(sqlArticulos, ConnDDBB, adOpenStatic, adLockReadOnly)
        
            dgArticulos.Visible = Not .EOF
        
            If Not .EOF = True Then
                Set dgArticulos.DataSource = rsArticulos
                Call FormatoGrilla("Articulos")
            Else
                Set dgArticulos.DataSource = Nothing
            End If
        
        End With
    
        sqlArticulos = ""

    Else
        Dim sqlProveedores As String
    
        Set rsProveedores = New ADODB.Recordset
    
        If Trim(vBusqueda) = "" Then
            sqlProveedores = "SELECT * FROM Proveedores WHERE 1=2"
        Else
            sqlProveedores = "SELECT * FROM Proveedores WHERE (Codigo LIKE '%" & Trim(vBusqueda) & "%') OR (Nombre LIKE '%" & Trim(vBusqueda) & "%')"
        End If
    
        With rsProveedores
            If .State = 1 Then .Close
    
            .CursorLocation = adUseClient
        
            Call .Open(sqlProveedores, ConnDDBB, adOpenStatic, adLockReadOnly)
        
            dgProveedores.Visible = Not .EOF
        
            If Not .EOF = True Then
                Set dgProveedores.DataSource = rsProveedores
                Call FormatoGrilla(vTipoBusqueda)
            Else
                Set dgProveedores.DataSource = Nothing
            End If
        
        End With
    
        sqlProveedores = ""
    
    End If

If Err Then GrabarLog "MostrarCoincidencias", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub FormatoGrilla(vtipo As String)
On Error Resume Next
    
    Dim i As Integer
    
    If vtipo = "Articulos" Then
    
        With dgArticulos
        
           
            'Lo Paso al Frente
            .ZOrder (0)
        
            'Lo Ubico justo debajo de donde escribo
            .Top = fraCargaDetalle.Top - 50
            '1.Left = f(1).Left + 500
            '.Width = 10485
      
            .HeadLines = 1.2
        
            For i = 0 To .Columns.Count - 1
        
                        Select Case i
                    Case 1
                        .Columns(i).Width = 3000
                    Case 4
                        .Columns(i).Width = 6000
                    Case 9
                        .Columns(i).Width = 1000
                    Case 23
                        .Columns(i).Width = 1000
                    Case Else
                        .Columns(i).Width = 0
                
                End Select
            Next

        End With
    
    Else
        With dgProveedores
        
            .ZOrder (0)
            '.Top = 485
            '.Left = 1650

            .HeadLines = 1.2
        
            For i = 0 To .Columns.Count - 1
        
                Select Case i
        
                    Case 3
                        .Columns("Nombre").Width = .Width - 100
                    
                    Case Else
                        .Columns(i).Width = 0
                
                End Select
            Next

        End With
    
    
    End If
    
If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub txtProveedor_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next

    If Index = 0 Then
        If KeyCode = 38 Then
            With rsProveedores
                If Not .EOF = True And Not .BOF = True Then
                    .MovePrevious
                Else
                    .MoveLast
                End If
            End With
        End If

        If KeyCode = 40 Then
            With rsProveedores
                If Not .EOF = True And Not .BOF = True Then
                    .MoveNext
                Else
                    .MoveFirst
                End If
            End With
        End If
    End If
    
    If KeyCode = 13 Then
        dgProveedores_DblClick
    End If
    
If Err Then GrabarLog "v_KeyUp", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtProveedor_LostFocus(Index As Integer)
On Error Resume Next
    
    Select Case Index

        Case 0
            vOpenGrilla(0) = False
            dgProveedores.Visible = Not True
        Case 1
        
        Case 2
        
        Case 3
        
        Case 4
        
        Case 5
    
    End Select

If Err Then GrabarLog "txtProveedor_LostFocus", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub vExento_Change()
CalcularTotales
End Sub

Private Sub vflete_Change()
CalcularTotales
End Sub

Private Sub vIBBsAs_Change()
CalcularTotales
End Sub

Private Sub vIBOtros_Change()
CalcularTotales
End Sub

Private Sub vNoGravado_Change()
CalcularTotales
End Sub

Private Sub vPerImpGanancia_Change()
CalcularTotales
End Sub

Private Sub vPerIngBrutoStaFe_Change()
CalcularTotales
End Sub

Private Sub vPerIva_Change()
CalcularTotales
End Sub



Private Sub Imprimir(vnroremito As Long, vTipoDocumento As String)
    On Error Resume Next

         Dim i, t As Integer
        
        'bdetalle.RecordSource = "select * from pfdetalle where remito=" + Str(vnroremito)
        'bdetalle.Refresh
        
       ' t = bdetalle.Recordset.RecordCount
        
        t = traerDatos2("select count(remito) as c from pfdetalle where remito = " + Str(vnroremito), "c", pathDBMySQL)
        'bdetalle.Refresh

        Dim vvsql As String
        
        
        vvsql = "delete from Relleno where remito=" + Str(vnroremito) '+ " where IdRelleno=1"
        Call EjecutarScript(vvsql, pathDBMySQL)
        
        vvsql = "insert into Relleno (remito) values (" + Str(vnroremito) + ")" ' " where IdRelleno=1"
       
        Call EjecutarScript(vvsql, pathDBMySQL)

        If Not LeerXml("MostrarSaldo") = "SI" Then margenfactura = (30 - t) * 208

        With Mantenimiento.rscfact
            If Not .State = 0 Then .Close
        
          .Source = "SHAPE {SELECT * FROM Factura WHERE remito = " & Str(vnroremito) & "}  AS cfact APPEND ((SHAPE {SELECT * FROM relleno}  AS crelleno APPEND ({SELECT FDetalle.*,relleno.* FROM relleno,FDetalle WHERE (fdetalle.remito = relleno.remito) AND (fdetalle.remito =" & Str(vnroremito) & ") ORDER BY idFDetalle ASC} AS cdetalle RELATE 'remito' TO  PARAMETER 0) AS cdetalle) AS crelleno RELATE 'Remito' TO 'remito') AS crelleno"

         
         '.Source = "SHAPE {SELECT * FROM Factura}  AS cfact APPEND (( SHAPE {SELECT * FROM `relleno`}  AS crelleno APPEND ({SELECT fdetalle.*,relleno.* FROM relleno,fdetalle WHERE fdetalle.remito = relleno.remito}  AS cdetalle RELATE 'remito' TO 'Remito') AS cdetalle) AS crelleno RELATE 'Remito' TO 'remito') AS crelleno"
            If Not .State = 1 Then .Open
            .Close
            .Open
        End With
        
        
          Unload Mantenimiento
          Load Mantenimiento
    
        'Me.WindowState = vbMinimized
        
        Select Case vTipoDocumento
    

            Case "Presupuesto"
                'ipresupuesto.Show
                
                mostrar_documentos
                'idocumento.Show

            Case "Remito"
                 mostrar_documentos
    
            Case "Documento"
               
                    mostrar_documentos
        End Select
    
        MousePointer = vbDefault
    
        'Me.WindowState = 1


 
   ' NuevoCliente

If Err Then GrabarLog "Imprimir", Err.Number & " " & Err.Description, Me.Name
End Sub


Private Sub mostrar_documentos()

With idocumento
'----------- titulos -------
.Sections("titulos").Controls("enroremito").Caption = Str(vnrocomprobante)
'.Sections("titulos").Controls("ecventa").Caption = Me.vcventa

.Sections("titulos").Controls("enombre").Caption = txtProveedor(0).Text
.Sections("titulos").Controls("edomicilio").Caption = txtProveedor(1).Text
.Sections("titulos").Controls("elocalidad").Caption = txtProveedor(2).Text
.Sections("titulos").Controls("ecuit").Caption = txtProveedor(3).Text
.Sections("titulos").Controls("efecha").Caption = Str(Me.dtpFecha(0))

'---------------------------

.Sections("totales").Controls("etotal").Caption = Format(vgTtotal, "#,###,##0.00")
.Sections("totales").Controls("esubtotal").Caption = Format(vgTsubtotal, "#,###,##0.00")
.Sections("Totales").Controls("edescuento").Caption = Format(vgTPdescuento, "#,###,##0.00")
.Sections("Totales").Controls("txtObservacion").Caption = vcomentario

.Show
End With
End Sub

Function ValirParaGuardar() As Boolean

ValirParaGuardar = True
Dim vmerror As String

vmerror = ""
If Me.dtpFecha(0).Text = "" Then
    ValirParaGuardar = False
    vmerror = vmerror + " - Fecha de emision"
End If

If Me.dtpFecha(1).Text = "" Then
    ValirParaGuardar = False
    vmerror = vmerror + " - Fecha de emision"
End If

If Me.opTipoDocumento(0).Value And _
(Val(Me.txtIva(0).Text) + Val(Me.txtIva(1).Text) + Val(Me.txtIva(2).Text)) = 0 And _
Me.txtProveedor(4).Text = "Iva Responsable Inscripto" Then

    
    If MsgBox("Quire hacer la factura sin IVA ?", vbYesNo) = vbYes Then
        ValirParaGuardar = True
    Else
        vmerror = vmerror + " - Debe ingresar IVA"
        ValirParaGuardar = False
    End If
    
   

End If

If valNroComprobante() = False Then

    ValirParaGuardar = False
    vmerror = vmerror + " - Este comprobante ya fue cargado."

End If



If Not ValirParaGuardar Then MsgBox vmerror

End Function


Function valNroComprobante() As Boolean
valNroComprobante = True

Dim Val1 As Boolean

Dim vsql, valor As String


vsql = "select count(*) as c from pfactura where " + _
"codigo = '" + Trim(Me.txtProveedor(0).Tag) + "'" + " and " + _
"Letra = '" + Trim(Me.cboLetra) + "' and " + _
"convert(PuntoDeVenta , unsigned)  = " + Str(Val(Me.cboPuntoDeVenta)) + " and " + _
"convert(NComprobante , unsigned) = " + Str(Val(Me.txtNroComprobante)) + ""


valor = traerDatos2(vsql, "c", pathDBMySQL)

If Val(valor) > 0 Then
    valNroComprobante = False
End If


End Function

