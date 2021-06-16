VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "KlexGrid.ocx"
Object = "{9746E3DA-06E1-4D26-9CE4-D9F6411A9C70}#1.0#0"; "SMGA_OcxTxt2009.ocx"
Begin VB.Form frmBancosMovimientos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos de Banco"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   240
   ClientWidth     =   13695
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   1.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBancoMovimientos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   13695
   Begin XtremeSuiteControls.GroupBox GroupBox4 
      Height          =   525
      Left            =   30
      TabIndex        =   49
      Top             =   0
      Width           =   13665
      _Version        =   851968
      _ExtentX        =   24104
      _ExtentY        =   926
      _StockProps     =   79
      Caption         =   "GroupBox4"
      Appearance      =   1
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton PBAcciones 
         Height          =   375
         Index           =   1
         Left            =   12360
         TabIndex        =   50
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cerrar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Picture         =   "frmBancoMovimientos.frx":000C
      End
      Begin XtremeSuiteControls.PushButton PBAcciones 
         Height          =   375
         Index           =   0
         Left            =   6330
         TabIndex        =   51
         Top             =   120
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Imprimir"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Picture         =   "frmBancoMovimientos.frx":040C
      End
      Begin XtremeSuiteControls.PushButton PBAcciones 
         Height          =   375
         Index           =   2
         Left            =   5100
         TabIndex        =   52
         Top             =   120
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Buscar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Picture         =   "frmBancoMovimientos.frx":0827
      End
      Begin XtremeSuiteControls.PushButton PBAcciones 
         Height          =   375
         Index           =   3
         Left            =   90
         TabIndex        =   53
         Top             =   120
         Width           =   1845
         _Version        =   851968
         _ExtentX        =   3254
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Agregar Movimiento"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Picture         =   "frmBancoMovimientos.frx":0C5E
      End
      Begin XtremeSuiteControls.PushButton PBAcciones 
         Height          =   375
         Index           =   4
         Left            =   2160
         TabIndex        =   54
         Top             =   120
         Width           =   1425
         _Version        =   851968
         _ExtentX        =   2514
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Borrar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
         Picture         =   "frmBancoMovimientos.frx":11F8
      End
      Begin XtremeSuiteControls.PushButton PBAcciones 
         Height          =   375
         Index           =   5
         Left            =   3600
         TabIndex        =   55
         Top             =   120
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Ver Asiento"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Picture         =   "frmBancoMovimientos.frx":1792
      End
   End
   Begin XtremeSuiteControls.PushButton PusActualizarSaldo 
      Height          =   315
      Left            =   12150
      TabIndex        =   39
      Top             =   870
      Width           =   1395
      _Version        =   851968
      _ExtentX        =   2461
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Actualizar Saldo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GBBusqueda 
      Height          =   1815
      Left            =   4020
      TabIndex        =   18
      Top             =   2700
      Visible         =   0   'False
      Width           =   7305
      _Version        =   851968
      _ExtentX        =   12885
      _ExtentY        =   3201
      _StockProps     =   79
      Caption         =   "Busqueda "
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   30
         Picture         =   "frmBancoMovimientos.frx":1D2C
         ScaleHeight     =   585
         ScaleWidth      =   7245
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1200
         Width           =   7245
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
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   120
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
            Index           =   1
            Left            =   75
            TabIndex        =   20
            Top             =   170
            Width           =   1770
         End
      End
      Begin XtremeSuiteControls.CheckBox chkOcultarNoBuscados 
         Height          =   255
         Left            =   1680
         TabIndex        =   22
         Top             =   840
         Visible         =   0   'False
         Width           =   5415
         _Version        =   851968
         _ExtentX        =   9551
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Solo mostrar criterios buscados"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtBusqueda 
         Height          =   315
         Left            =   1680
         TabIndex        =   23
         Top             =   360
         Width           =   5415
         _Version        =   851968
         _ExtentX        =   9551
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Texto a Buscar:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.TabControl TabIngreso 
      Height          =   2895
      Left            =   4590
      TabIndex        =   12
      Top             =   6210
      Visible         =   0   'False
      Width           =   7095
      _Version        =   851968
      _ExtentX        =   12515
      _ExtentY        =   5106
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   5
      Color           =   4
      PaintManager.Position=   2
      ItemCount       =   1
      Item(0).Caption =   "Ing. Movimiento"
      Item(0).ControlCount=   5
      Item(0).Control(0)=   "o1"
      Item(0).Control(1)=   "o2"
      Item(0).Control(2)=   "txtImporte"
      Item(0).Control(3)=   "txtComentario"
      Item(0).Control(4)=   "dtpAltaMovimiento"
      Begin VB.TextBox txtComentario 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   960
         TabIndex        =   16
         Top             =   1440
         Width           =   4680
      End
      Begin VB.TextBox txtImporte 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   1080
         TabIndex        =   15
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton o2 
         Caption         =   "Acreditar <F12>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   3720
         TabIndex        =   14
         Top             =   1080
         Value           =   -1  'True
         Width           =   1785
      End
      Begin VB.OptionButton o1 
         Caption         =   "Debitar <F11>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2280
         TabIndex        =   13
         Top             =   1080
         Width           =   1545
      End
      Begin MSComCtl2.DTPicker dtpAltaMovimiento 
         Height          =   315
         Left            =   1080
         TabIndex        =   17
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   61079553
         CurrentDate     =   38028
      End
   End
   Begin XtremeSuiteControls.TabControl TabBancos 
      Height          =   7425
      Left            =   0
      TabIndex        =   4
      Top             =   1350
      Width           =   13665
      _Version        =   851968
      _ExtentX        =   24104
      _ExtentY        =   13097
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AllowReorder    =   -1  'True
      PaintManager.BoldSelected=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      ItemCount       =   2
      Item(0).Caption =   "Filtrar Movimientos"
      Item(0).ControlCount=   14
      Item(0).Control(0)=   "lblBanco(0)"
      Item(0).Control(1)=   "lblBanco(1)"
      Item(0).Control(2)=   "pbCarga(0)"
      Item(0).Control(3)=   "pbCarga(1)"
      Item(0).Control(4)=   "txtBanco(0)"
      Item(0).Control(5)=   "txtBanco(1)"
      Item(0).Control(6)=   "txtBanco(2)"
      Item(0).Control(7)=   "txtBanco(3)"
      Item(0).Control(8)=   "PBFiltrar"
      Item(0).Control(9)=   "Frame1"
      Item(0).Control(10)=   "chkDiferidos"
      Item(0).Control(11)=   "GroupBox1"
      Item(0).Control(12)=   "GroupBox3"
      Item(0).Control(13)=   "barra"
      Item(1).Caption =   "Ver Datos"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "GroupBox2"
      Item(1).Control(1)=   "checkConciliacion"
      Begin XtremeSuiteControls.ProgressBar barra 
         Height          =   255
         Left            =   240
         TabIndex        =   57
         Top             =   4440
         Width           =   6135
         _Version        =   851968
         _ExtentX        =   10821
         _ExtentY        =   450
         _StockProps     =   93
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   615
         Left            =   600
         TabIndex        =   41
         Top             =   2070
         Width           =   3645
         _Version        =   851968
         _ExtentX        =   6429
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Conciliaciones bancarias :"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton rbConciliados 
            Height          =   315
            Left            =   60
            TabIndex        =   42
            Top             =   240
            Width           =   1245
            _Version        =   851968
            _ExtentX        =   2196
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Conciliados"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbNoConciliados 
            Height          =   315
            Left            =   1320
            TabIndex        =   43
            Top             =   240
            Width           =   1335
            _Version        =   851968
            _ExtentX        =   2355
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "No Conciliados"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbTodos 
            Height          =   315
            Left            =   2820
            TabIndex        =   44
            Top             =   240
            Width           =   705
            _Version        =   851968
            _ExtentX        =   1244
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Todos"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
      End
      Begin VB.CheckBox chkDiferidos 
         Caption         =   "Mostrar solamente cheques diferidos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   660
         TabIndex        =   38
         Top             =   1470
         Width           =   3165
      End
      Begin VB.Frame Frame1 
         Height          =   555
         Left            =   570
         TabIndex        =   33
         Top             =   2880
         Width           =   8655
         Begin Aplisoft_CajasDeTexto.TxF txtFecha 
            Height          =   315
            Index           =   0
            Left            =   1230
            TabIndex        =   34
            Top             =   120
            Width           =   1815
            _ExtentX        =   3201
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
         End
         Begin Aplisoft_CajasDeTexto.TxF txtFecha 
            Height          =   315
            Index           =   1
            Left            =   6390
            TabIndex        =   35
            Top             =   120
            Width           =   2145
            _ExtentX        =   3784
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
         End
         Begin XtremeSuiteControls.Label lblBanco 
            Height          =   255
            Index           =   2
            Left            =   90
            TabIndex        =   37
            Top             =   150
            Width           =   1125
            _Version        =   851968
            _ExtentX        =   1984
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Fecha Desde:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblBanco 
            Height          =   255
            Index           =   3
            Left            =   5310
            TabIndex        =   36
            Top             =   150
            Width           =   1125
            _Version        =   851968
            _ExtentX        =   1984
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Fecha Hasta:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtBanco 
         Height          =   315
         Index           =   0
         Left            =   2250
         TabIndex        =   0
         Top             =   630
         Width           =   975
         _Version        =   851968
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.FlatEdit txtBanco 
         Height          =   315
         Index           =   1
         Left            =   3810
         TabIndex        =   7
         Top             =   630
         Width           =   5445
         _Version        =   851968
         _ExtentX        =   9604
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtBanco 
         Height          =   315
         Index           =   2
         Left            =   2250
         TabIndex        =   1
         Top             =   1020
         Width           =   975
         _Version        =   851968
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   0
         Left            =   3330
         TabIndex        =   8
         Tag             =   "Banco"
         Top             =   630
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtBanco 
         Height          =   315
         Index           =   3
         Left            =   3810
         TabIndex        =   9
         Top             =   1020
         Width           =   5475
         _Version        =   851968
         _ExtentX        =   9657
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   1
         Left            =   3330
         TabIndex        =   10
         Tag             =   "BancoCuenta"
         Top             =   1020
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   7065
         Left            =   -70000
         TabIndex        =   30
         Top             =   360
         Visible         =   0   'False
         Width           =   13635
         _Version        =   851968
         _ExtentX        =   24051
         _ExtentY        =   12462
         _StockProps     =   79
         Caption         =   "GroupBox2"
         UseVisualStyle  =   -1  'True
         Begin Grid.KlexGrid KlexMovimientos 
            Height          =   6975
            Left            =   30
            TabIndex        =   31
            Top             =   90
            Width           =   13575
            _ExtentX        =   23945
            _ExtentY        =   12303
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
            MouseIcon       =   "frmBancoMovimientos.frx":6DDF
            Rows            =   10
         End
      End
      Begin XtremeSuiteControls.PushButton PBFiltrar 
         Height          =   435
         Left            =   120
         TabIndex        =   32
         Top             =   3750
         Width           =   13455
         _Version        =   851968
         _ExtentX        =   23733
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Filtrar Movimientos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Picture         =   "frmBancoMovimientos.frx":6DFB
      End
      Begin XtremeSuiteControls.CheckBox checkConciliacion 
         Height          =   255
         Left            =   -61630
         TabIndex        =   40
         Top             =   60
         Visible         =   0   'False
         Width           =   2955
         _Version        =   851968
         _ExtentX        =   5212
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Activar conciliación con doble clic"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   615
         Left            =   5400
         TabIndex        =   45
         Top             =   2100
         Width           =   3825
         _Version        =   851968
         _ExtentX        =   6747
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Conciliaciones Banacos -  Ctas Contables :"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton rbCbc 
            Height          =   315
            Left            =   150
            TabIndex        =   46
            Top             =   240
            Width           =   1245
            _Version        =   851968
            _ExtentX        =   2196
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Conciliados"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbNCbc 
            Height          =   285
            Left            =   1410
            TabIndex        =   47
            Top             =   240
            Width           =   1425
            _Version        =   851968
            _ExtentX        =   2514
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "No Conciliados"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbTCbc 
            Height          =   315
            Left            =   2970
            TabIndex        =   48
            Top             =   240
            Width           =   765
            _Version        =   851968
            _ExtentX        =   1349
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Todos"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
      End
      Begin XtremeSuiteControls.Label lblBanco 
         Height          =   255
         Index           =   1
         Left            =   570
         TabIndex        =   6
         Top             =   1020
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3087
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Seleccione la Cuenta:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblBanco 
         Height          =   255
         Index           =   0
         Left            =   630
         TabIndex        =   5
         Top             =   630
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3087
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Seleccione el Banco:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "Borrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1290
      Picture         =   "frmBancoMovimientos.frx":7395
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7740
      UseMaskColor    =   -1  'True
      Width           =   765
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   480
      Picture         =   "frmBancoMovimientos.frx":779A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7740
      UseMaskColor    =   -1  'True
      Width           =   825
   End
   Begin XtremeSuiteControls.PushButton cmdVerDetalle 
      Height          =   375
      Left            =   330
      TabIndex        =   11
      Top             =   7260
      Width           =   3975
      _Version        =   851968
      _ExtentX        =   7011
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Ver detalle del movimiento seleccionado"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
      Picture         =   "frmBancoMovimientos.frx":7B9B
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   795
      Left            =   30
      TabIndex        =   25
      Top             =   480
      Width           =   13665
      _Version        =   851968
      _ExtentX        =   24104
      _ExtentY        =   1402
      _StockProps     =   68
      Appearance      =   3
      Color           =   4
      Begin XtremeSuiteControls.GroupBox GroupBox5 
         Height          =   45
         Left            =   -90
         TabIndex        =   56
         Top             =   60
         Width           =   13845
         _Version        =   851968
         _ExtentX        =   24421
         _ExtentY        =   79
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
      End
      Begin VB.Label lblTituloSaldo 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Anterior :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   3090
         TabIndex        =   28
         Top             =   480
         Width           =   1185
      End
      Begin VB.Label lblTituloSaldo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Actual :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   6870
         TabIndex        =   29
         Top             =   480
         Width           =   1005
      End
      Begin XtremeSuiteControls.Label lblSaldo 
         Height          =   240
         Index           =   1
         Left            =   8100
         TabIndex        =   27
         Top             =   450
         Width           =   2715
         _Version        =   851968
         _ExtentX        =   4789
         _ExtentY        =   423
         _StockProps     =   79
         ForeColor       =   255
         BackColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblSaldo 
         Height          =   240
         Index           =   0
         Left            =   4320
         TabIndex        =   26
         Top             =   450
         Width           =   2355
         _Version        =   851968
         _ExtentX        =   4154
         _ExtentY        =   423
         _StockProps     =   79
         ForeColor       =   16744576
         BackColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
   End
End
Attribute VB_Name = "frmBancosMovimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vidBancos As Long
Dim rsBancos As ADODB.Recordset, rsBancosMovimientos As ADODB.Recordset
Const vColorInicial = &H80C0FF
Dim sqlBancosMovimientos As String
Dim vsqlFecha As String

Dim vColor As String
Private Sub CalcularSaldo(vSaldoParcial As Double)
On Error Resume Next
    
    ' ---- Variables ---------------
    Dim i As Integer
    Dim vsql, vvalores, vvcp  As String
    ' ------------------------------

    
    vsql = "delete from Movimientoscaja"
    Call EjecutarScript(vsql, PathDBListados)  ' vacio la tabla
    


   ' PBFiltrar.Enabled = Not True
    
    With rsBancosMovimientos
        If Not .EOF = True Then
            .MoveFirst
            FormatoGrilla (Val(GenerarDato("SELECT COUNT(idBancosMovimientos) as CantidadDeRegistros FROM BancosMovimientos bm WHERE (bm.idBancos = '" & Trim(txtBanco(0).Text) & "' AND bm.idBancosCuentas = " & Trim(Val(txtBanco(2).Text)) & ") AND " + vsqlFecha, "CantidadDeRegistros")))
        Else
            FormatoGrilla (1)
            PBFiltrar.Enabled = True
            Exit Sub
        End If
        
        i = 1
        
        If Not .EOF = False Then Exit Sub
        
        
        KlexMovimientos.Rows = .RecordCount
         
        barra.Max = .RecordCount + 1
        
        Do Until .EOF = True
        
            vSaldoParcial = vSaldoParcial - Val(Format(.Fields("Credito").Value, "#######0.00")) + Val(Format(.Fields("debito").Value, "#######0.00"))
                    
            vvalores = ""
                    
           
           
                    
            KlexMovimientos.TextMatrix(i, 0) = ""
            KlexMovimientos.TextMatrix(i, 1) = EsNulo(.Fields("idBancosMovimientos").Value)
            
            KlexMovimientos.TextMatrix(i, 2) = EsNulo(.Fields("Fecha").Value)
            vvalores = vvalores + "'" + strfecha2(.Fields("Fecha").Value) + "',"
            
            
            KlexMovimientos.TextMatrix(i, 3) = EsNulo(.Fields("NroInterno").Value)
            vvalores = vvalores + Str(Val(EsNulo(.Fields("NroInterno").Value))) + ","
                        
            
            
            vsql = "select TipoMovimiento as cp from asientos where nrointerno=" + EsNulo(.Fields("NroInterno").Value)
            vvcp = traerDatos2(vsql, "cp", pathDBMySQL)
            vvalores = vvalores + "'" + vvcp + "'," ' TipoMovimiento
            
            
            vsql = "select concat (CodigoProveedor,CodigoCliente) as cp from asientos where nrointerno=" + EsNulo(.Fields("NroInterno").Value)
            vvcp = traerDatos2(vsql, "cp", pathDBMySQL)
            vvalores = vvalores + "'" + vvcp + "'," ' Patner
            
            
            KlexMovimientos.TextMatrix(i, 4) = EsNulo(.Fields("Debito").Value)
            vvalores = vvalores + EsNulo(.Fields("Debito").Value) + "," ' Debito
            
            
                      
            
           KlexMovimientos.TextMatrix(i, 5) = EsNulo(.Fields("Credito").Value)
            vvalores = vvalores + EsNulo(.Fields("Credito").Value) + "," ' Credito
            
            
            KlexMovimientos.TextMatrix(i, 7) = EsNulo(.Fields("Comentario").Value)
            vvalores = vvalores + "'" + EsNulo(.Fields("Comentario").Value) + "'," ' Comentario
            
            KlexMovimientos.TextMatrix(i, 8) = EsNulo(.Fields("NroCheque").Value)
            vvalores = vvalores + Str(Val(EsNulo(.Fields("NroCheque").Value))) + ","  ' nrocheque
            
            
            
            KlexMovimientos.TextMatrix(i, 6) = vSaldoParcial
            vvalores = vvalores + Str(vSaldoParcial) ' Comentario
            
            KlexMovimientos.TextMatrix(i, 9) = EsNulo(.Fields("conciliado").Value)
            
            
            KlexMovimientos.Row = i
            KlexMovimientos.Col = 9
            
            If KlexMovimientos.TextMatrix(i, 9) = "OK" Then KlexMovimientos.CellBackColor = vbGreen
            If KlexMovimientos.TextMatrix(i, 9) = "NO" Then KlexMovimientos.CellBackColor = vbRed

            
              
            '--------------
            vvcp = ""
            
            vsql = "select nombre from cuentascorrientes where nrointerno=" + EsNulo(.Fields("NroInterno").Value)

            vvcp = traerDatos2(vsql, "nombre", pathDBMySQL)
            
            vsql = "select nombre from pcuentascorrientes where nrointerno=" + EsNulo(.Fields("NroInterno").Value)

            vvcp = vvcp + traerDatos2(vsql, "nombre", pathDBMySQL)
            
            '----------------
            
            vvalores = vvalores + ",'" + vvcp + "'"
            
            ' graba los datos en la tabla temporal del listado de bancocaja
            vsql = "insert into MovimientosCaja (" + vCampoMovimientosCaja + ") values (" + vvalores + ")"
            Call EjecutarScript(vsql, PathDBListados)
          
          
          
          
          '  .Fields("Saldo").Value = Val(Format(vSaldoParcial, "########0.00"))
            .MoveNext
        
        
            barra.Value = i
            i = i + 1
        Loop
        
        .Fields.Refresh
        
        If Not .EOF = True Then .MoveLast
    
        lblSaldo(1).Caption = Format(vSaldoParcial, "######0.000")
        lblSaldo(1).Alignment = xtpAlignRight
        
        PBFiltrar.Enabled = True
    
    End With
    
    If Err Then
        GrabarLog "CalcularSaldo", Err.Number & " " & Err.Description, Me.Name
    End If
End Sub
Private Function CalcularSaldoAnterior(vFechaLimite As Date) As Double
On Error Resume Next

    CalcularSaldoAnterior = Val(GenerarDato("SELECT Sum(Debito),Sum(Credito),Sum(Debito)-Sum(Credito) as TSaldo FROM BancosMovimientos WHERE ((idBancos  = '" & Trim(txtBanco(0).Text) & "') AND (idBancosCuentas = " & Val(txtBanco(2).Text) & ")) AND (Fecha < '" & strfechaMySQL(vFechaLimite) & "')", "TSaldo"))
    
    If Err Then
        GrabarLog "CalcularSaldoAnterior", Err.Number & " " & Err.Description, Me.Name
    End If
End Function
Private Sub GuardarDebito()
On Error Resume Next

    If Not Val(txtImporte.Text) = 0 Then
        
        With rsBancosMovimientos
            .AddNew
            
            .Fields("codigo").Value = txtBanco(0).Text
            .Fields("Fecha").Value = dtpAltaMovimiento.Value
            .Fields("Debito").Value = Val(txtImporte.Text)
            .Fields("Credito").Value = 0
            .Fields("Comentario").Value = Left(txtComentario.Text, 255)
            
            .Update

            CalcularSaldo (0)
     
        End With
        
        If vConfigGral.vIncluyeContabilidad = True Then
            With frmAsientosAlta
                .Show
                .ZOrder (0)
                .txtCuentaVieneDe.Text = Me.Caption
                .txtImporteVieneDe.Text = Val(txtImporte.Text)
                .dtpFecha.Value = dtpAltaMovimiento.Value
            End With
        End If
        
        cmdNuevo_Click (1)

    Else
        MsgBox "Debe ingresar un importe", vbInformation
        txtImporte.SetFocus
    End If

If Err Then GrabarLog "GuardarDebito", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub GuardarCredito()
On Error Resume Next

    If Not Val(txtImporte.Text) = 0 Then
    
        With rsBancosMovimientos
            .AddNew
            .Fields("codigo").Value = txtBanco(0).Text
            .Fields("fecha").Value = dtpAltaMovimiento.Value
            .Fields("credito").Value = Val(txtImporte.Text)
            .Fields("comentario").Value = Left(txtComentario.Text, 255)
            
            .Update
        
            cmdNuevo_Click (1)
            CalcularSaldo (0)
        
        End With
        
        If vConfigGral.vIncluyeContabilidad = True Then
            With frmAsientosAlta
                .Show
                .ZOrder (0)
                .txtCuentaVieneDe.Text = Me.Caption
                .txtImporteVieneDe.Text = Val(txtImporte.Text)
                .dtpFecha.Value = dtpAltaMovimiento.Value
            End With
        End If
    
    Else
        
        MsgBox "Debe ingresar un importe", vbInformation
        txtImporte.SetFocus
    
    End If

If Err Then GrabarLog "GuardarCredito", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdNuevo_Click(Index As Integer)
On Error Resume Next

    If Index = 0 Then
        txtBanco(0).Text = ""
        txtBanco(1).Text = ""
        txtImporte.Text = ""
        txtComentario.Text = ""
        lblSaldo(0).Caption = ""
        lblSaldo(1).Caption = ""
        txtBanco(0).SetFocus
        
        With Me
            .Top = 300
            .Left = 300
            .Width = 10260
            .Height = 2500
        End With
    
    Else
    
        txtImporte.Text = ""
        txtComentario.Text = ""
        txtImporte.SetFocus

    End If

If Err Then GrabarLog "cmdNuevo_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdBorrar_Click()
    On Error Resume Next
    
    If MsgBox("Confirma la baja del movimiento de Cuenta Corriente del Proveedor ? ", vbYesNo) = vbNo Then
        Exit Sub
    End If

    Dim vArreglo As Double, vSaldoProveedor As Double

    With rsBancosMovimientos
        If Not (.EOF = True) And Not (.BOF = True) Then
            vArreglo = Val(Format(.Fields("debito").Value, "#######0.00")) - Val(Format(.Fields("credito").Value, "#######0.00"))
            lblSaldo(1).Caption = Trim(Val(lblSaldo(1).Caption) + vArreglo)
        
            .Delete
            '.Refresh
        Else
            MsgBox "No tiene seleccionado ningun Movimiento...", vbExclamation, "Mensaje ...."
        End If
    
    End With

    CalcularSaldo (0)
    FormatoGrilla (0)
    
    If Err Then GrabarLog "cmdBorrar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub Imprimir()
    
    On Error Resume Next
    
    Dim vsql, vdiferido As String
    
    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox "Prepare la impresora ...", vbInformation, "Mensaje ..."
      
      
    vdiferido = ""
    If Me.chkDiferidos Then vdiferido = "Cheques diferidos"
      
    
     With drBancosMovimientos2
        .Sections("TituloEmpresa").Controls("lblTitulo").Caption = "Detalle de Movimientos de Banco"
        .Sections("TituloEmpresa").Controls("lblFechaDesde").Caption = txtFecha(0).Value
        .Sections("TituloEmpresa").Controls("lblFechaHasta").Caption = txtFecha(1).Value
        
        .Sections("TituloEmpresa").Controls("lblBanco").Caption = vdiferido + " - " + Me.txtBanco(1) + " / " + Me.txtBanco(3)
        
        .Sections("TituloEmpresa").Controls("lblSaldoAnterior").Caption = "$ " & Format(lblSaldo(0).Caption, "#######0.000")
        
        .Sections("PieInforme").Controls("lblSaldo").Caption = Format(lblSaldo(1).Caption, "#######0.000")
    
        .Show
    End With
    
    
    Exit Sub
        
    
    
    
    
    
    
    
    
    
    
    
    
    
    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox "Prepare la impresora ...", vbInformation, "Mensaje ..."

    'Listado de Movimientos
    With Mantenimiento.rsBancosMovimientos
        If Not .State = 0 Then .Close
        
        
        If vDatosEmpresa.Alias = "Wgestion" Then
            .Source = "SELECT BM.idBancosMovimientos, bm.NroCheque, B.idBancos, B.Descripcion, B.EsCaja, BC.idBancosCuentas, BC.Cuenta, BC.Descripcion, BM.Fecha, BM.Debito, BM.Credito, BM.Saldo,BM.Comentario,aa.TipoMovimiento, BM.NroInterno, concat (aa.`CodigoProveedor`,aa.`CodigoCliente` ) as ClienteProveedor FROM BancosMovimientos BM INNER JOIN Bancos B ON BM.idBancos=B.idBancos LEFT JOIN BancosCuentas BC ON BM.idBancosCuentas=BC.idBancosCuentas left join asientos aa on  bm.nrointerno = aa.nrointerno WHERE (B.idBancos = '" & Trim(txtBanco(0).Text) & "' AND BC.idBancosCuentas = " & Trim(txtBanco(2).Text) & ") and " + vsqlFecha '+ " ORDER BY fecha ASC, idBancosMovimientos ASC"
        Else
             .Source = "SELECT BM.idBancosMovimientos, bm.NroCheque, B.idBancos, B.Descripcion, B.EsCaja, BC.idBancosCuentas, BC.Cuenta, BC.Descripcion, BM.Fecha, BM.Debito, BM.Credito, BM.Saldo,BM.Comentario,aa.TipoMovimiento, BM.NroInterno, concat (aa.`CodigoProveedor`,aa.`CodigoCliente` ) as ClienteProveedor FROM BancosMovimientos BM INNER JOIN Bancos B ON BM.idBancos=B.idBancos LEFT JOIN BancosCuentas BC ON BM.idBancosCuentas=BC.idBancosCuentas left join asientos aa on  bm.nroasiento = aa.numero WHERE (B.idBancos = '" & Trim(txtBanco(0).Text) & "' AND BC.idBancosCuentas = " & Trim(txtBanco(2).Text) & ") and " + vsqlFecha '+ " ORDER BY fecha ASC, idBancosMovimientos ASC"
        End If
        ''.Source = "SELECT BM.idBancosMovimientos, bm.NroCheque, B.idBancos, B.Descripcion, B.EsCaja, BC.idBancosCuentas, BC.Cuenta, BC.Descripcion, BM.Fecha, BM.Debito, BM.Credito, BM.Saldo,BM.Comentario,BM.TipoMovimiento, BM.NroInterno FROM BancosMovimientos BM INNER JOIN Bancos B ON BM.idBancos=B.idBancos LEFT JOIN BancosCuentas BC ON BM.idBancosCuentas=BC.idBancosCuentas left join asientos aa on  bm.NroInterno = aa.NroInterno WHERE (B.idBancos = '" & Trim(txtBanco(0).Text) & "' AND BC.idBancosCuentas = " & Trim(txtBanco(2).Text) & ") and " + vsqlFecha '+ " ORDER BY fecha ASC, idBancosMovimientos ASC"
   
       ' .Source = "SELECT BM.idBancosMovimientos, bm.NroCheque, B.idBancos, B.Descripcion, B.EsCaja, BC.idBancosCuentas, BC.Cuenta, BC.Descripcion, BM.Fecha, BM.Debito, BM.Credito, BM.Saldo,BM.Comentario,BM.TipoMovimiento, BM.NroInterno FROM BancosMovimientos BM INNER JOIN Bancos B ON BM.idBancos=B.idBancos LEFT JOIN BancosCuentas BC ON BM.idBancosCuentas=BC.idBancosCuentas WHERE (B.idBancos = '" & Trim(txtBanco(0).Text) & "' AND BC.idBancosCuentas = " & Trim(txtBanco(2).Text) & ") AND (Fecha >= '" & strfechaMySQL(txtFecha(0).Value) & "' and fecha <= '" & strfechaMySQL(txtFecha(1).Value) & "') ORDER BY fecha ASC, idBancosMovimientos ASC"
        
        If Not .State = 1 Then .Open
        .Close
        .Open
    
    End With
    
    With drBancosMovimientos
        .Sections("TituloEmpresa").Controls("lblTitulo").Caption = "Detalle de Movimientos de Banco"
        .Sections("TituloEmpresa").Controls("lblFechaDesde").Caption = txtFecha(0).Value
        .Sections("TituloEmpresa").Controls("lblFechaHasta").Caption = txtFecha(1).Value
        .Sections("TituloEmpresa").Controls("lblBanco").Caption = Trim(txtBanco(0).Text) & " - " & Trim(txtBanco(1).Text)
        .Sections("TituloEmpresa").Controls("lblSaldoAnterior").Caption = "$ " & Format(lblSaldo(0).Caption, "###,###,##0.000")
        
        .Sections("PieInforme").Controls("lblSaldo").Caption = Format(lblSaldo(1).Caption, "###,###,##0.000")
    
        .Show

    End With

    If Err Then GrabarLog "cmdImprimir_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub chkDiferidos_LostFocus()
If chkDiferidos.Value = 1 Then
    txtFecha(0).Value = Date + 1
    txtFecha(1).Value = Date
    
    Me.Frame1.Enabled = False
Else
    Me.Frame1.Enabled = True
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, _
                       Shift As Integer)
On Error Resume Next

    If KeyCode = vbKeyF1 Then cmdNuevo_Click (0)
    If KeyCode = vbKeyF3 Then PbAcciones_Click (2)
    If KeyCode = vbKeyF11 Then o1.Value = True
    If KeyCode = vbKeyF12 Then o2.Value = True

If Err Then GrabarLog "Form_KeyUp", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Load()
    On Error Resume Next

    With Me
        .Show
        .Top = 0
        .Left = 0
        .Width = 13785
         Me.Height = 9180
        .KeyPreview = True
    End With
    
    
    Me.Left = (Screen.Width - Me.Width) / 2
    'Me.Top = (Screen.Height - Me.Height) / 2 - 4000
    
    dtpAltaMovimiento.Value = Date
    txtFecha(0).Value = Date
    txtFecha(1).Value = Date
    lblSaldo(0).Caption = "0.00"
    lblSaldo(1).Caption = "0.00"

    If Err Then GrabarLog "Form_load", Err.Number & " " & Err.Description, Me.Name
    
    
    init

End Sub

Private Sub init()

If vConfigGral.vIncluyeContabilidad Then
    Me.PbAcciones(4).Enabled = False
Else
    Me.PbAcciones(4).Enabled = True
End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub KlexMovimientos_DblClick()
On Error Resume Next

If Not Me.checkConciliacion.Value = 1 Then Exit Sub

Dim vidrow As Long
Dim vrow As Long
Dim vflag, vsql  As String

vrow = Me.KlexMovimientos.Row
vidrow = Me.KlexMovimientos.TextMatrix(vrow, 1)
vflag = Me.KlexMovimientos.TextMatrix(vrow, 9)

Me.KlexMovimientos.Row = vrow
Me.KlexMovimientos.Col = 9
    

If vflag = "OK" Then
    Me.KlexMovimientos.TextMatrix(vrow, 9) = "NO"
    Me.KlexMovimientos.CellBackColor = vbRed
    
Else
    Me.KlexMovimientos.TextMatrix(vrow, 9) = "OK"
      Me.KlexMovimientos.CellBackColor = vbGreen
End If

vsql = "update  bancosmovimientos set conciliado = " + "'" + Me.KlexMovimientos.TextMatrix(vrow, 9) + "' where idbancosmovimientos = " + Str(vidrow)
Call EjecutarScript(vsql, pathDBMySQL)

    
If Err Then Exit Sub
End Sub

Private Sub pbCarga_Click(Index As Integer)
 On Error Resume Next

    vVuelveBusqueda = Me.Name
    vVieneBusqueda = pbCarga(Index).Tag

    Select Case Index

        Case 0 To pbCarga.Count - 1
            frmBusqueda.Show
        
    End Select

    
    If Err Then GrabarLog "pbCarga_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub PBFiltrar_Click()
On Error Resume Next
    
    
    
    Me.lblSaldo(0).Caption = ""
    Me.lblSaldo(1).Caption = ""
    
    MousePointer = vbHourglass
    FiltrarMovimientos
    MousePointer = vbDefault
    
    Me.TabBancos.SelectedItem = 2
    vColor = vColorInicial
    Me.Height = 9180
    
If Err Then GrabarLog "PBFiltrar_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub FiltrarMovimientos()
On Error Resume Next

    Dim vconciliados As String
    
    If rbConciliados.Value Then vconciliados = " and  conciliado = 'OK'"
        
    If rbNoConciliados.Value Then vconciliados = " and conciliado = 'NO'"
        
    If Me.rbTodos Then vconciliados = " "
        
        

    Dim vsaldoanterior As Double
    If Not txtBanco(0).Text = "" And Trim(txtBanco(0).Text) Then
        
        Set rsBancosMovimientos = New ADODB.Recordset
        
        
        If Me.chkDiferidos.Value = 1 Then
                vsqlFecha = "(bm.Fecha > '" & strfechaMySQL(Date) & "' " + vconciliados + ") ORDER BY bm.Fecha ASC, bm.idBancosMovimientos ASC"

        Else
               vsqlFecha = "(bm.Fecha >= '" & strfechaMySQL(txtFecha(0).Value) & "' and bm.Fecha <= '" & strfechaMySQL(txtFecha(1).Value) & "'" + vconciliados + ") ORDER BY bm.Fecha ASC, bm.idBancosMovimientos ASC"
        
        End If

          
          
        sqlBancosMovimientos = "SELECT * FROM BancosMovimientos bm WHERE (bm.idBancos = '" & Trim(txtBanco(0).Text) & "' AND bm.idBancosCuentas = " & Trim(Val(txtBanco(2).Text)) & ") AND " + vsqlFecha
        
          
          
          
         If Me.rbCbc.Value = True Then
            sqlBancosMovimientos = fSQLConciliaBancoCtas(txtFecha(0).Value, txtFecha(1).Value, "")
         End If
          
          
           
         If Me.rbNCbc.Value = True Then
            sqlBancosMovimientos = fSQLConciliaBancoCtas(txtFecha(0).Value, txtFecha(1).Value, "not")
         End If
           
        
         'If Me.rbTCbc.Value = True Then
         '   sqlBancosMovimientos = fSQLConciliaBancoCtas(txtFecha(0).Value, txtFecha(1).Value, "")
         'End If
         
        
        With rsBancosMovimientos
            .CursorLocation = adUseServer
                        
            Call .Open(sqlBancosMovimientos, ConnDDBB, adOpenDynamic, adLockPessimistic)
            
            lblTituloSaldo(0).Caption = "Saldo anterior al " & txtFecha(0).Value
            vsaldoanterior = CalcularSaldoAnterior(Me.txtFecha(0).Value)
            lblSaldo(0).Caption = Format(vsaldoanterior, "############0.00")
            lblSaldo(0).Alignment = xtpAlignRight
            
            
            
            If Me.rbCbc Or Me.rbNCbc Then
            
                Set Me.KlexMovimientos.Recordset = rsBancosMovimientos
                
            End If
            
            If Me.rbTCbc Then CalcularSaldo (vsaldoanterior)   ' llena la grilla
        
        End With

    Else
        MsgBox "Seleccione al menos un Banco", vbInformation, "Mensaje ..."
        FormatoGrilla (1)
    End If

    
If Err Then GrabarLog "FiltrarMovimientos", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub MostrarCoincidencias(vBusqueda As String)
On Error Resume Next

    Dim sqlBancos As String

    Set rsBancos = New ADODB.Recordset

    If Trim(vBusqueda) = "" Then
        sqlBancos = "SELECT * FROM Bancos WHERE 1=2"
    Else
        sqlBancos = "SELECT * FROM Bancos WHERE (idBancos LIKE '%" & Trim(vBusqueda) & "%') OR (Descripcion LIKE '%" & Trim(vBusqueda) & "%')"
    End If

    With rsBancos
        If .State = 1 Then .Close

        .CursorLocation = adUseClient
    
        Call .Open(sqlBancos, ConnDDBB, adOpenStatic, adLockReadOnly)
    
        'dgBancos.Visible = Not .EOF
    
        If Not .EOF = True Then
            'Set dgBancos.DataSource = rsBancos
            Call FormatoGrilla(1)
        Else
            'Set dgBancos.DataSource = Nothing
        End If
    
    End With

    sqlBancos = ""

If Err Then GrabarLog "MostrarCoincidencias", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub FormatoGrilla(vCantidadRenglones As Integer)
On Error Resume Next

    Dim i As Integer
    
    With KlexMovimientos
        
        .FixedRows = 1
        .FixedCols = 1
    
        .Cols = 10
        .Rows = vCantidadRenglones + 1
        
        If vCantidadRenglones = 1 Then
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
                .ColWidth(i) = 0
            Next
        End If
        
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 400
        
        .TextMatrix(0, 1) = "idBancosMovimientos"
        .ColWidth(1) = 0
               
        .TextMatrix(0, 2) = "Fecha"
        .ColWidth(2) = 1150
        
        .TextMatrix(0, 3) = "Nro Interno"
        .ColWidth(3) = 1000
        
        .TextMatrix(0, 4) = "Debito"
        .ColWidth(4) = 1250
        .ColDisplayFormat(4) = "###,##0.00"
        .ColAlignment(4) = 9
        
        .TextMatrix(0, 5) = "Credito"
        .ColWidth(5) = 1250
        .ColDisplayFormat(5) = "###,##0.00"
        .ColAlignment(5) = 9
        
        .TextMatrix(0, 6) = "Saldo"
        .ColWidth(6) = 1250
        .ColDisplayFormat(6) = "###,##0.00"
        .ColAlignment(6) = 9
        
        .TextMatrix(0, 7) = "Observaciones"
        .ColWidth(7) = 5000

        .TextMatrix(0, 8) = "N.Cheque"
        .ColWidth(8) = 950
        
        .TextMatrix(0, 9) = "Conciliado"
        .ColWidth(9) = 1000
        


        .BackColorAlternate = 14737632
    End With
    
If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub PbAcciones_Click(Index As Integer)
Dim vnroasiento As Long
On Error Resume Next

    Select Case Index
    
        Case 0
            Imprimir
            
        Case 1
            Unload Me
            
        Case 2
            GBBusqueda.Visible = True
            txtBusqueda.SetFocus
        Case 3
            frmIngresosEgresos.Show
            frmIngresosEgresos.cmdGuardar.Enabled = False
            'frmIngresosEgresos.cmdCerrar.Enabled = False
        Case 4
        
        If MsgBox("Confirma el eliminar el movimiento de Banco ? ", vbYesNo) = vbNo Then
                Exit Sub
         End If

        Dim vmotivo As String
        
        vmotivo = InputBox("Ing. el motivo por el que Usd. borra este movimiento", "Borrado ...")

        Call BorrarBase("bancosMovimientos where idBancosMovimientos=" + KlexMovimientos.TextMatrix(KlexMovimientos.RowSel, 1), pathDBMySQL)
        GrabarLog "Borrar Caja", vmotivo, Me.Caption
        Call PBFiltrar_Click
        
        Case 5
        vnroasiento = traerDatos2("select * from bancosmovimientos where idBancosMovimientos=" + Me.KlexMovimientos.TextMatrix(Me.KlexMovimientos.RowSel, 1), "NroAsiento", pathDBMySQL)
        frmAsientos.txtBusqueda(0) = vnroasiento
        frmAsientos.txtBusqueda(1) = vnroasiento
        Call frmAsientos.PbAcciones_Click(4)
    
    End Select

If Err Then GrabarLog "PBAcciones_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub PusActualizarSaldo_Click()
On Error Resume Next

Dim vdebito, vcredito As Double
Dim vsql, vvalores As String


vdebito = Val(InputBox("Debito para ajuste:", "Ajuste de Saldo"))
vcredito = Val(InputBox("Credito para ajuste:", "Ajuste de Saldo"))


If vdebito + vcredito = 0 Then Exit Sub


vvalores = " ('" + Trim(Me.txtBanco(0)) + "'," + Trim(Me.txtBanco(2)) + "," + "'2011-01-01'," + Str(vdebito) + "," + Str(vcredito) + ",'Ajuste de Saldo')"

vsql = "insert into bancosmovimientos (idBancos,idBancosCuentas,fecha,debito,credito,comentario) Values " + vvalores

Call EjecutarScript(vsql, pathDBMySQL)

Call PBFiltrar_Click

If Err Then
    MsgBox "Error al intentar modificar saldo", vbCritical
End If
End Sub

Private Sub txtBanco_Click(Index As Integer)
On Error Resume Next

    txtBanco(Index).SelStart = 0
    txtBanco(Index).SelLength = Len(txtBanco(Index).Text)

If Err Then GrabarLog "txtBanco_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtBanco_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        
        Select Case Index
        
            Case 0
1                If Not (txtBanco(Index).Text) = "" Then
                    txtBanco(Index + 1).Text = TraerDato("Bancos", "idBancos = '" & Trim(txtBanco(Index).Text) & "' AND (EsCaja = 'N')", "Descripcion")
                    txtBanco(Index + 2).SetFocus
                Else
                    txtBanco(Index).SetFocus
                End If
                
            Case 1
            
            Case 2
                If Not txtBanco(Index).Text = "" Then
                    txtBanco(Index + 1).Text = TraerDato("BancosCuentas", "idBancosCuentas = " & Trim(txtBanco(Index).Text) & " AND (idBancos = '" & Trim(txtBanco(Index - 2).Text) & "')", "Cuenta")
                    If txtBanco(Index + 1).Text = "" Then
                        txtBanco(Index).Text = ""
                        txtBanco(Index).SetFocus
                    Else
                        txtFecha(0).SetFocus
                    End If
                Else
                    txtBanco(Index + 1).Text = ""
                    txtBanco(Index).SetFocus
                End If
                

            Case 3
        
        End Select
    
    End If
        
If Err Then GrabarLog "txtBanco_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Function SeleccionarBancoyCuenta(Index As Integer, vValor As Variant)
On Error Resume Next

    vValor = Trim(vValor)
    
    Dim rsBancoyCuenta As New ADODB.Recordset, sqlBancoyCuenta As String
    
    Select Case Index
    
        Case 0
            sqlBancoyCuenta = "SELECT * FROM Bancos WHERE (idBancos = '" & vValor & "')"
        
        Case 2
            sqlBancoyCuenta = "SELECT * FROM BancosCuentas WHERE (Cuenta = '" & vValor & "')"
    
    End Select
    
    With rsBancoyCuenta
        Call .Open(sqlBancoyCuenta, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then
            If Not vValor = "" Then
                Select Case Index
    
                    Case 0
                        txtBanco(0).Text = EsNulo(.Fields("idBancos").Value)
                        txtBanco(1).Text = EsNulo(.Fields("Descripcion").Value)
                        txtBanco(2).SetFocus
                    Case 2
                        txtBanco(2).Text = EsNulo(.Fields("Cuenta").Value)
                        txtBanco(3).Text = EsNulo(.Fields("Descripcion").Value)
                        txtFecha(0).SetFocus
    
                End Select

            End If
        End If
    End With
    
    sqlBancoyCuenta = ""
    
    If rsBancoyCuenta.State = 1 Then
        rsBancoyCuenta.Close
        Set rsBancoyCuenta = Nothing
    End If
    
If Err Then GrabarLog "SeleccionarBancoyCuenta", Err.Number & " " & Err.Description, Me.Caption
End Function

Private Sub txtBanco_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
    If KeyCode = vbKeyF3 Then
        If Index = 2 Then
            pbCarga_Click (1)
        End If
    End If

If Err Then GrabarLog "txtBanco_KeyUp", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtBusqueda_GotFocus()
On Error Resume Next

    With txtBusqueda
        .SelStart = 0
        .SelLength = Len(txtBusqueda.Text)
    End With

If Err Then GrabarLog "txtBusqueda_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
On Error Resume Next

    Dim i As Integer, j As Integer

    If KeyAscii = 13 Then
        
        If Not Trim(txtBusqueda.Text) = "" Then
            
            GBBusqueda.Visible = False
            
            With KlexMovimientos
                .Row = 1
                For i = 1 To Val(.Rows - 1)
                    If (Val(.TextMatrix(i, 3)) = Val(txtBusqueda.Text)) Or (InStr(1, LCase(.TextMatrix(i, 7)), LCase(txtBusqueda.Text)) > 0) Then
                        .Row = i
                        
                        For j = 1 To Val(.Cols - 1)
                            .Col = j
                            .CellBackColor = vColor
                        Next
                        
                    Else
                        If chkOcultarNoBuscados.Value = xtpChecked Then
                            'For j = 1 To Val(.Cols - 1)
                                .Row = i
                                '.RowIsVisible = True
                            'Next
                        End If
                    End If
                    
                Next
           
                Randomize
                vColor = Val(Rnd * vColorInicial)
                
            End With
        
        End If
    
    End If

If Err Then GrabarLog "txtBusqueda_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        If Index = 0 Then txtFecha(Index + 1).SetFocus
        If Index = 1 Then PBFiltrar.SetFocus
        
    End If

If Err Then GrabarLog "txtFecha_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtComentario.SetFocus
    End If

End Sub

