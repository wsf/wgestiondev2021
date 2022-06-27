VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.ShortcutBar.v13.0.0.Demo.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "Copia de KlexGrid.ocx"
Object = "{9746E3DA-06E1-4D26-9CE4-D9F6411A9C70}#1.0#0"; "SMGA_OcxTxt2008.ocx"
Begin VB.Form frmCtaCteC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuenta Corrientes de Clientes"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   16890
   FillColor       =   &H00000080&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   16890
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox vctipo 
      Height          =   315
      Left            =   1650
      TabIndex        =   76
      Top             =   2130
      Width           =   945
   End
   Begin MSDataGridLib.DataGrid dgClientes 
      Height          =   1620
      Left            =   900
      TabIndex        =   12
      Top             =   1410
      Visible         =   0   'False
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   2858
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
   Begin XtremeSuiteControls.GroupBox GroupBox4 
      Height          =   150
      Left            =   60
      TabIndex        =   63
      Top             =   480
      Width           =   16815
      _Version        =   851968
      _ExtentX        =   29660
      _ExtentY        =   265
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1515
      Left            =   0
      TabIndex        =   35
      Top             =   540
      Width           =   8925
      Begin VB.CommandButton cmdSaldosClientes 
         Height          =   315
         Left            =   8310
         Picture         =   "CuentasCorrientes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Impre saldos de clientes"
         Top             =   570
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   315
      End
      Begin XtremeSuiteControls.FlatEdit txtCUIT 
         Height          =   285
         Left            =   900
         TabIndex        =   37
         Top             =   840
         Width           =   3075
         _Version        =   851968
         _ExtentX        =   5424
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtComentario 
         Height          =   285
         Left            =   900
         TabIndex        =   38
         Top             =   1140
         Width           =   4245
         _Version        =   851968
         _ExtentX        =   7488
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   16777215
         BackColor       =   16777215
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtCliente 
         Height          =   285
         Left            =   900
         TabIndex        =   39
         Top             =   540
         Width           =   7245
         _Version        =   851968
         _ExtentX        =   12779
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vDesRepartidor 
         Height          =   315
         Left            =   2310
         TabIndex        =   72
         Top             =   150
         Width           =   5835
         _Version        =   851968
         _ExtentX        =   10292
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4210752
         BackColor       =   -2147483633
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit vcodRepartidor 
         Height          =   315
         Left            =   960
         TabIndex        =   73
         Top             =   150
         Width           =   795
         _Version        =   851968
         _ExtentX        =   1402
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4210752
         BackColor       =   -2147483633
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton PushButton7 
         Height          =   345
         Left            =   1890
         TabIndex        =   74
         Top             =   120
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "F7"
         Appearance      =   6
      End
      Begin VB.Label lblCtaCte 
         Caption         =   "Vendedor :"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   75
         Top             =   210
         Width           =   915
      End
      Begin VB.Label Label27 
         Caption         =   "Obs.: "
         Height          =   195
         Left            =   330
         TabIndex        =   42
         Top             =   1155
         Width           =   525
      End
      Begin VB.Label lblCtaCte 
         Caption         =   "Persona:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   41
         Top             =   570
         Width           =   705
      End
      Begin VB.Label lblCtaCte 
         Caption         =   "C.U.I.T :"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   40
         Top             =   885
         Width           =   735
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   615
      Left            =   0
      TabIndex        =   28
      Top             =   -90
      Width           =   16875
      _Version        =   851968
      _ExtentX        =   29766
      _ExtentY        =   1085
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   375
         Left            =   8880
         TabIndex        =   71
         Top             =   180
         Width           =   1035
         _Version        =   851968
         _ExtentX        =   1826
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Excel"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdImprimir 
         Height          =   375
         Index           =   0
         Left            =   2880
         TabIndex        =   29
         Top             =   180
         Width           =   1185
         _Version        =   851968
         _ExtentX        =   2090
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
         Picture         =   "CuentasCorrientes.frx":0102
      End
      Begin XtremeSuiteControls.PushButton cmdPagos 
         Height          =   375
         Left            =   5760
         TabIndex        =   30
         Top             =   180
         Width           =   1125
         _Version        =   851968
         _ExtentX        =   1984
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Pagos"
         UseVisualStyle  =   -1  'True
         Picture         =   "CuentasCorrientes.frx":069C
      End
      Begin XtremeSuiteControls.PushButton cmdVerMovimientos 
         DragIcon        =   "CuentasCorrientes.frx":6EFE
         Height          =   375
         Left            =   1710
         TabIndex        =   31
         Top             =   180
         Visible         =   0   'False
         Width           =   1155
         _Version        =   851968
         _ExtentX        =   2037
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Ver detalle"
         UseVisualStyle  =   -1  'True
         Picture         =   "CuentasCorrientes.frx":7488
      End
      Begin XtremeSuiteControls.PushButton cmdCerrar 
         Height          =   375
         Index           =   1
         Left            =   15780
         TabIndex        =   32
         Top             =   180
         Visible         =   0   'False
         Width           =   1035
         _Version        =   851968
         _ExtentX        =   1826
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cerrar"
         Appearance      =   2
         Picture         =   "CuentasCorrientes.frx":7A22
      End
      Begin XtremeSuiteControls.PushButton cmdImprimir 
         Height          =   375
         Index           =   1
         Left            =   4080
         TabIndex        =   33
         Top             =   180
         Width           =   1665
         _Version        =   851968
         _ExtentX        =   2937
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Imprimir C/Detalle"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
         Picture         =   "CuentasCorrientes.frx":7E22
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   375
         Left            =   60
         TabIndex        =   34
         Top             =   180
         Width           =   1635
         _Version        =   851968
         _ExtentX        =   2884
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Ver. Transacción"
         UseVisualStyle  =   -1  'True
         Picture         =   "CuentasCorrientes.frx":E684
      End
      Begin XtremeSuiteControls.PushButton PusCambiarSaldo 
         Height          =   375
         Left            =   6900
         TabIndex        =   50
         Top             =   180
         Width           =   1905
         _Version        =   851968
         _ExtentX        =   3360
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Agregar Movimientos"
         UseVisualStyle  =   -1  'True
         Picture         =   "CuentasCorrientes.frx":EC1E
      End
   End
   Begin XtremeSuiteControls.GroupBox GBPagos 
      Height          =   30
      Left            =   16305
      TabIndex        =   15
      Top             =   7935
      Visible         =   0   'False
      Width           =   30
      _Version        =   851968
      _ExtentX        =   -53
      _ExtentY        =   -53
      _StockProps     =   79
      BackColor       =   -2147483632
      Appearance      =   6
      BorderStyle     =   1
      Begin XtremeSuiteControls.TabControl TabControl1 
         Height          =   4125
         Left            =   60
         TabIndex        =   16
         Top             =   150
         Width           =   11535
         _Version        =   851968
         _ExtentX        =   20346
         _ExtentY        =   7276
         _StockProps     =   68
         AllowReorder    =   -1  'True
         Appearance      =   10
         Color           =   32
         PaintManager.BoldSelected=   -1  'True
         PaintManager.OneNoteColors=   -1  'True
         PaintManager.HotTracking=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         ItemCount       =   5
         Item(0).Caption =   "Item"
         Item(0).ControlCount=   0
         Item(1).Caption =   "Item"
         Item(1).ControlCount=   0
         Item(2).Caption =   "Item"
         Item(2).ControlCount=   0
         Item(3).Caption =   "Item"
         Item(3).ControlCount=   0
         Item(4).Caption =   "Detalles del documento asociado al movimiento seleccionado"
         Item(4).ControlCount=   2
         Item(4).Control(0)=   "KlexCobros"
         Item(4).Control(1)=   "cmdSalirPagos"
         Begin Grid.KlexGrid KlexCobros 
            Height          =   3135
            Left            =   -69880
            TabIndex        =   17
            Top             =   420
            Visible         =   0   'False
            Width           =   11295
            _ExtentX        =   19923
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
            MouseIcon       =   "CuentasCorrientes.frx":F1B8
            Rows            =   10
            ScrollBars      =   2
            SelectionMode   =   1
         End
         Begin XtremeSuiteControls.PushButton cmdSalirPagos 
            Height          =   375
            Left            =   -69850
            TabIndex        =   18
            Top             =   3660
            Visible         =   0   'False
            Width           =   11340
            _Version        =   851968
            _ExtentX        =   20002
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Ocultar Detalle"
            Transparent     =   -1  'True
            Appearance      =   6
            ImageAlignment  =   4
            TextImageRelation=   4
         End
      End
   End
   Begin XtremeSuiteControls.TabControl TabControl2 
      Height          =   6555
      Left            =   60
      TabIndex        =   19
      Top             =   2520
      Width           =   16755
      _Version        =   851968
      _ExtentX        =   29554
      _ExtentY        =   11562
      _StockProps     =   68
      ItemCount       =   2
      SelectedItem    =   1
      Item(0).Caption =   "Filtrar Datos"
      Item(0).ControlCount=   9
      Item(0).Control(0)=   "Frame2"
      Item(0).Control(1)=   "cmdFiltroMovimientos"
      Item(0).Control(2)=   "GroupBox3"
      Item(0).Control(3)=   "rbRetenciones"
      Item(0).Control(4)=   "rbTodos"
      Item(0).Control(5)=   "Label3"
      Item(0).Control(6)=   "gretencion"
      Item(0).Control(7)=   "rbPago"
      Item(0).Control(8)=   "rbDeuda"
      Item(1).Caption =   "Ver Datos"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "KlexCtaCte"
      Item(1).Control(1)=   "PusVerFactuas"
      Begin XtremeSuiteControls.GroupBox gretencion 
         Height          =   735
         Left            =   -69550
         TabIndex        =   57
         Top             =   2700
         Visible         =   0   'False
         Width           =   15495
         _Version        =   851968
         _ExtentX        =   27331
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Seleccionar retención para filtrar:"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.PushButton PushButton2 
            Height          =   315
            Left            =   3960
            TabIndex        =   59
            Top             =   300
            Width           =   615
            _Version        =   851968
            _ExtentX        =   1085
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit vCretencion 
            Height          =   345
            Left            =   2670
            TabIndex        =   58
            Top             =   270
            Width           =   1155
            _Version        =   851968
            _ExtentX        =   2037
            _ExtentY        =   609
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vdretencion 
            Height          =   345
            Left            =   4680
            TabIndex        =   60
            Top             =   240
            Width           =   10695
            _Version        =   851968
            _ExtentX        =   18865
            _ExtentY        =   609
            _StockProps     =   77
            BackColor       =   -2147483643
         End
      End
      Begin XtremeSuiteControls.RadioButton rbRetenciones 
         Height          =   435
         Left            =   -63040
         TabIndex        =   54
         Top             =   2220
         Visible         =   0   'False
         Width           =   2385
         _Version        =   851968
         _ExtentX        =   4207
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Retenciones/Persepciones"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   825
         Left            =   -69505
         TabIndex        =   51
         Top             =   1170
         Visible         =   0   'False
         Width           =   15495
         _Version        =   851968
         _ExtentX        =   27331
         _ExtentY        =   1455
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.FlatEdit vcomentario 
            Height          =   375
            Left            =   1170
            TabIndex        =   53
            Top             =   270
            Width           =   9645
            _Version        =   851968
            _ExtentX        =   17013
            _ExtentY        =   661
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "Vcomentario"
         End
         Begin XtremeSuiteControls.RadioButton RadioButton1 
            Height          =   435
            Left            =   11250
            TabIndex        =   77
            Top             =   225
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
            Left            =   12465
            TabIndex        =   78
            Top             =   225
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
            Left            =   14040
            TabIndex        =   79
            Top             =   315
            Width           =   1305
            _Version        =   851968
            _ExtentX        =   2302
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Sólo FACT"
            Appearance      =   6
         End
         Begin VB.Label Label1 
            Caption         =   "Comentarios:"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   360
            Width           =   1215
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid KlexCtaCte 
         Height          =   5775
         Left            =   90
         TabIndex        =   49
         Top             =   720
         Width           =   16515
         _ExtentX        =   29131
         _ExtentY        =   10186
         _Version        =   393216
         BackColor       =   16777215
         BackColorFixed  =   -2147483644
         BackColorBkg    =   -2147483645
         GridColorFixed  =   8421504
         GridLinesFixed  =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Frame Frame2 
         Height          =   675
         Left            =   -69520
         TabIndex        =   20
         Top             =   510
         Visible         =   0   'False
         Width           =   15465
         Begin VB.CheckBox chkSolo_Saldo 
            Caption         =   "Solo Saldo"
            Height          =   255
            Left            =   9450
            TabIndex        =   21
            Top             =   240
            Width           =   1275
         End
         Begin Aplisoft_CajasDeTexto.TxF dtpDesde 
            Height          =   300
            Left            =   1080
            TabIndex        =   22
            Top             =   210
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   529
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
         Begin Aplisoft_CajasDeTexto.TxF dtpHasta 
            Height          =   300
            Left            =   4710
            TabIndex        =   23
            Top             =   210
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   529
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
         Begin XtremeSuiteControls.PushButton PusFiltroAvanzado 
            Height          =   375
            Left            =   12630
            TabIndex        =   24
            Top             =   180
            Visible         =   0   'False
            Width           =   2745
            _Version        =   851968
            _ExtentX        =   4842
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Opiones de filtros avanzado >>>"
            UseVisualStyle  =   -1  'True
            Picture         =   "CuentasCorrientes.frx":F1D4
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "> Hasta:"
            ForeColor       =   &H00004040&
            Height          =   195
            Left            =   3900
            TabIndex        =   26
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   735
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "> Desde:"
            ForeColor       =   &H00004040&
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   795
         End
      End
      Begin XtremeSuiteControls.PushButton cmdFiltroMovimientos 
         Height          =   405
         Left            =   -69550
         TabIndex        =   27
         Top             =   3720
         Visible         =   0   'False
         Width           =   15495
         _Version        =   851968
         _ExtentX        =   27331
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Filtrar Movimentos de CtaCte"
         UseVisualStyle  =   -1  'True
         Picture         =   "CuentasCorrientes.frx":15A36
      End
      Begin XtremeSuiteControls.RadioButton rbTodos 
         Height          =   435
         Left            =   -67660
         TabIndex        =   55
         Top             =   2220
         Visible         =   0   'False
         Width           =   975
         _Version        =   851968
         _ExtentX        =   1720
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Todos"
         Appearance      =   6
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton rbPago 
         Height          =   405
         Left            =   -66460
         TabIndex        =   61
         Top             =   2220
         Visible         =   0   'False
         Width           =   1305
         _Version        =   851968
         _ExtentX        =   2302
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Sólo pagos"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton rbDeuda 
         Height          =   255
         Left            =   -64870
         TabIndex        =   62
         Top             =   2310
         Visible         =   0   'False
         Width           =   1305
         _Version        =   851968
         _ExtentX        =   2302
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Sólo deudas"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton PusVerFactuas 
         Height          =   345
         Left            =   15420
         TabIndex        =   69
         Top             =   360
         Width           =   1125
         _Version        =   851968
         _ExtentX        =   1984
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Ver Factuas"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   465
         Left            =   -69550
         TabIndex        =   56
         Top             =   2190
         Visible         =   0   'False
         Width           =   1545
         _Version        =   851968
         _ExtentX        =   2725
         _ExtentY        =   820
         _StockProps     =   79
         Caption         =   "Tipos de datos:"
      End
   End
   Begin MSAdodcLib.Adodc bprueba 
      Height          =   855
      Left            =   16920
      Top             =   3720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
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
      Connect         =   $"CuentasCorrientes.frx":15FD0
      OLEDBString     =   $"CuentasCorrientes.frx":1606E
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM CuentasCorrientes   WHERE Codigo = '0018';"
      Caption         =   "Adodc1"
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
   Begin VB.Frame fraVerDetalles 
      ForeColor       =   &H8000000F&
      Height          =   3135
      Left            =   14880
      TabIndex        =   13
      Top             =   6360
      Visible         =   0   'False
      Width           =   10815
      Begin MSDataGridLib.DataGrid dgDetalles 
         Height          =   2535
         Left            =   360
         TabIndex        =   14
         Top             =   360
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   4471
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         Appearance      =   0
         BackColor       =   -2147483633
         Enabled         =   0   'False
         ColumnHeaders   =   0   'False
         HeadLines       =   0
         RowHeight       =   19
         RowDividerStyle =   6
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
   End
   Begin VB.Frame ados 
      BackColor       =   &H8000000D&
      Height          =   1635
      Left            =   14640
      TabIndex        =   7
      Top             =   6720
      Visible         =   0   'False
      Width           =   8505
      Begin MSAdodcLib.Adodc bfacturas 
         Height          =   330
         Left            =   5640
         Top             =   1080
         Width           =   2745
         _ExtentX        =   4842
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
         Caption         =   "bfacturas"
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
      Begin MSAdodcLib.Adodc bcliente 
         Height          =   330
         Left            =   120
         Top             =   360
         Width           =   2750
         _ExtentX        =   4842
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
         Caption         =   "bcliente"
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
      Begin MSAdodcLib.Adodc basignacion 
         Height          =   330
         Left            =   2880
         Top             =   720
         Width           =   2750
         _ExtentX        =   4842
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
         Caption         =   "basignacion"
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
      Begin MSAdodcLib.Adodc blibreta_pagos 
         Height          =   330
         Left            =   5640
         Top             =   360
         Width           =   2750
         _ExtentX        =   4842
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
         Caption         =   "blibreta_pagos"
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
      Begin MSAdodcLib.Adodc bremito 
         Height          =   330
         Left            =   120
         Top             =   720
         Width           =   2745
         _ExtentX        =   4842
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
         Caption         =   "bremito"
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
      Begin MSAdodcLib.Adodc bfdetalle 
         Height          =   330
         Left            =   2880
         Top             =   1080
         Width           =   2745
         _ExtentX        =   4842
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
         Caption         =   "bfdetalle"
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
      Begin MSAdodcLib.Adodc bPagoPorMes 
         Height          =   330
         Left            =   5640
         Top             =   720
         Width           =   2745
         _ExtentX        =   4842
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
         Caption         =   "bPagoPorMes"
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
      Begin MSAdodcLib.Adodc barreglo 
         Height          =   330
         Left            =   2880
         Top             =   360
         Visible         =   0   'False
         Width           =   2685
         _ExtentX        =   4736
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\vbprog\WGestion (La Surgente)\Datos\WGestion.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\vbprog\WGestion (La Surgente)\Datos\WGestion.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "CtateCreditosIgualDebitos"
         Caption         =   "barreglo"
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
   End
   Begin VB.CommandButton cmdCarga_Creditos 
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   16800
      Picture         =   "CuentasCorrientes.frx":1610C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Ver crédito"
      Top             =   3120
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdCarga_Cheques 
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   16800
      Picture         =   "CuentasCorrientes.frx":1620E
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Ver crédito"
      Top             =   2880
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   255
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1005
      Left            =   8910
      TabIndex        =   43
      Top             =   630
      Width           =   7965
      _Version        =   851968
      _ExtentX        =   14049
      _ExtentY        =   1773
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton PusW 
         Height          =   915
         Left            =   8460
         TabIndex        =   44
         Top             =   150
         Visible         =   0   'False
         Width           =   150
         _Version        =   851968
         _ExtentX        =   265
         _ExtentY        =   1614
         _StockProps     =   79
         Caption         =   "w"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption lblsaltoTotal 
         Height          =   435
         Left            =   240
         TabIndex        =   70
         Top             =   480
         Width           =   7575
         _Version        =   851968
         _ExtentX        =   13361
         _ExtentY        =   767
         _StockProps     =   14
         ForeColor       =   15591427
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
         Alignment       =   1
         GradientColorLight=   8421504
         GradientColorDark=   8421504
         ForeColor       =   15591427
      End
      Begin XtremeSuiteControls.Label lblSaldo 
         Height          =   450
         Index           =   0
         Left            =   3570
         TabIndex        =   48
         Top             =   60
         Width           =   1575
         _Version        =   851968
         _ExtentX        =   2778
         _ExtentY        =   794
         _StockProps     =   79
         Caption         =   "0.00"
         ForeColor       =   16711680
         BackColor       =   -2147483644
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
      End
      Begin XtremeSuiteControls.Label lblSaldo 
         Height          =   450
         Index           =   1
         Left            =   6150
         TabIndex        =   47
         Top             =   60
         Width           =   1935
         _Version        =   851968
         _ExtentX        =   3413
         _ExtentY        =   794
         _StockProps     =   79
         Caption         =   "0.00"
         ForeColor       =   255
         BackColor       =   -2147483644
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
      End
      Begin VB.Label lblTituloSaldo 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Saldo Anterior al"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   270
         TabIndex        =   46
         Top             =   210
         Width           =   3480
      End
      Begin VB.Label lblTituloSaldo 
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Saldo Hasta :"
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
         Left            =   5160
         TabIndex        =   45
         Top             =   210
         Width           =   960
      End
   End
   Begin XtremeSuiteControls.PushButton pbCarga 
      Height          =   285
      Index           =   0
      Left            =   2760
      TabIndex        =   64
      Tag             =   "TipoMovimientosBanco"
      Top             =   2160
      Width           =   375
      _Version        =   851968
      _ExtentX        =   661
      _ExtentY        =   503
      _StockProps     =   79
      Caption         =   "..."
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit vdtipo 
      Height          =   315
      Left            =   3300
      TabIndex        =   65
      Top             =   2130
      Width           =   5595
      _Version        =   851968
      _ExtentX        =   9869
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.ComboBox vtipoProveedor 
      Height          =   315
      Left            =   15150
      TabIndex        =   67
      Top             =   1680
      Width           =   1665
      _Version        =   851968
      _ExtentX        =   2937
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "Proveedor"
   End
   Begin VB.Label lblTipoDe 
      Alignment       =   1  'Right Justify
      Caption         =   "Tipo de personas: "
      Height          =   255
      Left            =   13200
      TabIndex        =   68
      Top             =   1710
      Width           =   1905
   End
   Begin XtremeSuiteControls.Label lblAltaCaja 
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   66
      Top             =   2190
      Width           =   1575
      _Version        =   851968
      _ExtentX        =   2778
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Tipo de Movimientos:"
      Alignment       =   1
      Transparent     =   -1  'True
   End
   Begin VB.Label credotorgado 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   17760
      TabIndex        =   4
      Top             =   4320
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label saldo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   15360
      TabIndex        =   11
      Top             =   3840
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label dsaldo 
      Appearance      =   0  'Flat
      Caption         =   "Saldo anterior a la fecha "
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   17400
      TabIndex        =   6
      Top             =   2880
      Visible         =   0   'False
      Width           =   3795
   End
   Begin VB.Label vsanterior 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   15480
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label7 
      Caption         =   "Saldo de creditos otorgados :"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   17400
      TabIndex        =   3
      Top             =   2640
      Visible         =   0   'False
      Width           =   3795
   End
   Begin VB.Label Label2 
      Caption         =   "Saldo de Cheques no acreditados :"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   17400
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   3915
   End
   Begin VB.Label saldocheque 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      DataSource      =   "bcliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   15480
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label lbldisplay 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      ForeColor       =   &H008080FF&
      Height          =   195
      Left            =   14880
      TabIndex        =   0
      Top             =   3480
      Width           =   3765
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      ForeColor       =   &H00C0FFFF&
      Height          =   870
      Left            =   14880
      TabIndex        =   8
      Top             =   5160
      Width           =   3195
   End
End
Attribute VB_Name = "frmCtaCteC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCtaCte As ADODB.Recordset
Dim vfsaldo, vlsaldo, vsaldoactual As Double

Dim vnombre, vcomentario_imputacion, vcomentario_imputacion_cambia As String
Dim vsaldoanterior, vtfactura, vtfactura2 As Double
Dim vanomes, vtablaCP, vctacteCP, vid, vfacturaCP, vcobrospagos, vfdetalleCP As String
Dim ttotal_ctacte As Double
Dim vuremito As Long

Dim vfechamanual As Date
Public docheque, vorden As String
Dim vidcreditoctacte As Long

Public vidd As Long
Dim vid_ctacte As Long
Dim i As Integer
Dim vacum As Double 'Acumulador de pagos en la imputación automatica de pago en Ficha Cliente
Dim vCodigoRepartidor As String
Dim vSaldoCliente As String
Dim rsClientes As ADODB.Recordset
Dim rec As ADODB.Recordset
Dim vIdCtaCte As Long
Dim vLeyendaAsiento As String, vTotalAsiento As Double
Dim sqlFiltroGral As String
Dim vnrointerno As Long
Dim vcriterio As String

Private Sub ActualizarVistas()
On Error Resume Next

    Dim vSaldoParcialAnterior As Double
    
    lblSaldo(0).Caption = "0.000"
    lblSaldo(1).Caption = "0.000"
    vSaldoParcialAnterior = ""
    vSaldoCliente = 0
    vlsaldo = 0
    
    'Vista << Ficha de Cliente >>
    Dim sqlCtaCte As String
    
    sqlCtaCte = sqlFiltroGral + "WHERE (cuentascorrientes.codigo = '" & Trim(txtCliente.Tag) & "' ORDER BY fecha ASC, id ASC"
    
    With rsCtaCte
        If .State = 1 Then .Close
        .CursorLocation = adUseServer
        Call .Open(sqlCtaCte, ConnDDBB, adOpenDynamic, adLockPessimistic)
        
        If .EOF = True Then
            lblSaldo(0).Caption = "0.000"
            lblSaldo(1).Caption = "0.000"
        Else
            FiltrarMovimientos
        End If
        
    
    End With
    
    'Vista << Pago pago por mes >>
    Dim vParcial As Double
    
    vParcial = 0
    
    'saldo.Caption = vSaldoCliente
    'lblSaldo(0).Caption = vSaldoCliente
    
If Err Then GrabarLog "ActualizarVistas", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub GuardarSaldoParcialPPM(vSaldoParcial As Double, vIdCCC As Long)
On Error Resume Next

    Dim rsccc As New ADODB.Recordset, sqlCCC As String
    
    sqlCCC = "SELECT * FROM CuentasCorrientes WHERE (id = " & vIdCCC & ")"
    
    With rsccc
        Call .Open(sqlCCC, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not .EOF Then
            .Fields("Saldo_PPM").Value = Val(vSaldoParcial)
            .Update
        End If
        
    End With
    
    sqlCCC = ""
    
    If rsccc.State = 1 Then
        rsccc.Close
        Set rsccc = Nothing
    End If
    
If Err Then GrabarLog "SaldoParcialPPM", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub GuardarSaldoParcialFacturas(vSaldoParcial As Double, vIdCCC As Long)
On Error Resume Next

    Dim rsccc As New ADODB.Recordset, sqlCCC As String
    
    sqlCCC = "SELECT * FROM CuentasCorrientes WHERE (id = " & vIdCCC & ")"
    
    With rsccc
        Call .Open(sqlCCC, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not .EOF Then
            .Fields("Saldo_Facturas").Value = Val(vSaldoParcial)
            .Update
        End If
        
    End With
    
    sqlCCC = ""
    
    If rsccc.State = 1 Then
        rsccc.Close
        Set rsccc = Nothing
    End If
    
If Err Then GrabarLog "SaldoParcialPPM", Err.Number & " " & Err.Description, Me.Name
End Sub
Function BuscarCliente(vViene As String) As Boolean
    On Error Resume Next

    If txtCliente.Text = "" Then Exit Function
    
    MousePointer = vbHourglass
    
    With bcliente
        .ConnectionString = pathDBMySQL
        If vViene = "Normal" Then
            .RecordSource = "SELECT * FROM " + vtablaCP + " WHERE (codigo = '" & Trim(txtCliente.Text) & "') OR (nombre = '" & Trim(txtCliente.Text) & "')"
        Else
            If Not rsClientes.EOF = True Then
                .RecordSource = "SELECT * FROM " + vtablaCP + " WHERE (codigo = '" & Trim(rsClientes.Fields("Codigo").Value) & "')"
            Else
                .RecordSource = "SELECT * FROM " + vtablaCP + " WHERE 1=2"
            End If
        End If
        .Refresh
    
        If .Recordset.EOF = True Then
    
        Else
            
            
            MousePointer = vbHourglass
            
            txtCliente.Tag = EsNulo(.Recordset("codigo").Value)
            txtCliente.Text = EsNulo(.Recordset("Nombre").Value)
            txtComentario.Text = EsNulo(.Recordset("Observaciones").Value)
            
            vnombre = txtCliente.Text
            txtCuit.Text = EsNulo(.Recordset("cuit").Value)
            txtCuit.Locked = True
            
            BuscarCliente = True
            
            dtpDesde.SetFocus
            
        End If
        
    End With
    
    dgClientes.Visible = Not BuscarCliente
    
    MousePointer = vbDefault
    
If Err Then GrabarLog "BuscarCliente", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub cmdSalirPagos_Click()
    On Error Resume Next
    
    GBPagos.Visible = False
    
    If Err Then GrabarLog "cmdSalirPagos_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdVerMovimientos_Click()
On Error Resume Next

If Err Then GrabarLog "cmdVerMovimientos_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub dgClientes_DblClick()
On Error Resume Next

    If BuscarCliente("Grilla") = True Then
        dgClientes.Visible = Not True
        CargoDatosClientes
    End If
        
If Err Then GrabarLog "dgClientes_DblClick", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CargarDetalleCobros(vnroremito As Long)
On Error Resume Next

    Dim rsCobros As New ADODB.Recordset, sqlCobros As String
    
    sqlCobros = "SELECT * FROM " + vcobrospagos + " WHERE (NroInterno = " & vnroremito & ")"
    
    With rsCobros
        .CursorLocation = adUseClient
        
        Call .Open(sqlCobros, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then
            .MoveFirst
            FormatoGrillaCobros (.RecordCount)
        Else
            FormatoGrillaCobros (1)
        End If
        
        Do Until .EOF = True
            KlexCobros.TextMatrix(.AbsolutePosition, 1) = EsNulo(.Fields("id" + vcobrospagos).Value)
            KlexCobros.TextMatrix(.AbsolutePosition, 2) = EsNulo(.Fields("Fecha").Value)
            KlexCobros.TextMatrix(.AbsolutePosition, 3) = EsNulo(.Fields("Remito").Value)
            KlexCobros.TextMatrix(.AbsolutePosition, 4) = EsNulo(.Fields("idMedioPago").Value)
            KlexCobros.TextMatrix(.AbsolutePosition, 5) = EsNulo(.Fields("Importe").Value)
            KlexCobros.TextMatrix(.AbsolutePosition, 6) = EsNulo(.Fields("TipoMovimiento").Value)
            KlexCobros.TextMatrix(.AbsolutePosition, 7) = EsNulo(.Fields("NroInterno").Value)
        
            .MoveNext
        Loop

    End With
    
    sqlCobros = ""

    If rsCobros.State = 1 Then
        rsCobros.Close
        Set rsCobros = Nothing
    End If
    
If Err Then GrabarLog "CargarDetalleCobros", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub FormatoGrillaCobros(vCantidadRenglones As Integer)
On Error Resume Next

    Dim i As Integer

    GBPagos.Visible = True
    GBPagos.Top = 1530
    GBPagos.Left = 120
    
    With KlexCobros
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
        
        .TextMatrix(0, 1) = "idCobros"
        .ColWidth(1) = 1100
               
        .TextMatrix(0, 2) = "Fecha"
        .ColWidth(2) = 2340
        
        .TextMatrix(0, 3) = "Remito"
        .ColWidth(3) = 1500
        
        .TextMatrix(0, 4) = "idMedioPago"
        .ColWidth(4) = 1000
        
        .TextMatrix(0, 5) = "Importe"
        .ColWidth(5) = 1000
        .ColDisplayFormat(5) = "#0.000"
        
        .TextMatrix(0, 6) = "Tipo Movimiento"
        .ColWidth(6) = 1000
        
        .TextMatrix(0, 7) = "Nro Compr."
        .ColWidth(7) = 1000

    End With
    
If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub FormatoGrilla(vCantidadRenglones As Integer)
On Error Resume Next

    Dim i As Integer

    'GBPagos.Visible = True
    
    With KlexCtaCte
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
        .ColWidth(0) = 200
        
        .TextMatrix(0, 1) = "idCtaCte"
        .ColWidth(1) = 0
               
        .TextMatrix(0, 2) = "Fecha"
        .ColWidth(2) = 1000
        
        .TextMatrix(0, 3) = "Nro Interno"
        .ColWidth(3) = 900
        
        .TextMatrix(0, 4) = "Debito"
        .ColWidth(4) = 1350
        
       ' .ColDisplayFormat(4) = "#0.000"
        .ColAlignment(4) = 6
        
        .TextMatrix(0, 5) = "Credito"
        .ColWidth(5) = 1350
       ' .ColDisplayFormat(5) = "#0.000"
        .ColAlignment(5) = 6
        
        .TextMatrix(0, 6) = "Saldo"
        .ColWidth(6) = 1350
        '.ColDisplayFormat(6) = "#0.000"
        .ColAlignment(6) = 6
        
        .TextMatrix(0, 7) = "Observaciones"
        .ColWidth(7) = 7500
        

        .TextMatrix(0, 8) = "Tipo"
        .ColWidth(8) = 350
        
        .TextMatrix(0, 9) = "NroComp"
        .ColWidth(9) = 1000
        

     '   .BackColorAlternate = 14737632
    End With
    
If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub FormatoGrillaRetenciones()
On Error Resume Next

    Dim i As Integer

    'GBPagos.Visible = True
    
    
    
    
    With KlexCtaCte
    

    
        .FixedRows = 1
        .FixedCols = 1
    
        .Cols = 9
        .Rows = 2
       
        
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 200
        
        .TextMatrix(0, 1) = "idCtaCte"
        .ColWidth(1) = 0
               
        .TextMatrix(0, 2) = "Fecha"
        .ColWidth(2) = 1000
        
        .TextMatrix(0, 3) = "Nro Interno"
        .ColWidth(3) = 900
        
        .TextMatrix(0, 4) = "Codigo"
        .ColWidth(4) = 1000
        
       ' .ColDisplayFormat(4) = "#0.000"
       ' .ColAlignment(4) = 6
        
        .TextMatrix(0, 5) = "Retención"
        .ColWidth(5) = 2500
       ' .ColDisplayFormat(5) = "#0.000"
       ' .ColAlignment(5) = 6
        
        
        
        .TextMatrix(0, 6) = "Importe"
        .ColWidth(6) = 1000
        '.ColDisplayFormat(6) = "#0.000"
        .ColAlignment(6) = 6
    
    
        .TextMatrix(0, 7) = "Persona"
        .ColWidth(7) = 3000
        '.ColDisplayFormat(6) = "#0.000"
        .ColAlignment(7) = 1
    
    
        
        .TextMatrix(0, 8) = "Comentario"
        .ColWidth(8) = 4000
        '.ColDisplayFormat(6) = "#0.000"
        .ColAlignment(8) = 1
        
        
     '   .BackColorAlternate = 14737632
    End With
    
If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Caption
End Sub


Private Sub dtpDesde_KeyPress(KeyAscii As Integer)
On Error Resume Next
    
    If KeyAscii = 13 Then
        dtpHasta.SetFocus
    End If

If Err Then GrabarLog "dtpDesde_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub dtpHasta_KeyPress(KeyAscii As Integer)
On Error Resume Next
    
    If KeyAscii = 13 Then
        cmdFiltroMovimientos.SetFocus
    End If

If Err Then GrabarLog "dtpHasta_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub KlexCtaCte_Click()
vIdCtaCte = Val(KlexCtaCte.TextMatrix(KlexCtaCte.Row, 1))
vnrointerno = Val(KlexCtaCte.TextMatrix(KlexCtaCte.Row, 3))
End Sub

Private Sub KlexCtaCte_DblClick()
On Error Resume Next

    With KlexCtaCte
        'COntrolo el Rows
            If Not IsNull(.TextMatrix(.Row, 1)) = "" Then
                CargarDetalleCobros (.TextMatrix(.Row, 3))
            End If
            
            
        'Else
        
        'End If
    
    End With
    
If Err Then GrabarLog "KlexCtaCte_DblClick", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cmdCerrar_Click(Index As Integer)
On Error Resume Next

    Unload Me

If Err Then GrabarLog "PusCerrar_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cmdPagos_Click()
On Error Resume Next
    
    If GBPagos.Visible = True Then
        GBPagos.Visible = False
    Else
        With rsCtaCte
            If Not .EOF = True Then
                CargarDetalleCobros (.Fields("Remito").Value)
            Else
        
            End If
        End With
        GBPagos.Visible = True
    End If
    
If Err Then GrabarLog "cmdPagos_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub PusBorrar_Click()

    On Error Resume Next
    Dim vmensaje, vborrado As String
    vmensaje = ""






    If (vctacteCP = "cuentascorrientes") And (Me.KlexCtaCte.TextMatrix(Me.KlexCtaCte.Row, 8) = "CC" Or Me.KlexCtaCte.TextMatrix(Me.KlexCtaCte.Row, 8) = "CD") And Val(Me.KlexCtaCte.TextMatrix(Me.KlexCtaCte.Row, 5)) > 0 Then
        MsgBox "No se puedo borrar un crédito de un documento de contado. " + Chr(13) + "Debe borrarlo desde el movimiento de debito", vbInformation, "Cuidado...."
        Exit Sub
    End If


    If (vctacteCP = "pcuentascorrientes") And (Me.KlexCtaCte.TextMatrix(Me.KlexCtaCte.Row, 8) = "CC" Or Me.KlexCtaCte.TextMatrix(Me.KlexCtaCte.Row, 8) = "CD") And Val(Me.KlexCtaCte.TextMatrix(Me.KlexCtaCte.Row, 4)) > 0 Then
        MsgBox "No se puedo borrar un débito de un documento de contado. " + Chr(13) + "Debe borrarlo desde el movimiento de Crédito", vbInformation, "Cuidado...."
        Exit Sub
    End If



 
   If MsgBox("Confirma el eliminar el movimiento de Cuenta Corriente del Cliente ? ", vbYesNo) = vbNo Then
        Exit Sub
    End If
   
    PushButton1_Click

    Exit Sub


    Dim vArreglo As Double, vSaldoProveedor As Double

    With rsCtaCte
        .Close
        Call .Open("select *  from " + vctacteCP + "  where  " + vid + "=" + Str(vIdCtaCte), ConnDDBB, adOpenDynamic, adLockPessimistic)
        
        If Not (.EOF = True) And Not (.BOF = True) Then
            vArreglo = Val(Format(.Fields("debito").Value, "#######0.00")) - Val(Format(.Fields("credito").Value, "#######0.00"))
            saldo.Caption = Str(Val(saldo.Caption) + vArreglo)
        
            If Not IsNull(.Fields("idCheques").Value) = True Then
                Call BorrarBase("Cheques WHERE (idCheques = " & .Fields("idCheques").Value & ")", pathDBMySQL)
                vmensaje = vmensaje + Chr(13) + "# Se ha borrado el movimiento de CHEQUE asociado"
            End If
            If Not IsNull(.Fields("ReMito").Value) = True Or Not (.Fields("Remito").Value = 0) Then
                Call BorrarBase(vfacturaCP + " WHERE (Remito = " & .Fields("Remito").Value & ")", pathDBMySQL)
                Call BorrarBase(vfdetalleCP + " WHERE (Remito = " & .Fields("Remito").Value & ")", pathDBMySQL)
                vmensaje = vmensaje + Chr(13) + "# Se ha borrado el DOCUMENTO de venta/compra asociado"
            End If
            
            
            vborrado = .Fields("codigo") & "   " & Str(.Fields("fecha")) & "  " & Str(.Fields("nrointerno"))
            
            GrabarLog "Borrar.CtaCte", vborrado, Me.Name
            
            Call BorrarBase(vctacteCP + " WHERE (" + vid + " = " & Str(vIdCtaCte) & ")", pathDBMySQL)
            
            
            '.Refresh
            'Me.dgMovimientos.Refresh
            
            MsgBox "Fueron Borrado los siguientes datos: " + Chr(13) + vborrado + Chr(13) + vmensaje, vbInformation

            
            
        
        Else
            MsgBox "No tiene seleccionado ningun Movimiento...", vbExclamation, "Mensaje ...."
        End If
    
    End With
    
    

    
    vSaldoProveedor = 0
    'vSaldoProveedor = Val(TraerDato("Proveedores", "Codigo = '" & Trim(txtProveedor.Tag) & "'", "Saldo")) + vArreglo
    
    'Call EjecutarScript("UPDATE Proveedores SET saldo = '" & vSaldoProveedor & "' WHERE (codigo = '" & Trim(txtProveedor.Tag) & "')")

    ' refresh de grilla
    FiltrarMovimientos

    'CalcularSaldo (0)

    'MsgBox "Esta accion NO BORRA ningun DOCUMENTO relacionado!!", vbExclamation, "Mensaje ...."

    If Err Then GrabarLog "cmdBorrar_Click", Err.Number & " " & Err.Description, Me.Name


End Sub

Private Sub pbCarga_Click(Index As Integer)
Call fbuscarGrilla("(select * from tipomovimientos where ProveedorClienteBanco ='P') as t", "TipoMovimiento", "Codigo", Me.vdtipo.Name, Me)     ' ema:
End Sub

Private Sub PusCambiarSaldo_Click()
On Error Resume Next
Dim vsaldo As Double

'vsaldo = Val(InputBox("Ingrese el saldo que desea fijar: ", "Fijar Saldo"))


frmCambioSaldoCtate.vcodigo = Me.txtCliente.Tag
frmCambioSaldoCtate.vnombre = Me.txtCliente.Text
frmCambioSaldoCtate.vsaldoactual = CDbl(Me.lblSaldo(1).Caption)
frmCambioSaldoCtate.vctacteCP = vctacteCP
frmCambioSaldoCtate.Show

On Error Resume Next
End Sub

Private Sub PushButton1_Click()
    frmTransaccionMantenimiento.vnrointerno = vnrointerno
    frmTransaccionMantenimiento.Show
End Sub

Private Sub PushButton3_Click()
On Error Resume Next
    
'Call Buscar("", "todos")
  
  
Call grillaToExcel(Me.KlexCtaCte)

If Err Then Exit Sub
End Sub

Private Sub PushButton7_Click()
Dim vsql, vc1, vc2 As String

vsql = "(Select * from proveedores where tipocliente  = 'Vendedor') t"
vc1 = "Nombre"
vc2 = "Codigo"

Call fbuscarGrilla(vsql, vc1, vc2, Me.vDesRepartidor.Name, Me)
End Sub

Private Sub PusVerFactuas_Click()
   
    frmBuscarFactura.txtCliente = Me.txtCliente
    frmBuscarFactura.txtcodigoCliente = Me.txtCliente.Tag
    frmBuscarFactura.txtCliente.Tag = Me.txtCliente.Tag
    
    If vctacteCP = "pcuentascorrientes" Then
        frmBuscarFactura.cpFactura = "pfactura"
        frmBuscarFactura.CP = "proveedores"
    End If
    
    If vctacteCP = "cuentascorrientes" Then
        frmBuscarFactura.cpFactura = "factura"
        frmBuscarFactura.CP = "clientes"
    End If
    
    Call frmBuscarFactura.cmdFiltrar_Click
End Sub

Private Sub PusW_Click()
If PusCambiarSaldo.Enabled = False Then

    If InputBox("Ingresar clave para poder ajustar saldo:") = "wsf.2011" Then
        PusCambiarSaldo.Enabled = True
    End If
End If
End Sub

Private Sub rbRetenciones_Click()
Me.gRetencion.Visible = True
End Sub

Public Sub txtCliente_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    If KeyCode = 38 Then
        With rsClientes
            If Not .EOF = True And Not .BOF = True Then
                .MovePrevious
            Else
                .MoveLast
            End If
        End With
    End If

    If KeyCode = 40 Then
        With rsClientes
            If Not .EOF = True And Not .BOF = True Then
                .MoveNext
            Else
                .MoveFirst
            End If
        End With
    End If
    
    If KeyCode = 13 Then
        
        If BuscarCliente("Grilla") = True Then
           ' CargoDatosClientes
        End If
    End If
    
If Err Then GrabarLog "txtCliente_KeyUp", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CalcularDocumento()
    On Error Resume Next
    
    Dim rsCheques As New ADODB.Recordset, sqlCheques As String
    Dim vMontoTotal As Double

    sqlCheques = "SELECT SUM(Monto) as MontoTotal FROM cheques WHERE (estado = 'No Acreditado') and (codigo = '" & Trim(txtCliente.Tag) & "') and (cp = 'c')"

    With rsCheques
        Call .Open(sqlCheques, ConnDDBB, adOpenStatic, adLockReadOnly)

        vMontoTotal = 0

        If Not .EOF = True Then vMontoTotal = Val(Format(.Fields("MontoTotal").Value, "#####0.000"))

        saldocheque.Caption = (vMontoTotal)

    End With
    
    sqlCheques = ""
    
    If rsCheques.State = 1 Then
        rsCheques.Close
        Set rsCheques = Nothing
    End If
    
    If Err Then GrabarLog "CalcularDcumento", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Function CalcularSaldo(vSaldoParcial As Double) As Double  ' Alfredo: ac'a llena la grilla y calcula el total del saldo mientra la llena
     On Error Resume Next
    
    Dim vvvsaldo, vsa, vd, vc, vs, vsg As Double
    
    '------------ control -----------------
    vsa = vSaldoParcial
    '-------------------------------------
    
    Dim i As Integer
    
    
    Me.KlexCtaCte.Clear
    
    With rsCtaCte
        If Not .EOF = True Then
            .MoveFirst
            dtpDesde.Value = strfechaMySQL(.Fields("Fecha").Value)
           If Me.txtCliente.Tag = "" Then
                FormatoGrilla (Val(GenerarDato("SELECT COUNT(" + vid + ") as CantidadDeRegistros FROM " + vctacteCP + " WHERE (Fecha >= '" & strfechaMySQL(dtpDesde.Value) & "' and Fecha <= '" & strfechaMySQL(dtpHasta.Value) & "') ", "CantidadDeRegistros")))
           Else
                FormatoGrilla (Val(GenerarDato("SELECT COUNT(" + vid + ") as CantidadDeRegistros FROM " + vctacteCP + " WHERE (codigo = '" & Trim(txtCliente.Tag) & "') AND (Fecha >= '" & strfechaMySQL(dtpDesde.Value) & "' and Fecha <= '" & strfechaMySQL(dtpHasta.Value) & "') ", "CantidadDeRegistros")))
           End If
        
            'FormatoGrilla (vnregistros)
        Else
            dtpDesde.Value = Date
            dtpHasta.Value = Date
        End If
        i = 1
        
        vvvsaldo = 0
        
                
         
        Do Until .EOF = True
        
            'If Me.Tag = "Proveedores" Then
            '    vSaldoParcial = vSaldoParcial + Val(Format(.Fields("Credito").Value, "#######0.000")) - Val(Format(.Fields("debito").Value, "#######0.000"))
            'Else
                vSaldoParcial = vSaldoParcial - Val(Format(.Fields("Credito").Value, "#######0.000")) + Val(Format(.Fields("debito").Value, "#######0.000"))
            'End If
            
            
            ' '------------ control ---------------------------------------------------
            vd = vd + Val(Format(.Fields("debito").Value, "#######0.000"))
            vc = vc + Val(Format(.Fields("Credito").Value, "#######0.000"))
            '-------------------------------------------------------------------------
            
            
            
            
            vlsaldo = vSaldoParcial
            ' panic !!! sacar el llenado dela grilla del calculo de saldo
            KlexCtaCte.TextMatrix(i, 0) = ""
            KlexCtaCte.TextMatrix(i, 1) = EsNulo(.Fields(vid).Value) ' vid contiene el nombre de la id para ctacte provee - clie
            KlexCtaCte.TextMatrix(i, 2) = EsNulo(.Fields("Fecha").Value)
            KlexCtaCte.TextMatrix(i, 3) = EsNulo(.Fields("NroInterno").Value)
            
            KlexCtaCte.TextMatrix(i, 4) = formatNumero(EsNulo(.Fields("Debito").Value))
            KlexCtaCte.TextMatrix(i, 5) = formatNumero(EsNulo(.Fields("Credito").Value))
            
            KlexCtaCte.TextMatrix(i, 6) = formatNumero(vSaldoParcial)

            KlexCtaCte.TextMatrix(i, 7) = EsNulo(.Fields("Comentario").Value)
            
           ' KlexCtaCte.TextMatrix(i, 7) = fbanco(EsNulo(.Fields("NroAsiento").Value)) + EsNulo(.Fields("Comentario").Value)
            KlexCtaCte.TextMatrix(i, 8) = EsNulo(.Fields("TipoMovimiento").Value)
            KlexCtaCte.TextMatrix(i, 9) = TraerDato(vfacturaCP, "remito=" + EsNulo(.Fields("remito").Value), "NComprobante", pathDBMySQL) ' Alfredo: acá voy a buscar el nror de factura enla tabla factua teniendo el nro de remito de la tabla cuentascorrientes
            
            ' ----------------------- saldo del filtro
            vvvsaldo = vvvsaldo + Val(EsNulo(.Fields("Debito").Value)) - Val(EsNulo(.Fields("Credito").Value))
            '-------------------------
            
            '.Fields("Saldo").Value = Val(Format(vSaldoParcial, "###,###,##0.00"))
             
            If Not Me.rbRetenciones Then
                .Fields("Saldo").Value = vSaldoParcial
            End If
            
            
            vsg = .Fields("Saldo").Value
            
            
            
            i = i + 1
            KlexCtaCte.Rows = i + 1
            
            
            .MoveNext
            
        Loop
        
        .Fields.Refresh
        
        If .EOF = True Then
            .MoveLast
            KlexCtaCte.Row = i - 1
            
        End If
    End With
    
    vSaldoCliente = vSaldoParcial
    CalcularSaldo = vSaldoParcial
    
    
    '------------ control -----------------
    vs = vsa + vd - vc
    '--------------------------------------
    
    
    If vsg - vs > 0.01 Then
           MsgBox "Hay un problema con esta cuenta. Verifique los datos", vbCritical
    End If
    

If Err Then GrabarLog "CalcularSaldo", Err.Number & " " & Err.Description, Me.Name
End Function

Function CalcularRetenciones() As Double    ' Alfredo: ac'a llena la grilla y calcula el total del saldo mientra la llena
    On Error Resume Next
    
    Dim vsaldo, vsa, vd, vc, vs, vsg As Double
    
    Dim i As Integer
    
    
    Me.KlexCtaCte.Clear
    
    FormatoGrillaRetenciones
    
   
    
    With rsCtaCte
        If Not .EOF = True Then
            .MoveFirst
            dtpDesde.Value = strfechaMySQL(.Fields("Fecha").Value)
          
        Else
            Exit Function
        End If
        i = 1
        
        vsaldo = 0
        
        KlexCtaCte.Rows = 2
        
        Do Until .EOF = True
            KlexCtaCte.TextMatrix(i, 0) = ""
            KlexCtaCte.TextMatrix(i, 1) = EsNulo(.Fields(vid).Value) ' vid contiene el nombre de la id para ctacte provee - clie
            KlexCtaCte.TextMatrix(i, 2) = EsNulo(.Fields("Fecha").Value)
            KlexCtaCte.TextMatrix(i, 3) = EsNulo(.Fields("NroInterno").Value)
            KlexCtaCte.TextMatrix(i, 4) = EsNulo(.Fields("idRetencion").Value)
            KlexCtaCte.TextMatrix(i, 5) = EsNulo(.Fields("Descrip").Value)
            KlexCtaCte.TextMatrix(i, 6) = formatNumero(EsNulo(.Fields("importe").Value))
            KlexCtaCte.TextMatrix(i, 7) = formatNumero(EsNulo(.Fields("Nombre").Value))
            KlexCtaCte.TextMatrix(i, 8) = EsNulo(.Fields("Comentario").Value)
            
            i = i + 1
            KlexCtaCte.Rows = i + 1
            vsaldo = vsaldo + EsNulo(.Fields("importe").Value)
            
            .MoveNext
        Loop
        
        .Fields.Refresh
        
        If .EOF = True Then
            .MoveLast
            KlexCtaCte.Row = i - 1
            
        End If
    End With
    
    CalcularRetenciones = vsaldo
    
 
    
If Err Then GrabarLog "CalcularSaldo", Err.Number & " " & Err.Description, Me.Name
End Function





Function fbanco(v As String) As String
On Error Resume Next
Dim vsql As String

vsql = " SELECT " & _
        "  bancos.descripcion " & _
        " From " & _
        " cuentascorrientes " & _
        " INNER JOIN bancosmovimientos ON (cuentascorrientes.NroAsiento=bancosmovimientos.NroAsiento) " & _
        " INNER JOIN bancos ON (bancosmovimientos.idBancos=bancos.idBancos) where cuentascorrientes.NroAsiento=" + Trim(v)
        
fbanco = "[" + traerDatos2(vsql, "descripcion", pathDBMySQL) + "]"

If Err Then
fbanco = "[]"
End If
End Function
Private Sub cmdActualizar_Click()
On Error Resume Next

    If txtCliente.Text = "" Then Exit Sub
    chkSolo_Saldo.Value = 0
    If MsgBox(" ¿ Desea actualizar todas las vistas ?", vbInformation + vbYesNo, "Mensaje ...") = vbYes Then
        
        ActualizarVistas
            
    End If
If Err Then GrabarLog "cmdActualizar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdCalcularMonto_Click()
On Error Resume Next

    Dim vSaldoArreglo  As Double
    
    'vfecha_arreglo.Caption = Str(bpagopormes.Recordset("últimodefecha").Value)
    'vsumadedebito_arreglo.Caption = Str(bpagopormes.Recordset("sumadedebito").Value)
    'vsumadecredito_arreglo.Caption = Str(bpagopormes.Recordset("sumadecredito").Value)
    'vSaldoArreglo = bpagopormes.Recordset("saldo").Value
    'vdiferencia_arreglo.Caption = (vSaldoArreglo) - Val(vimporte_arreglo.Text)

If Err Then GrabarLog "cmdCalcularMonto_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdCarga_Cheques_Click()
    On Error Resume Next
    
    With frmCheques
        .txtNombre = Trim(txtCliente.Tag)
        .txtNombre_KeyPress 13
        .TabCheques.tab = 2
    End With
    
    If Err Then GrabarLog "cmdCarga_Cheques_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdCarga_Creditos_Click()
    On Error Resume Next
    
   
    
    If Err Then GrabarLog "cmdCarga_Creditos_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdfiltrardoc_Click()
On Error Resume Next
    With bFacturas
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        '.RecordSource = "Select * from cuentascorrientes where codigo = '" & trim(txtcliente.tag) & "' and (debito > 0) and (Fecha >= '" & strfechaMySQl(fdesdedoc.Value) + "' and Fecha <= '" & strfechaMySQl(fhastadoc.Value) + "')"
        .RecordSource = "Select Fecha, codigo, Comentario, credito, remito, sumadetotaliva, sumadepago, sumaderesta, diferencia, últimodefecha, id, ncomprobante from factura_ctacte where codigo = '" & Trim(txtCliente.Tag) & "' and (Comentario like '%Documento%' or Comentario like '%Dif. por Consumo%') and credito = 0 ORDER BY fecha ASC, id ASC"
        .Refresh
        If Not .Recordset.RecordCount = 0 Then
            .Recordset.MoveFirst
            'cmdejecutardoc.Enabled = True
            'Mostrar Cantidad total de Documentos para generar Factura.....
        Else
            'cmdejecutardoc.Enabled = False
            'Mostrar que no existen Documentos paraa generar Factura.......
        End If

        Do Until .Recordset.EOF = True
            If Not Val(.Recordset("SumaDePago").Value) = 0 Then
                MsgBox "NO PUEDE AGRUPAR DOCUMENTOS QUE TIENE PAGOS", vbExclamation, "Mensaje ..."
                'cmdejecutardoc.Enabled = False
            End If
            
            .Recordset.MoveNext
        Loop
        
    
    End With
If Err Then GrabarLog "cmdfiltrardoc_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Function CalcularSaldoAnterior(vFechaLimite As Date) As Double
On Error Resume Next

   ' If Me.Tag = "Proveedores" Then
        
   '         CalcularSaldoAnterior = Val(GenerarDato("SELECT Sum(Debito),Sum(Credito),Sum(Credito)-Sum(Debito) as TSaldo FROM " + vctacteCP + " WHERE (Codigo  = '" & Trim(txtCliente.Tag) & "') AND (Fecha < '" & strfechaMySQL(vFechaLimite) & "')", "TSaldo"))
   ' Else
            CalcularSaldoAnterior = Val(GenerarDato("SELECT Sum(Debito),Sum(Credito),Sum(Debito)-Sum(Credito) as TSaldo FROM " + vctacteCP + " WHERE (Codigo  = '" & Trim(txtCliente.Tag) & "') AND (Fecha < '" & strfechaMySQL(vFechaLimite) & "')", "TSaldo"))
    
    
   ' End If
    
    If Err Then
        GrabarLog "CalcularSaldoAnterior", Err.Number & " " & Err.Description, Me.Name
    End If
End Function
Public Sub cmdFiltroMovimientos_Click()
    On Error Resume Next
    
    
    If Me.radioTodoDoc.Value Then Me.vcomentario = "Documento"
    
    If Me.Radsolofact.Value Then Me.vcomentario = "Fact"
    
    If Me.RadioButton1.Value Then Me.vcomentario = ""
    
    
   ' If Not Trim(txtCliente.Tag) = "" Then
        FiltrarMovimientos
   ' Else
   '     MsgBox "No puede ejecutar el filtro si todavia no ha elegido un cliente!!!", vbExclamation, "Mensaje ..."
   ' End If
    
    TabControl2.Item(1).Selected = True
    
    If Me.Tag = "Proveedores" Then
        Me.lblsaltoTotal.Caption = "Saldo Actual: " + Format(getSaldoProveedor2(Me.txtCliente.Tag), "###,###,##0.00")
    Else
        Me.lblsaltoTotal.Caption = "Saldo Actual: " + Format(getSaldoCliente2(Me.txtCliente.Tag), "###,###,##0.00")
    End If
   
    
    If Err Then GrabarLog "cmdFiltroMovimientos_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdNoPagarFactura_Click()
On Error Resume Next

    With bFacturas
        If (.Recordset.EOF = True) Or (.Recordset.BOF = True) Then
            MsgBox "No tiene una Factura Seleccionada para deshacer el pago", vbInformation, "Mensaje ..."
            Exit Sub
        End If
    End With
    With bfdetalle
        .RecordSource = "SELECT * FROM fdetalle WHERE (remito = " & Val(bFacturas.Recordset("remito").Value) & ")"
        .Refresh

        Do Until .Recordset.EOF = True
            
            .Recordset("pagado").Value = "NO"
            .Recordset("resta").Value = .Recordset("totaliva").Value
            .Recordset("pago").Value = 0

            .Recordset.MoveNext
        Loop
    
    End With
    
    'CalcularSaldoFacturas
    
If Err Then GrabarLog "cmdNoPagarFactura_Click", Err.Number & " " & Err.Description, Me.Name
End Sub



Private Sub cmdVerDebitos_Click()
On Error Resume Next
    
    If (bFacturas.Recordset.BOF = True) Or (bFacturas.Recordset.EOF = True) Then Exit Sub
    'With frmCtaCteAgrupados
    '    .bctacte_agrupados.ConnectionString = pathDBMySQL
    '    .bctacte_agrupados.RecordSource = "select * from ctacte_agrupados where ctacte_padre = " & Trim(bfacturas.Recordset("remito").value)
    '    .bctacte_agrupados.Refresh
    '
    '    .Show
    'End With
    
If Err Then GrabarLog "cmdVerDebitos_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdVerMovimientosPM_Click()
On Error Resume Next
    
    'With frmMovimientosXmes
     
    '    .bctacte.ConnectionString = pathDBMySQL
        
        ' .bctacte.Refresh
        
        'If Not .bctacte.Recordset.EOF = True Then
        '    .bctacte.Recordset.MoveLast
        '    .Show
        'Else
        '    MsgBox "No tiene detalles para ver!!", vbInformation, "Mensaje ..."
        '    Unload frmMovimientosXmes
        'End If
        
        
    'End With
    
If Err Then GrabarLog "cmdVerMovimientosPM_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdVerRemitos_Click()
On Error Resume Next
    
    If (bFacturas.Recordset.BOF = True) Or (bFacturas.Recordset.EOF = True) Then Exit Sub

    'With frmFacturasAgrupadas
    '
    '    .bfactu_agrupadas.ConnectionString = pathDBMySQL
    '    .bfactu_agrupadas.RecordSource = "select * from factu_agrupadas where factura_padre = " + Str(bfacturas.Recordset("remito"))
    '    .bfactu_agrupadas.Refresh
    '
    '    If Not .bfactu_agrupadas.Recordset.EOF = True Then
    '        .bfactu_agrupadas.Recordset.MoveLast
    '        .Show
    '    Else
    '        MsgBox "No tiene detalles para ver!!", vbInformation, "Mensaje ..."
    '        Unload frmFacturasAgrupadas
    '    End If
    'End With
    
If Err Then GrabarLog "cmdVerRemitos_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdImprimirRecibo_Click()
    On Error Resume Next
    
    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox "Prepare la Impresora !", vbInformation, "Mensaje ..."
    
    With Mantenimiento.rscimpresión
        If .State = 0 Then .Open
        .Close
        .Open
    End With
    
    'With drctacteimpresion
    '    .Show
    'End With
    
    If Err Then GrabarLog "cmdImprimirRecibo_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdSaldosClientes_Click()
    On Error Resume Next
    
    'frmSaldosClientes.Show

    If Err Then GrabarLog "Command12_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdPagoCheques_Click()
    On Error Resume Next
    
    With frmCheques
        .dedonde = "ctacte"
        .TabCheques.tab = 0
        .txtNombre = Trim(txtCliente.Tag)
        .txtNombre_KeyPress 13
        docheque = "cheque"
    End With
    
    If Err Then GrabarLog "cmdPagoCheques_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdImprimir_Click(Index As Integer)
    On Error Resume Next
    
    MousePointer = vbHourglass
    
    Select Case Index
    
        Case 0
            
            If Me.rbRetenciones Then
                ImprimirRetMovimientos
            Else
                ImprimirFichaCliente
            End If
        
        Case 1
            ImprimirDetalle
    
    End Select
    
    
    MousePointer = vbDefault

    If Err Then GrabarLog "cmdImprimir_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub ImprimirDetalle()
On Error Resume Next
    
    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox "Prepare la Impresora ", vbInformation, "Mensaje ..."
    
    With Mantenimiento.rsCuentasCorrientesDetalle
        If .State = 1 Then .Close
        
        '.Source = "SHAPE {SELECT F.Remito, CC.id, CC.Fecha, CC.Codigo, CC.Nombre, CC.Debito, CC.Credito, CC.Saldo, CC.Comentario AS ComentarioCC, F.Comentario AS ComentarioF FROM CuentasCorrientes CC INNER JOIN Factura F ON CC.Remito = F.Remito WHERE CC.Codigo = '" & Trim(txtCliente.Tag) & "' GROUP BY CC.id ORDER BY Codigo;}  AS CuentasCorrientesDetalle APPEND ({SELECT FD.idFDetalle,FD.Remito, FD.Codigo, Cantidad, Detalle AS Descrip, FD.Precio, FD.Descuento, FD.Total, Comentario FROM FDetalle FD INNER JOIN Factura F ON FD.Remito = F.remito GROUP BY FD.idFDetalle;}  AS FDetalleDetalle RELATE 'Remito' TO 'Remito') AS FDetalleDetalle"

        .Source = "SHAPE {SELECT F.Remito, CC.id, CC.Fecha, CC.Codigo, CC.Nombre, CC.Debito, CC.Credito, CC.Saldo, CC.Comentario AS ComentarioCC, F.Comentario AS ComentarioF FROM CuentasCorrientes CC left join Factura F ON CC.Remito = F.Remito WHERE CC.Codigo = '" & Trim(txtCliente.Tag) & "' GROUP BY CC.id ORDER BY Codigo;}  AS CuentasCorrientesDetalle APPEND ({SELECT FD.idFDetalle,FD.Remito, FD.Codigo, Cantidad, Detalle AS Descrip, FD.Precio, FD.Descuento, FD.Total, Comentario FROM FDetalle FD left join Factura F ON FD.Remito = F.remito GROUP BY FD.idFDetalle;}  AS FDetalleDetalle RELATE 'Remito' TO 'Remito') AS FDetalleDetalle"
        
        
        If .State = 0 Then .Open
        .Close
        .Open
    
    End With
    'Err.Clear

    With drcuentascorrientes_detalles
        .Sections(2).Controls("lblCliente").Caption = Trim(txtCliente.Tag) & " - " & Trim(txtCliente.Text)
        .Sections(2).Controls("snombre").Caption = vDatosEmpresa.Nombre
        .Sections(2).Controls("sdirtel").Caption = vDatosEmpresa.Direccion & "  /  " & vDatosEmpresa.Telefono
        .Sections(2).Controls("slocalidad").Caption = vDatosEmpresa.Localidad
        .Sections(2).Controls("semail").Caption = vDatosEmpresa.Email

        .Sections("Sección3").Controls("lblSaldoCliente").Caption = lblSaldo(1).Caption
        .Refresh
        
        .Show
    End With
If Err Then GrabarLog "ImprimirDetalle", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cmdLimpiar_Click()
    On Error Resume Next
    
    txtCliente.Text = ""
    txtCuit.Text = ""

    txtCliente.Tag = ""
  
    lblSaldo(0).Caption = "0.000"
    lblSaldo(1).Caption = "0.000"
    
    txtCliente.SetFocus
    
    If Err Then GrabarLog "cmdLimpiar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Function CopiarDebito(vuremito As Long) As Boolean
On Error Resume Next
    
    Dim i As Integer
    
    i = 0
    
    Dim rsCtaCteDebito As New ADODB.Recordset, sqlCtaCteDebito As String
    
    sqlCtaCteDebito = "SELECT * FROM cuentascorrientes_debito WHERE 1=2"
    
    With rsCtaCteDebito
        Call .Open(sqlCtaCteDebito, ConnDDBB, adOpenStatic, adLockPessimistic)

        .AddNew
        
        .Fields("ctacte_padre").Value = vuremito
        
        For i = 1 To 18
            If Not IsNull(bFacturas.Recordset(i).Value) = True Then
                .Fields(i).Value = bFacturas.Recordset(i).Value
            End If
        Next
        
        .Update
    
    End With
    
    sqlCtaCteDebito = ""
    
    If rsCtaCteDebito.State = 1 Then
        rsCtaCteDebito.Close
        Set rsCtaCteDebito = Nothing
    End If
    
If Err.Number Then
    GrabarLog "CopiarDebito", Err.Number & " " & Err.Description, Me.Name
    CopiarDebito = True
Else
    CopiarDebito = False
End If
End Function
Function copiadoc(vnremito As Long, ByRef vuremito As Long) As Boolean
Dim i As Integer
On Error Resume Next

    If vuremito = 0 Then
        MsgBox "No se pueden agrupar los documentos!!!", vbCritical, "Error..."
        Exit Function
    End If

    i = 0
    With bremito
        .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM factura WHERE (remito = " & vnremito & ")"
        .Refresh
        If Not .Recordset.EOF = True Then
            If Not .Recordset("remito").Value > 0 Then
                MsgBox "No existe el Documento asociado a este débito.", vbInformation, "Mensaje ..."
                Exit Function
            End If
        End If
    End With
    
    'Cambio los nremito de Fdetalle
    With bfdetalle
        .RecordSource = "Select * from fdetalle where remito = " & vnremito & " order by Fecha"
        .Refresh
        If Not .Recordset.RecordCount = 0 Then .Recordset.MoveFirst
        Do Until .Recordset.EOF = True
            .Recordset("remito_ant").Value = vnremito 'Nº Remito que tenia previo a la factura
            .Recordset("remito").Value = vuremito     'Nuevo Nº de remito - El que trae la FACT
            
            ttotal_ctacte = ttotal_ctacte + .Recordset("totaliva").Value
            
            .Recordset.MoveNext
        Loop
        
    End With

    ' --------- estás son los documentos que corresponden a cada factura -------------------
    Dim rsFacturaRemito As New ADODB.Recordset, sqlFacturaRemito As String
    
    sqlFacturaRemito = "SELECT * FROM Factura_remito WHERE 1=2"
    
    With rsFacturaRemito
        Call .Open(sqlFacturaRemito, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        .AddNew
        
        .Fields("factura_padre").Value = vuremito
        For i = 0 To 25
            If Not IsNull(bremito.Recordset(i).Value) = True Then
                .Fields(i).Value = bremito.Recordset(i).Value
            End If
        Next
        
        .Update
    
    End With

    sqlFacturaRemito = ""
    
    If rsFacturaRemito.State = 1 Then
        rsFacturaRemito.Close
        Set rsFacturaRemito = Nothing
    End If
    '---------------------------------------------------------------------------------------
    
If Err.Number Then
    GrabarLog "copiadoc " & bremito.Recordset("remito").Value, Err.Number & " " & Err.Description, Me.Name
    copiadoc = True
Else
    copiadoc = False
End If
End Function
Private Sub credotorgado_Change()
    credotorgado = Format(credotorgado, "#####0.000")
End Sub

Private Sub VerDetalles()
    On Error Resume Next
    
    'Dim mouse As PointAPI
    
    With fraVerDetalles
        .Visible = True
        
      
    End With
    
    CargarDetalles
    
    If Err Then GrabarLog "VerDetalles ", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CargarDetalles()
    On Error Resume Next
    
    With bfdetalle
        .CursorLocation = adUseClient
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM fdetalle WHERE (remito = " & bFacturas.Recordset("Remito").Value & ")"
        .Refresh
    End With

    Set dgDetalles.DataSource = bfdetalle.Recordset
    
    Call FormatoGrillas(0)
    
    If Err Then GrabarLog "CargarDetalles ", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub FormatoGrillas(vGrilla As Integer)
    On Error Resume Next
    
    Select Case vGrilla
    
        Case 0
            With dgDetalles
                .HeadLines = 2
                .Columns(0).Width = 1000
                .Columns(1).Width = 0
                .Columns(2).Width = 750
                .Columns(3).Width = 750
                .Columns(4).Width = 7000
                .Columns(5).Width = 1000
                .Columns(6).Width = 0
                
                .Columns(13).Width = 0
                .Columns(21).Width = 500
                .Columns(22).Width = 500
                .Columns(23).Width = 0
                .Columns(24).Width = 0
            End With
        Case 1
        Case 2
        Case 3
        Case 4
        
    
    End Select
    
    If Err Then GrabarLog "FormatoGrillas ", Err.Number & " " & Err.Description, Me.Name
End Sub

Function ffechahasta(vtempfecha) As Date

    Select Case Mid(vtempfecha, 4, 2)
    
        Case "01"
            ffechahasta = "01/02/" + Right(vtempfecha, 4)

        Case "02"
            ffechahasta = "01/03/" + Right(vtempfecha, 4)

        Case "03"
            ffechahasta = "01/04/" + Right(vtempfecha, 4)

        Case "04"
            ffechahasta = "01/05/" + Right(vtempfecha, 4)

        Case "05"
            ffechahasta = "01/06/" + Right(vtempfecha, 4)

        Case "06"
            ffechahasta = "01/07/" + Right(vtempfecha, 4)

        Case "07"
            ffechahasta = "01/08/" + Right(vtempfecha, 4)

        Case "08"
            ffechahasta = "01/09/" + Right(vtempfecha, 4)

        Case "09"
            ffechahasta = "01/10/" + Right(vtempfecha, 4)

        Case "10"
            ffechahasta = "01/11/" + Right(vtempfecha, 4)

        Case "11"
            ffechahasta = "01/12/" + Right(vtempfecha, 4)

        Case "12"
            ffechahasta = "01/01/" + Trim(Str(Val(Right(vtempfecha, 4)) + 1))
    End Select

End Function

Function ftipoMovimiento(vtabla As String) As String
If Me.rbRetenciones Then
    ftipoMovimiento = vtabla + " t inner join retencionesmovimientos rm on (rm.nrointerno=t.nrointerno) " & _
    " inner join retenciones r on (r.idretencion=rm.idretenciones) "
Else
    ftipoMovimiento = vtabla
End If
End Function

Function fcampos() As String
If Me.rbRetenciones Then
    fcampos = "*,rm.`importe` as importeR "
Else
    fcampos = "*"
End If
End Function
Private Sub FiltrarMovimientos()
    On Error Resume Next
    
    Dim vsqlTipo As String
    
    vsqlTipo = ""
    
    Me.KlexCtaCte.Clear
    
    Dim sqlFiltro, vtabla As String
    
    If Me.Tag = "Proveedores" Then
        vtabla = "pcuentascorrientes"
        vid = "IdPcuentascorrientes"
    Else
        vtabla = "cuentascorrientes"
        vid = "id"
    End If
    'sqlFiltro = sqlFiltroGral + " where  (cuentascorrientes.codigo = '" & Trim(txtCliente.Tag) & "') and (cuentascorrientes.Fecha >= '" & strfechaMySQL(dtpDesde.Value) + "' and cuentascorrientes.fecha <= '" & strfechaMySQL(dtpHasta.Value) + "' and (cuentascorrientes.Comentario like '%" + Trim(txtComentario.Text) + "%') ) order by cuentascorrientes.fecha, cuentascorrientes.id"
    
    If Not vctipo = "" Then vsqlTipo = vsqlTipo + " and (tipomovimiento = '" + Me.vctipo + "') "
    
   vtabla = ftipoMovimiento(vtabla) ' ingreso las retencione
   
   
    If Me.rbDeuda.Value = True Then vsqlTipo = " and (debito > 0) "
    
    If Me.rbPago.Value = True Then vsqlTipo = " and (credito > 0) "

    
   If Me.txtCliente.Tag = "" Then

    
    If Val(Me.vcodRepartidor.Tag) > 0 Then
    
    
            Dim vCriterioVendedor As String
            
            vCriterioVendedor = " nrointerno in (select nrointerno from t_rel where idVendedor = " + Me.vcodRepartidor.Tag
            
            sqlFiltro = "SELECT " + fcampos + " FROM " + vtabla + " WHERE (comentario like '%" + Me.vcomentario + "%') and (Fecha >= '" & strfechaMySQL(dtpDesde.Value) + "' and fecha <= '" & strfechaMySQL(dtpHasta.Value) + "')" + vsqlTipo + _
            " and (" + vCriterioVendedor + ") " + _
            " order by fecha," + vid

            
    Else
       
            sqlFiltro = "SELECT " + fcampos + " FROM " + vtabla + " WHERE (comentario like '%" + Me.vcomentario + "%') and (Fecha >= '" & strfechaMySQL(dtpDesde.Value) + "' and fecha <= '" & strfechaMySQL(dtpHasta.Value) + "')" + vsqlTipo + " order by fecha," + vid
            vcriterio = vcriterio + "Comentario: " + vcomentario + " - " + vsqlTipo
                     
    End If
    
     
    
   Else
       ' sqlFiltro = "SELECT " + fcampos + " FROM " + vtabla + " WHERE (comentario like '%" + Me.vcomentario + "%') and  (codigo = '" & Trim(txtCliente.Tag) & "') and (Fecha >= '" & strfechaMySQL(dtpDesde.Value) + "' and fecha <= '" & strfechaMySQL(dtpHasta.Value) + "') order by fecha," + vid
       
       Dim vcomentario1 As String
       
       vcomentario1 = "  and (comentario like '%" + Me.vcomentario + "%')"
       
       sqlFiltro = "SELECT " + fcampos + " FROM " + vtabla + " WHERE (codigo = '" & Trim(txtCliente.Tag) & "') and (Fecha >= '" & strfechaMySQL(dtpDesde.Value) + "' and fecha <= '" & strfechaMySQL(dtpHasta.Value) + "')" _
       + vcomentario1 _
       + vsqlTipo + " order by fecha," + vid
    
    
    vcriterio = vcriterio + "Comentario: " + vcomentario + " - " + vsqlTipo
    
   End If
    
    
   'vcriterio = sqlFiltro
    
    
    
    
    If Me.vtipoProveedor = "" Or Me.vtipoProveedor.Text = "Todos" Then
      
    
    End If
    
    
    
    
    With rsCtaCte
        If .State = 1 Then .Close

        .CursorLocation = adUseServer
        Call .Open(sqlFiltro, ConnDDBB, adOpenDynamic, adLockPessimistic)
        
        
        vsaldoanterior = 0
        lblSaldo(0).Caption = Format(vsaldoanterior, "##,###,##0.000")
        
        vsaldoanterior = Val(CalcularSaldoAnterior(dtpDesde.Value)) 'Alfredo: en esta funciòn llena la grilla
        
        lblSaldo(0).Caption = Format(vsaldoanterior, "##,###,##0.000")
        
        If .EOF = True Then ' si no hay movimiento el saldo es idem al saldo anterior
            lblSaldo(1).Caption = lblSaldo(0).Caption
            Exit Sub
        End If
        
        'vsaldoanterior = Val(CalcularSaldoAnterior(dtpDesde.Value)) 'Alfredo: en esta funciòn llena la grilla
        
        'lblSaldo(0).Caption = Format(vsaldoanterior, "##,###,##0.000")
        
        
        ' llena grilla con el detalle
        
        If Me.rbRetenciones Then
            lblSaldo(1).Caption = Format(CalcularRetenciones, "#,###,##0.000")
        Else
            lblSaldo(1).Caption = Format(CalcularSaldo(Val(vsaldoanterior)), "#,###,##0.000")
        End If
        
        
    
    End With
    
    sqlFiltro = ""
    
    Me.KlexCtaCte.TopRow = Me.KlexCtaCte.Rows - 1
    Call LastKlexRow(Me.KlexCtaCte)
    
    If Err Then GrabarLog "FiltrarMovimientos", Err.Number & " " & Err.Description, Me.Name
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, _
                       Shift As Integer)
    On Error Resume Next

    If KeyCode = vbKeyF1 Then
        txtCliente.Text = ""
        txtCliente.Tag = ""
        txtCuit.Text = ""
        txtComentario.Text = ""
        
        dtpDesde.Value = Date - 90
        dtpHasta.Value = Date
    
        txtCliente.SetFocus
    
    End If
    
    If (KeyCode = 27) And (GBPagos.Visible = True) Then
        GBPagos.Visible = False
    End If
    vIdCtaCte = 0
    If Err Then GrabarLog "Form_KeyUp", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Load()
    On Error Resume Next

    'Set rec = New ADODB.Recordset
    'rec.Open "SELECT * FROM factura_ctacte", ConnDDBB, adOpenKeyset, adLockOptimistic
    '-----------------------------------------------------------
    'With bfdetalle
    '    .ConnectionString = pathDBMySQL
    '    .RecordSource = "SELECT * FROM fdetalle"
    '    .Refresh
    'End With

    '-----------------------------------------------------------
    
    dtpDesde.Value = Date - 90
    dtpHasta.Value = Date
 
    With Me
        .Show
        .Top = 0
        .Left = 0
        .Width = 16980
        .Height = 9495
        .KeyPreview = True
    End With
    
    i = 0
    
    Set rsCtaCte = New ADODB.Recordset
        
    If vidd > 0 Then
        
        With rsCtaCte
            .CursorLocation = adUseClient
            Call .Open("SELECT * FROM cuentascorrientes WHERE (id = " & Val(vidd) & ")", ConnDDBB, adOpenDynamic, adLockBatchOptimistic)
            If Not .State = 1 Then
                MsgBox Err.Description
                Exit Sub
            End If
        End With
    
    Else
    
    
    End If
    
    FormatoGrilla (1)
    
    txtCliente.SetFocus
    
    
    Call CentrarFormulario(Me)
    
    init
    
    

If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub

Public Sub init()


Me.vcomentario.Text = ""

If Me.Tag = "Proveedores" Then
    vctacteCP = "pcuentascorrientes"
    vid = "Idpcuentascorrientes"
    vfacturaCP = "pfactura"
    vcobrospagos = "pagos"
    vtablaCP = "proveedores"
    vfdetalleCP = "pfdetalle"

Else
    vctacteCP = "cuentascorrientes"
    vid = "id"
    vfacturaCP = "factura"
    vfdetalleCP = "fdetalle"
    vcobrospagos = "cobros"
    vtablaCP = "clientes"
End If

Me.WindowState = 0


Call CentrarFormulario(Me)
Me.Top = 0
Me.Width = 16980
Me.Height = 9495

Me.Caption = "Cuentas Corrientes de: " + Me.Tag



Me.vtipoProveedor.Clear
Me.vtipoProveedor.AddItem ("Proveedor")
Me.vtipoProveedor.AddItem ("Eventuales")
Me.vtipoProveedor.AddItem ("Personal")
Me.vtipoProveedor.AddItem ("Funcionarios")
Me.vtipoProveedor.AddItem ("Externos")
Me.vtipoProveedor.AddItem ("Rol1")
Me.vtipoProveedor.AddItem ("Rol2")

Me.vtipoProveedor.Text = "Proveedor"



Me.TabControl2.SelectedItem = 0

sqlFiltroGral = "SELECT " + vctacteCP + ".*, " + vfacturaCP + ".NComprobante From " + vctacteCP + " LEFT OUTER JOIN " + vfacturaCP + " ON (" + vctacteCP + ".Remito=" + vfacturaCP + ".Remito)"
valertaModulo = "->"
End Sub
Function BuscarEmpleado(vempleado As String, vestado As String) As String
    On Error Resume Next
    
    Dim rsempleados As New ADODB.Recordset, sqlEmpleados As String
            
    If vestado = "M" Then
        
        sqlEmpleados = "SELECT Codigo, Nombre FROM empleados WHERE (Codigo = '" + vempleado + "')"
    
    Else
        
        sqlEmpleados = "SELECT Codigo, Nombre FROM empleados WHERE (Nombre = '" + vempleado + "')"
    
    End If
    
    With rsempleados
        Call .Open(sqlEmpleados, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then
    
            If vestado = "M" Then
        
                BuscarEmpleado = .Fields("Nombre").Value
        
            Else
        
                BuscarEmpleado = .Fields("Codigo").Value
        
            End If
        
        End If
        
    End With

    sqlEmpleados = ""

    If rsempleados.State = 1 Then
        rsempleados.Close
        Set rsempleados = Nothing
    End If
    
    If Err Then GrabarLog "BuscarEmpleado", Err.Number & "-" & Err.Description, Me.Name
End Function
Function CalcularSaldoActual(vdfecha As Date, vhfecha As Date) As String
    On Error Resume Next
    
    Dim rsSaldosClientes As New ADODB.Recordset, sqlSaldosClientes As String
    
    sqlSaldosClientes = "SELECT Max(" + vctacteCP + ".Fecha) AS Fecha, " + vctacteCP + ".Codigo, Sum(" + vctacteCP + ".Debito) AS SumaDeDebito, Sum(" + vctacteCP + ".Credito) AS SumaDeCredito, (Sum(" + vctacteCP + ".Debito)-Sum(cuentascorrientes.Credito)) AS Saldo FROM " + vtablaCP + " LEFT JOIN " + vtablaCP + " ON " + vctacteCP + ".Codigo = " + vctacteCP + ".Codigo WHERE (((" + vctacteCP + ".FechaInput)>= '" & strfechaMySQL(vdfecha) + "' And (" + vctacteCP + ".FechaInput)<= '" & strfechaMySQL(vhfecha) + "') AND ((" + vctacteCP + ".Noimputar)= false) And ((" + vctacteCP + ".Codigo) = '" & Trim(txtCliente.Tag) & "')) GROUP BY " + vctacteCP + ".Codigo"
    
    With rsSaldosClientes
        Call .Open(sqlSaldosClientes, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then
            CalcularSaldoActual = Val(EsNulo(.Fields("Saldo").Value))
        Else
            CalcularSaldoActual = 0
        End If

    End With

    sqlSaldosClientes = ""

    If rsSaldosClientes.State = 1 Then
        rsSaldosClientes.Close
        Set rsSaldosClientes = Nothing
    End If
    
    If Err Then GrabarLog "CalcularSaldoActual", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub ImprimirRetMovimientos()
    On Error Resume Next

    If Not Me.rbRetenciones Then
        cambioFiltroParaReporte
    End If

    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox "   Prepare la impresora   ", vbInformation, "Mensaje ..."
    
    With Mantenimiento.rsRetMovimientos
        If Not .State = 0 Then .Close
        
        .Source = rsCtaCte.Source
            
        If Not .State = 1 Then .Open
        .Close
        .Open
    
    End With
    
    With drRetencionesMovimientos
        .Sections("TituloEmpresa").Controls("lblTitulo").Caption = "Detalle de Cuentas Corrientes " + Trim(vtablaCP)
        .Sections("TituloEmpresa").Controls("lblFechaDesde").Caption = dtpDesde.Value
        .Sections("TituloEmpresa").Controls("lblFechaHasta").Caption = dtpHasta.Value
        .Sections("TituloEmpresa").Controls("lblcliente").Caption = Trim(txtCliente.Tag) & " - " & Trim(txtCliente.Text)
        .Sections("TituloEmpresa").Controls("lblSaldoAnterior").Caption = "$ " & lblSaldo(0).Caption
        
        .Sections("PieInforme").Controls("lblSaldo").Caption = lblSaldo(1).Caption
        '.Refresh
        .Show
        
    End With
    
    If Err Then GrabarLog "ImprimirFichaCliente", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub ImprimirFichaCliente()
    On Error Resume Next

    If Not Me.rbRetenciones Then
        cambioFiltroParaReporte
    End If

    Unload Mantenimiento
    Load Mantenimiento
    
    
    Wait (1000)
    
  
    
    With Mantenimiento.rsccc
        If Not .State = 0 Then .Close
        
        .Source = rsCtaCte.Source
            
        If Not .State = 1 Then .Open
        .Close
        .Open
    
    End With
    
      MsgBox "   Prepare la impresora   ", vbInformation, "Mensaje ..."
      
    
    With drcuentascorrientes
    
        .Sections("TituloEmpresa").Controls("lblTitulo").Caption = "Detalle de Cuentas Corrientes " + Trim(vtablaCP)
        .Sections("TituloEmpresa").Controls("lblFechaDesde").Caption = dtpDesde.Value
        .Sections("TituloEmpresa").Controls("lblFechaHasta").Caption = dtpHasta.Value
        .Sections("TituloEmpresa").Controls("lblcliente").Caption = Trim(txtCliente.Tag) & " - " & Trim(txtCliente.Text)
         
        .Sections("TituloEmpresa").Controls("lblSaldoAnterior").Caption = "$ " & lblSaldo(0).Caption
        
        
        .Sections("TituloEmpresa").Controls("efiltro").Caption = vcriterio
   
        vcriterio = ""
        
        
        .Sections("PieInforme").Controls("lblSaldo").Caption = lblSaldo(1).Caption
        .Refresh
        .Show
        
    End With
    
    'If Err Then GrabarLog "ImprimirFichaCliente", Err.Number & " " & Err.Description, Me.Name
    If Err Then
        MsgBox Err.Description
    End If
        
End Sub
Private Sub cambioFiltroParaReporte()
Dim sqlFiltro As String
Dim vsqlTipo As String

vsqlTipo = ""



If Me.rbDeuda.Value = True Then vsqlTipo = " and (debito > 0) "
    
If Me.rbPago.Value = True Then vsqlTipo = " and (credito > 0) "
    
    
  ' sqlFiltro = sqlFiltroGral + " where  (cuentascorrientes.codigo = '" & Trim(txtCliente.Tag) & "') and (cuentascorrientes.Fecha >= '" & strfechaMySQL(dtpDesde.Value) + "' and cuentascorrientes.fecha <= '" & strfechaMySQL(dtpHasta.Value) + "' and (cuentascorrientes.Comentario like '%" + Trim(txtComentario.Text) + "%') ) order by cuentascorrientes.fecha, cuentascorrientes.id"
    
   If Me.txtCliente.Tag = "" Then
   
    sqlFiltro = sqlFiltroGral + " where  (" + vctacteCP + ".Fecha >= '" & strfechaMySQL(dtpDesde.Value) + "' and " + vctacteCP + ".fecha <= '" & strfechaMySQL(dtpHasta.Value) + "' and (" + vctacteCP + ".Comentario like '%" + Trim(vcomentario.Text) + "%') ) " + vsqlTipo + " order by " + vctacteCP + ".fecha, " + vctacteCP + "." + vid
   
   Else
   
    sqlFiltro = sqlFiltroGral + " where  (" + vctacteCP + ".codigo = '" & Trim(txtCliente.Tag) & "') and (" + vctacteCP + ".Fecha >= '" & strfechaMySQL(dtpDesde.Value) + "' and " + vctacteCP + ".fecha <= '" & strfechaMySQL(dtpHasta.Value) + "' and (" + vctacteCP + ".Comentario like '%" + Trim(vcomentario.Text) + "%') ) order by " + vctacteCP + ".fecha, " + vctacteCP + "." + vid
    
   End If
    
   'sqlFiltro = "SELECT * FROM cuentascorrientes  WHERE (codigo = '" & Trim(txtCliente.Tag) & "') and (Fecha >= '" & strfechaMySQL(dtpDesde.Value) + "' and fecha <= '" & strfechaMySQL(dtpHasta.Value) + "' and (Comentario like '%" + Trim(txtComentario.Text) + "%') ) order by fecha, id"
    
    
    With rsCtaCte
        If .State = 1 Then .Close

        .CursorLocation = adUseServer
        Call .Open(sqlFiltro, ConnDDBB, adOpenDynamic, adLockPessimistic)
    End With
End Sub
Private Sub imprime_pagoporfactura()
    On Error Resume Next
        
    With Mantenimiento.rsfactura_ctacte

        If Not .State = 1 Then .Open
        .Close
        .Open

        .Filter = "(codigo = '" & Trim(txtCliente.Tag) & "') and (remito <> 0) and ((Fecha >= '" + strfecha2(dtpDesde.Value) + "' and Fecha <= '" + strfecha2(dtpHasta.Value) + "'))"
        .Sort = "Remito ASC, Fecha ASC, Id ASC"
    End With

    'With drPagoPorFactura
    '    .Sections("TituloEmpresa").Controls("vcliente").Caption = Trim(txtCliente.Tag) & "-" & Trim(txtCliente.Text)
    '    .Show
    'End With

    If Err Then GrabarLog "Imprime_pagopormes", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub imprime_pagopormes()
    On Error Resume Next

    With Mantenimiento.rspagopormes

        If .State = 1 Then
            .Close
            .Open
        Else
            .Open
            .Close
            .Open
        End If

        .Filter = "Codigo = '" & Trim(txtCliente.Tag) & "' and Últimodefecha <= '" + strfecha2(dtpHasta) + "' and Últimodefecha >= '" + strfecha2(dtpDesde) + "'"
        .Sort = "Anomes ASC"
    End With

    'With drctactepagopormes
    '    .Sections(2).Controls("vcliente").Caption = Trim(txtCliente.Tag) & " - " & Trim(txtCliente.Text)
    '    .Sections(4).Controls("vsaldo").Caption = saldo.Caption
    '    .Show
    'End With
    
    If Err Then GrabarLog "Imprime_pagofactura", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub saldo_Change()
    On Error Resume Next
    
    saldo.Caption = Format(Val(saldo.Caption), "#,###,##0.000")
    lblSaldo(0).Caption = Format(Val(saldo.Caption) + Val(saldocheque) + Val(credotorgado.Caption), "#,###,##0.000")

    If Err Then GrabarLog "saldo_Change", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub saldocheque_Change()
    On Error Resume Next
    
    saldocheque = Format(saldocheque, "#####0.000")
    'lblSaldo.Caption = Format(Val(saldo.Caption) + Val(saldocheque) + Val(credotorgado.Caption), "######0.000")

    If Err Then GrabarLog "saldocheque_Change", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub txtCliente_Change()
On Error Resume Next

    If txtCliente.Text = "" Then
        txtCliente.Tag = ""
    Else
        Call MostrarCoincidencias(txtCliente.Text)
    End If
    
    Me.TabControl2.SelectedItem = 0
    
If Err Then GrabarLog "txtCliente_Change", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CargoDatosClientes()
On Error Resume Next

    If BuscarCliente("Normal") = True Then
        vSaldoCliente = 0
        If chkSolo_Saldo.Value = 1 Then
            lblSaldo(0).Caption = Format((CalcularSaldoActual("2000-01-01", Date)), "###,###,###0.000")
        Else
            ActualizarVistas
            'rsaldo.Caption = Format(CalcularSaldoActual(Me.fdesde.Value, Date), "######0.000")
            'asaldo.Caption = Format(CalcularSaldoAnterior(Date - Left(Date, 2)), "######0.000")
        End If
            
        'Set lblSaldo(.DataSource = Nothing
        'Set saldo.DataSource = Nothing
    
        If rsCtaCte.EOF = True Then
            cmdVerMovimientos.Enabled = False
        Else
            cmdVerMovimientos.Enabled = True
        End If
        
    End If

    Me.KlexCtaCte.TopRow = Me.KlexCtaCte.Rows - 1

If Err Then GrabarLog "CargoDatosClientes", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CargarMovimientosCtaCte()
On Error Resume Next

If Err Then GrabarLog "CargarMovimientosCtaCte", Err.Number & " " & Err.Description, Me.Caption
End Sub
Function FuncionMes(vMesActual) As String
On Error Resume Next
    
    Select Case vMesActual
    
        Case 1
            FuncionMes = "12/" & Year(Date) - 1
            
        Case 2, 3, 4, 5, 6, 7, 8, 9, 10, 11
            FuncionMes = vMesActual - 1
    
    End Select

If Err Then GrabarLog "FuncionMes", Err.Number & " " & Err.Description, Me.Name
End Function

Private Sub vdtipo_Change()
    vctipo.Text = vdtipo.Tag
End Sub

Private Sub vsanterior_Change()
    On Error Resume Next
    
    dsaldo.Caption = "Saldo anterior a la fecha " & (dtpDesde.Value) & " :"

    If Err Then GrabarLog "vsanterior_Change", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub MostrarCoincidencias(vBusqueda As String)
On Error Resume Next
    
    Dim sqlClientes As String
    
    Set rsClientes = New ADODB.Recordset
    
    If Trim(vBusqueda) = "" Then
        sqlClientes = "SELECT * FROM " + vtablaCP + " WHERE 1=2"
    Else
        
        If Val(vBusqueda) > 0 Then
        
                sqlClientes = "SELECT * FROM " + vtablaCP + "  WHERE (Codigo = '" & Trim(vBusqueda) & "')"
        Else
        
            sqlClientes = "SELECT * FROM " + vtablaCP + "  WHERE (Codigo LIKE '%" & Trim(vBusqueda) & "%') OR (Nombre LIKE '%" & Trim(vBusqueda) & "%')"
    
        End If
    End If
    
    With rsClientes
        If .State = 1 Then .Close
    
        .CursorLocation = adUseClient
    
        Call .Open(sqlClientes, ConnDDBB, adOpenStatic, adLockReadOnly)
    
        dgClientes.Visible = Not .EOF
    
        If Not .EOF = True Then
            Set dgClientes.DataSource = rsClientes
            Call FormatoGrillaClientes
        Else
            Set dgClientes.DataSource = Nothing
        End If
    
    End With
    
    sqlClientes = ""
    
If Err Then GrabarLog "MostrarCoincidencias", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub FormatoGrillaClientes()
On Error Resume Next
    
    Dim i As Integer
    
    With dgClientes
    
        .ZOrder (0)
        '.Top = txtCliente.Top + txtCliente.Height + 500
        '.Left = txtCliente.Left + 500
    
        .HeadLines = 1.2
    
        For i = 0 To .Columns.Count - 1
    
            Select Case i
    
                Case 3
                    .Columns("Nombre").Width = .Width - 1000
                
                Case Else
                    .Columns(i).Width = 0
            
            End Select
        Next
    
    End With

    
If Err Then GrabarLog "FormatoGrillaClientes", Err.Number & " " & Err.Description, Me.Name
End Sub
