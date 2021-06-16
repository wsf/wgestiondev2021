VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{9746E3DA-06E1-4D26-9CE4-D9F6411A9C70}#1.0#0"; "SMGA_OcxTxt2008.ocx"
Begin VB.Form frmAsientos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Asientos"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   13035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   13035
   Begin XtremeSuiteControls.GroupBox GroupBox4 
      Height          =   165
      Left            =   -30
      TabIndex        =   50
      Top             =   240
      Width           =   13005
      _Version        =   851968
      _ExtentX        =   22939
      _ExtentY        =   291
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
   End
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   525
      Left            =   0
      TabIndex        =   48
      Top             =   -90
      Width           =   13005
      _Version        =   851968
      _ExtentX        =   22939
      _ExtentY        =   926
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   6
         Left            =   11550
         TabIndex        =   49
         Top             =   120
         Visible         =   0   'False
         Width           =   1395
         _Version        =   851968
         _ExtentX        =   2461
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Cerrar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmAsientos.frx":0000
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   525
      Left            =   30
      TabIndex        =   45
      Top             =   3900
      Width           =   12945
      _Version        =   851968
      _ExtentX        =   22834
      _ExtentY        =   926
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton PusModificarAsiento 
         Height          =   345
         Left            =   60
         TabIndex        =   46
         Top             =   150
         Width           =   1665
         _Version        =   851968
         _ExtentX        =   2937
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Modificar Asiento"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmAsientos.frx":0400
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   5
         Left            =   1740
         TabIndex        =   47
         Top             =   150
         Width           =   3975
         _Version        =   851968
         _ExtentX        =   7011
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Ver movimientos de otros módulos relacionados"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmAsientos.frx":099A
         BorderGap       =   10
      End
   End
   Begin VB.OptionButton vopprovee 
      BackColor       =   &H80000004&
      Caption         =   "Proveedores:"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   4890
      TabIndex        =   34
      Top             =   4860
      Width           =   1245
   End
   Begin VB.OptionButton vopcli 
      BackColor       =   &H80000004&
      Caption         =   "Cliente:"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   4020
      TabIndex        =   33
      Top             =   4860
      Value           =   -1  'True
      Width           =   825
   End
   Begin XtremeSuiteControls.PushButton PusModificarLey 
      Height          =   315
      Left            =   11340
      TabIndex        =   29
      Top             =   4560
      Width           =   1545
      _Version        =   851968
      _ExtentX        =   2725
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Modificar Leyenda"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   3075
      Left            =   30
      TabIndex        =   25
      Top             =   5280
      Width           =   12975
      _Version        =   851968
      _ExtentX        =   22886
      _ExtentY        =   5424
      _StockProps     =   68
      Appearance      =   8
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      ItemCount       =   1
      Item(0).Caption =   "Datos del Asiento seleccionado"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "dgDetalle"
      Begin MSDataGridLib.DataGrid dgDetalle 
         Height          =   2475
         Left            =   60
         TabIndex        =   26
         Top             =   360
         Width           =   12885
         _ExtentX        =   22728
         _ExtentY        =   4366
         _Version        =   393216
         BackColor       =   16777215
         BorderStyle     =   0
         ForeColor       =   8388608
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
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1575
      Left            =   30
      TabIndex        =   22
      Top             =   2310
      Width           =   12975
      _Version        =   851968
      _ExtentX        =   22886
      _ExtentY        =   2778
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin MSDataGridLib.DataGrid dgAsientos 
         Height          =   1395
         Left            =   30
         TabIndex        =   23
         Top             =   120
         Width           =   12825
         _ExtentX        =   22622
         _ExtentY        =   2461
         _Version        =   393216
         BorderStyle     =   0
         ForeColor       =   4210752
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
   End
   Begin XtremeSuiteControls.TabControl TabAsientos 
      Height          =   1455
      Left            =   30
      TabIndex        =   0
      Top             =   450
      Width           =   12975
      _Version        =   851968
      _ExtentX        =   22886
      _ExtentY        =   2566
      _StockProps     =   68
      AllowReorder    =   -1  'True
      PaintManager.BoldSelected=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      ItemCount       =   3
      Item(0).Caption =   "Busqueda de Asientos"
      Item(0).ControlCount=   24
      Item(0).Control(0)=   "txtBusqueda(3)"
      Item(0).Control(1)=   "lblBusqueda(0)"
      Item(0).Control(2)=   "lblBusqueda(1)"
      Item(0).Control(3)=   "lblBusqueda(2)"
      Item(0).Control(4)=   "lblBusqueda(3)"
      Item(0).Control(5)=   "lblBusqueda(4)"
      Item(0).Control(6)=   "lblBusqueda(5)"
      Item(0).Control(7)=   "lblBusqueda(6)"
      Item(0).Control(8)=   "txtBusqueda(4)"
      Item(0).Control(9)=   "txtBusqueda(0)"
      Item(0).Control(10)=   "txtBusqueda(1)"
      Item(0).Control(11)=   "txtBusqueda(2)"
      Item(0).Control(12)=   "dtpFecha(0)"
      Item(0).Control(13)=   "dtpFecha(1)"
      Item(0).Control(14)=   "PbAcciones(0)"
      Item(0).Control(15)=   "PbAcciones(1)"
      Item(0).Control(16)=   "PbAcciones(3)"
      Item(0).Control(17)=   "PbAcciones(2)"
      Item(0).Control(18)=   "lblBusqueda(7)"
      Item(0).Control(19)=   "lblBusqueda(8)"
      Item(0).Control(20)=   "lblBusqueda(9)"
      Item(0).Control(21)=   "vTxtnrobalance"
      Item(0).Control(22)=   "vdescripcion"
      Item(0).Control(23)=   "vperiodo"
      Item(1).Caption =   "Impresion de Asientos"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "PBImpresion(1)"
      Item(1).Control(1)=   "PBImpresion(2)"
      Item(1).Control(2)=   "PBImpresion(0)"
      Item(2).Caption =   "Acciones"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "log"
      Begin XtremeSuiteControls.FlatEdit vTxtnrobalance 
         Height          =   285
         Left            =   10320
         TabIndex        =   42
         Top             =   390
         Width           =   2595
         _Version        =   851968
         _ExtentX        =   4577
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin VB.ListBox log 
         Height          =   840
         Left            =   -69970
         TabIndex        =   27
         Top             =   390
         Visible         =   0   'False
         Width           =   12915
      End
      Begin XtremeSuiteControls.PushButton PBImpresion 
         Height          =   435
         Index           =   0
         Left            =   -69820
         TabIndex        =   9
         Top             =   630
         Visible         =   0   'False
         Width           =   2625
         _Version        =   851968
         _ExtentX        =   4630
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Imprimir Listado Actual"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmAsientos.frx":0F34
      End
      Begin XtremeSuiteControls.PushButton PBImpresion 
         Height          =   435
         Index           =   1
         Left            =   -67150
         TabIndex        =   8
         Top             =   630
         Visible         =   0   'False
         Width           =   2955
         _Version        =   851968
         _ExtentX        =   5212
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Imprimir Asientos Con Movimientos"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmAsientos.frx":1394
      End
      Begin XtremeSuiteControls.PushButton PBImpresion 
         Height          =   435
         Index           =   2
         Left            =   -64180
         TabIndex        =   10
         Top             =   630
         Visible         =   0   'False
         Width           =   2775
         _Version        =   851968
         _ExtentX        =   4895
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Imprimir Asiento Seleccionado"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmAsientos.frx":17F4
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   0
         Left            =   -69850
         TabIndex        =   11
         Top             =   480
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Nuevo Asiento"
         Appearance      =   3
         Picture         =   "frmAsientos.frx":1C54
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   1
         Left            =   -68140
         TabIndex        =   12
         Top             =   480
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Borrar Asiento"
         Appearance      =   3
         Picture         =   "frmAsientos.frx":206D
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   375
         Index           =   3
         Left            =   -68140
         TabIndex        =   13
         Top             =   870
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Ver Movimientos"
         Appearance      =   3
         Picture         =   "frmAsientos.frx":2513
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   375
         Index           =   2
         Left            =   -69850
         TabIndex        =   14
         Top             =   870
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Editar Asiento"
         Enabled         =   0   'False
         Appearance      =   3
         Picture         =   "frmAsientos.frx":29D2
      End
      Begin XtremeSuiteControls.FlatEdit txtBusqueda 
         Height          =   285
         Index           =   0
         Left            =   4920
         TabIndex        =   17
         Top             =   420
         Width           =   1005
         _Version        =   851968
         _ExtentX        =   1764
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtBusqueda 
         Height          =   285
         Index           =   1
         Left            =   4920
         TabIndex        =   18
         Top             =   720
         Width           =   1005
         _Version        =   851968
         _ExtentX        =   1764
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtBusqueda 
         Height          =   285
         Index           =   2
         Left            =   8010
         TabIndex        =   19
         Top             =   420
         Width           =   1005
         _Version        =   851968
         _ExtentX        =   1764
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtBusqueda 
         Height          =   285
         Index           =   3
         Left            =   8010
         TabIndex        =   20
         Top             =   720
         Width           =   1005
         _Version        =   851968
         _ExtentX        =   1764
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtBusqueda 
         Height          =   285
         Index           =   4
         Left            =   1320
         TabIndex        =   21
         Top             =   1020
         Width           =   7695
         _Version        =   851968
         _ExtentX        =   13573
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
      End
      Begin Aplisoft_CajasDeTexto.TxF dtpFecha 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   15
         Top             =   420
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
         Left            =   1320
         TabIndex        =   16
         Top             =   720
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
      Begin VB.Label vperiodo 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   10320
         TabIndex        =   44
         Top             =   1020
         Width           =   2595
      End
      Begin VB.Label vdescripcion 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   10320
         TabIndex        =   43
         Top             =   690
         Width           =   2595
      End
      Begin VB.Label lblBusqueda 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> Período:"
         Height          =   195
         Index           =   9
         Left            =   8940
         TabIndex        =   41
         Top             =   1080
         Width           =   1305
      End
      Begin VB.Label lblBusqueda 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> Nº Balance:"
         Height          =   165
         Index           =   8
         Left            =   9210
         TabIndex        =   40
         Top             =   450
         Width           =   1035
      End
      Begin VB.Label lblBusqueda 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> Descripción:"
         Height          =   195
         Index           =   7
         Left            =   9030
         TabIndex        =   39
         Top             =   765
         Width           =   1215
      End
      Begin VB.Label lblBusqueda 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> Leyenda:"
         Height          =   195
         Index           =   6
         Left            =   60
         TabIndex        =   7
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label lblBusqueda 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> Nº Int. Hasta:"
         Height          =   255
         Index           =   5
         Left            =   6615
         TabIndex        =   6
         Top             =   735
         Width           =   1305
      End
      Begin VB.Label lblBusqueda 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> Nº Int. Desde:"
         Height          =   195
         Index           =   4
         Left            =   6615
         TabIndex        =   5
         Top             =   480
         Width           =   1305
      End
      Begin VB.Label lblBusqueda 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> Nº Asiento Hasta:"
         Height          =   195
         Index           =   3
         Left            =   3360
         TabIndex        =   4
         Top             =   735
         Width           =   1455
      End
      Begin VB.Label lblBusqueda 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> Nº Asiento Desde:"
         Height          =   195
         Index           =   2
         Left            =   3360
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblBusqueda 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> F. Hasta :"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   2
         Top             =   735
         Width           =   1100
      End
      Begin VB.Label lblBusqueda 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> F. Desde :"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   1
         Top             =   480
         Width           =   1100
      End
   End
   Begin XtremeSuiteControls.PushButton PbAcciones 
      Height          =   345
      Index           =   4
      Left            =   30
      TabIndex        =   24
      Top             =   1950
      Width           =   12945
      _Version        =   851968
      _ExtentX        =   22834
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "&Filtrar Asientos "
      UseVisualStyle  =   -1  'True
      TextAlignment   =   6
      Picture         =   "frmAsientos.frx":2F1C
   End
   Begin XtremeSuiteControls.FlatEdit vleyenda 
      Height          =   285
      Left            =   6240
      TabIndex        =   28
      Top             =   4560
      Width           =   4815
      _Version        =   851968
      _ExtentX        =   8493
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   285
      Left            =   11340
      TabIndex        =   31
      Top             =   4890
      Width           =   1545
      _Version        =   851968
      _ExtentX        =   2725
      _ExtentY        =   503
      _StockProps     =   79
      Caption         =   "Modificar Persona"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit vcp1 
      Height          =   285
      Left            =   6240
      TabIndex        =   32
      Top             =   4860
      Width           =   4815
      _Version        =   851968
      _ExtentX        =   8493
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.PushButton PusModificarFecha 
      Height          =   315
      Left            =   2430
      TabIndex        =   38
      Top             =   4710
      Width           =   1425
      _Version        =   851968
      _ExtentX        =   2514
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Modificar Fecha"
      UseVisualStyle  =   -1  'True
   End
   Begin Aplisoft_CajasDeTexto.TxF vfechaAsiento 
      Height          =   285
      Left            =   660
      TabIndex        =   36
      Top             =   4710
      Width           =   1665
      _ExtentX        =   2937
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
   Begin VB.Label lblFecha 
      BackColor       =   &H80000004&
      Caption         =   "Fecha:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   60
      TabIndex        =   37
      Top             =   4770
      Width           =   585
   End
   Begin VB.Label lblLeyenda 
      BackColor       =   &H80000004&
      Caption         =   "Leyenda:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   5430
      TabIndex        =   35
      Top             =   4590
      Width           =   705
   End
   Begin VB.Label lblNuevaLeyenda 
      Alignment       =   1  'Right Justify
      Caption         =   "Nueva Leyenda:"
      Height          =   195
      Left            =   870
      TabIndex        =   30
      Top             =   4650
      Width           =   1245
   End
End
Attribute VB_Name = "frmAsientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsAsientos As ADODB.Recordset, rsAsientosDetalle As ADODB.Recordset
Dim vnumeroasiento As Long
Dim vnrobalance, vnrointerno As Long
Dim vidAsientos As Long

Private Sub VerDetalle()
On Error Resume Next
    


    With rsAsientos
        If Not .EOF = True And Not .BOF = True Then
        
            ' rodri: acá están todos los datos del encabezado del asient puesto en variables
            vnumeroasiento = .Fields("Numero").Value
            vleyenda.Text = .Fields("leyenda").Value
            vTxtnrobalance = .Fields("NroBalance").Value
            Me.vfechaAsiento = .Fields("fecha").Value
            vnrointerno = .Fields("nrointerno").Value
            vidAsientos = .Fields("idAsientos").Value
            
            vcp1.Text = "" ' pongo vacio el label
            
            If Not Trim(EsNulo(.Fields("CodigoCliente"))) = "" Then
                 vcp1 = .Fields("CodigoCliente")
                 Me.vopcli.Value = True
            End If
            
            
            If Not Trim(EsNulo(.Fields("CodigoProveedor"))) = "" Then
                 Me.vcp1.Text = .Fields("CodigoProveedor")
                 Me.vopprovee.Value = True
            End If
            
            
            Call CargarDetalleAsiento(.Fields("Numero").Value, .Fields("TimeStamp"), .Fields("nrobalance").Value)
        End If
    End With
    
If Err Then GrabarLog "VerDetalle", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub dgAsientos_Click()
On Error Resume Next
    
    VerDetalle

If Err Then GrabarLog "dgAsientos_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub dgDetalle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
End If
End Sub

Private Sub dtpFecha_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
    
    
    If KeyAscii = 13 Then
    
        Select Case Index
    
            Case 0
                dtpFecha(1).SetFocus
            Case 1
                txtBusqueda(0).SetFocus
        
        End Select

    End If
    
If Err Then GrabarLog "dtpFecha_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub Form_Load()
On Error Resume Next
    
    Me.Show
    Call CentrarFormulario(Me)
    
    init
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub init()


    
vnrobalance = TraerDato("balances", " Activo='S' order by NroBalance Desc", "NroBalance", pathDBMySQL)


If vnrobalance = 15 Then

    PusModificarFecha.Enabled = False
    PusModificarLey.Enabled = False
    PushButton1.Enabled = False

Else

    PusModificarFecha.Enable = True
    PusModificarLey.Enabled = True
    PushButton1.Enabled = True
End If

End Sub
Private Sub CargarAsientos(ByRef vsql As String)
On Error Resume Next

    Set rsAsientos = New ADODB.Recordset
    Dim sqlAsientos As String
    
    sqlAsientos = "SELECT  A.idAsientos, A.Fecha, A.Numero, A.NroInterno, SUM(Debe), SUM(Haber), Leyenda, A.NroBalance, A.CodigoCliente, A.CodigoProveedor, A.TimeStamp FROM Asientos A INNER JOIN AsientosDetalle AD ON A.Numero = AD.Numero WHERE 1=1 " & vsql & " GROUP BY A.Numero"
    
    With rsAsientos
        If .State = 1 Then .Close
        
        .CursorLocation = adUseClient
        
        Call .Open(sqlAsientos, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If .State = 1 Then
            
            FormatoGrilla (0)
            
            If Not .EOF = True Then .MoveLast
        Else
        
        
        End If
    
    End With

If Err Then GrabarLog "CargarAsientos", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub FormatoGrilla(Index As Integer)
On Error Resume Next

    If Index = 0 Then
    
        With dgAsientos
            Set .DataSource = Nothing
            Set .DataSource = rsAsientos
        
            .HeadLines = 1.5
        
            .Columns(0).Width = 0
            .Columns(1).Width = 1000
            .Columns(1).DataFormat.Format = "DD/MM/YYYY"
            
            .Columns(2).Width = 850
            .Columns(2).Caption = "Nº Asiento"
            .Columns(2).Alignment = dbgRight
            
            .Columns(3).Width = 850
            .Columns(3).Caption = "Nº Int."
            .Columns(3).Alignment = dbgRight
            
            .Columns(4).Width = 1000
            .Columns(4).Caption = "Debe"
            .Columns(4).DataFormat.Format = "$ #######0.00"
            .Columns(4).Alignment = dbgRight
            
            .Columns(5).Width = 1000
            .Columns(5).Caption = "Haber"
            .Columns(5).DataFormat.Format = "$ #######0.00"
            .Columns(5).Alignment = dbgRight
            
            .Columns(6).Width = 3500
        
        End With
    
    Else
    
        With dgDetalle
            Set .DataSource = rsAsientosDetalle  ' Rodri: se pasa la table delos detalles del asiento a la grilla
        
            .HeadLines = 1.5
        
            .Columns(0).Width = 0
            .Columns(1).Width = 0
            
            .Columns(2).Width = 2250
            .Columns(2).Caption = "Cod. Cuenta"
            .Columns(2).Alignment = dbgLeft
            
            .Columns(3).Width = 4500
            .Columns(3).Caption = "Nombre de la Cuenta"
            .Columns(3).Alignment = dbgLeft
            
            .Columns(4).Width = 1500
            .Columns(4).Caption = "Debe"
            .Columns(4).DataFormat.Format = "$ #######0.00"
            .Columns(4).Alignment = dbgRight
            
            .Columns(5).Width = 1500
            .Columns(5).Caption = "Haber"
            .Columns(5).DataFormat.Format = "$ #######0.00"
            .Columns(5).Alignment = dbgRight
            
            
        End With
    
    End If
    
If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CargarDetalleAsiento(vnumeroasiento As Long, vtimestamp As Date, vnrobalance As Integer)
On Error Resume Next

    Set rsAsientosDetalle = New ADODB.Recordset
    Dim sqlAsientosDetalle As String
    
   ' sqlAsientosDetalle = "SELECT idAsientosDetalle, AsientosDetalle.Numero, AsientosDetalle.CodigoCuenta, Cuenta, Debe, Haber FROM AsientosDetalle INNER JOIN Cuentas ON (AsientosDetalle.CodigoCuenta = Cuentas.CodigoCuenta)  WHERE (date(AsientosDetalle.TIMESTAMP) >= '" + Format(vtimestamp, "yyyy-mm-dd") + "') and (AsientosDetalle.Numero = " & vnumeroasiento & ") and (AsientosDetalle.nrobalance=" + Str(vnrobalance) + ")  ORDER BY idAsientosDetalle"
    
  '  If vnrobalance = 15 Then
    
       ' sqlAsientosDetalle = "SELECT idAsientosDetalle, AsientosDetalle.Numero, AsientosDetalle.CodigoCuenta, Cuenta, Debe, Haber FROM AsientosDetalle INNER JOIN Cuentas ON (AsientosDetalle.CodigoCuenta = Cuentas.CodigoCuenta)  WHERE (date(AsientosDetalle.TIMESTAMP) >= '" + Format(vtimestamp, "yyyy-mm-dd") + "') and (AsientosDetalle.Numero = " & vnumeroasiento & ")  ORDER BY idAsientosDetalle"
  '      sqlAsientosDetalle = "SELECT idAsientosDetalle, AsientosDetalle.Numero, AsientosDetalle.CodigoCuenta, Cuenta, Debe, Haber FROM AsientosDetalle INNER JOIN Cuentas ON (AsientosDetalle.CodigoCuenta = Cuentas.CodigoCuenta)  WHERE asientosdetalle.`idAsientosDetalle` >= 87246 and (AsientosDetalle.Numero = " & vnumeroasiento & ")  ORDER BY idAsientosDetalle"
          
  '  Else
    
        'sqlAsientosDetalle = "SELECT idAsientosDetalle, AsientosDetalle.Numero, AsientosDetalle.CodigoCuenta, Cuenta, Debe, Haber FROM AsientosDetalle INNER JOIN Cuentas ON (AsientosDetalle.CodigoCuenta = Cuentas.CodigoCuenta)  WHERE (date(AsientosDetalle.TIMESTAMP) >= '" + Format(vtimestamp, "yyyy-mm-dd") + "') and (AsientosDetalle.Numero = " & vnumeroasiento & ")  ORDER BY idAsientosDetalle"
        sqlAsientosDetalle = "SELECT idAsientosDetalle, AsientosDetalle.Numero, AsientosDetalle.CodigoCuenta, Cuenta, Debe, Haber FROM AsientosDetalle INNER JOIN Cuentas ON (AsientosDetalle.CodigoCuenta = Cuentas.CodigoCuenta)  WHERE (AsientosDetalle.nrobalance=" + Str(vnrobalance) + ") and (AsientosDetalle.Numero = " & vnumeroasiento & ")  ORDER BY idAsientosDetalle"
    
   ' End If
    
    
    
    With rsAsientosDetalle
        If .State = 1 Then .Close
        
        .CursorLocation = adUseClient
        
        Call .Open(sqlAsientosDetalle, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If .State = 1 Then
            FormatoGrilla (1)
        End If
    
    End With

If Err Then GrabarLog "CargarDetalleAsiento", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub

Public Sub PbAcciones_Click(Index As Integer)
On Error Resume Next
Dim vcborrraasiento As String
vcborrraasiento = Chr(13) + "En esta versión debe borrar manualmente (si corresponde) los movimientos de:" + Chr(13) + _
" - CtaCte Cliente/Provedor" + Chr(13) + _
" - Doc. Venta/Compra" + Chr(13) + _
" - Cheques, etc"
 

log.Clear


    
    Select Case Index
    
        Case 0
            frmAsientosAlta.Show
            
        Case 1
            With rsAsientos
                If Not .EOF = True And Not .BOF = True Then
                    If Not MsgBox("¿Esta seguro que desea borrar el Asiento Nº " & EsNulo(.Fields("Numero").Value) & "?", vbInformation + vbYesNo) = vbYes Then
                        Exit Sub
                    End If
                    
                    PbAcciones(Index).Tag = EsNulo(.Fields("NroInterno").Value)
                    
                    If Not Val(PbAcciones(Index).Tag) = 0 Then
                    
                        log.AddItem "Borrando detalle del asiento ... (ok!)"
                        Call BorrarBase("AsientosDetalle WHERE Numero = " & Val(.Fields("Numero").Value) & "", pathDBMySQL)
                        
                        log.AddItem "Borrando asiento ... (ok!)"
                        Call BorrarBase("Asientos WHERE Numero = " & Val(.Fields("Numero").Value) & "", pathDBMySQL)
                                    
                        log.AddItem "Borrando ctacte cliente ... (ok!)"
                        Call BorrarBase("CuentasCorrientes WHERE (NroInterno = " & Val(PbAcciones(Index).Tag) & ")", pathDBMySQL)
                        
                        log.AddItem "Borrando doc. venta ... (ok!)"
                        Call BorrarBase("Factura WHERE (NroInterno = " & Val(PbAcciones(Index).Tag) & ")", pathDBMySQL)
                        
                        log.AddItem "Borrando ctacte proveedores ... (ok!)"
                        Call BorrarBase("PCuentasCorrientes WHERE (NroInterno = " & Val(PbAcciones(Index).Tag) & ")", pathDBMySQL)
                        
                        log.AddItem "Borrando doc. compras ... (ok!)"
                        Call BorrarBase("PFactura WHERE (NroInterno = " & Val(PbAcciones(Index).Tag) & ")", pathDBMySQL)
                        
                        log.AddItem "Borrando movimiento de banco/caja ... (ok!)"
                        Call BorrarBase("BancosMovimientos WHERE (NroInterno = " & Val(PbAcciones(Index).Tag) & ")", pathDBMySQL)
                
                        MsgBox "Registro borrado correctamente !!!" + vcborrraasiento, vbInformation, "Mensaje ..."
                    
                        .Requery
                    
                        FormatoGrilla (0)
                    
                    End If
                End If
                

            End With
        
        Case 2
            Modificar
            
        Case 3
            VerDetalle
        
        Case 4
            Me.PbAcciones(Index).Enabled = False
            Buscar
            Me.PbAcciones(Index).Enabled = True
            
        Case 5
        
        frmTransaccionMantenimiento.vViene = Me.Name
        frmTransaccionMantenimiento.vnrointerno = vnrointerno
        frmTransaccionMantenimiento.Show
        
                    
'        Call BorrarBase("asientos  WHERE (numero = " & Str(vnumeroasiento) & ")", pathDBMySQL)
'        Call BorrarBase("asientosdetalle WHERE (numero = " & Str(vnumeroasiento) & ")", pathDBMySQL)
'        Call PbAcciones_Click(4)
'
'        MsgBox "El asiento fue borrado correctamente", vbInformation, "Asiento Nro:" + Str(vnumeroasiento)
        
        Case 6
        Unload Me
    End Select

If Err Then GrabarLog "PBAcciones_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub PBImpresion_Click(Index As Integer)
On Error Resume Next
    
    Select Case Index

        Case 0
            'Imprimir vSQL
            
        Case 1
            'Imprimir Shape (Asientos APPEND AsientosDetalle)
            
        Case 2
            'Imprimir Shape (Asientos APPEND AsientosDetalle) WHERE Asientos.Numero = n
    
    End Select

If Err Then GrabarLog "PBImpresion_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Buscar()
On Error Resume Next
    
    Dim vsql As String
    
    vsql = ""
    
    If Not Trim(dtpFecha(0).Text) = "" Then
        vsql = " AND ((A.Fecha >= '" & strfechaMySQL(dtpFecha(0).Value) & "') AND (A.Fecha <= '" & strfechaMySQL(dtpFecha(1).Value) & "'))"
    End If
    
    'Por Numero de Asiento
    If Not Val(txtBusqueda(0).Text) = 0 And Not Val(txtBusqueda(1).Text) = 0 Then
        vsql = vsql & " AND (A.Numero >= " & Val(txtBusqueda(0).Text) & " AND A.Numero <= " & Val(txtBusqueda(1).Text) & ")"
    End If
    
    'Por Numero Interno
    If Not Val(txtBusqueda(2).Text) = 0 And Not Val(txtBusqueda(3).Text) = 0 Then
        vsql = vsql & " AND (A.NroInterno >= " & Val(txtBusqueda(2).Text) & " AND A.NroInterno <= " & Val(txtBusqueda(3).Text) & ")"
    End If
    
    'Por Leyenda
    If Not Trim(txtBusqueda(4).Text) = "" Then
        vsql = vsql & " AND A.Leyenda LIKE '%" & Trim(txtBusqueda(4).Text) & "%'"
    End If
    
    CargarAsientos (vsql)
    
If Err Then GrabarLog "PBAcciones_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Modificar()
On Error Resume Next

    With frmAsientosAlta
        If Not rsAsientos.EOF = True And Not rsAsientos.BOF = True Then
        
        Else
        
        End If
    
    End With
    

If Err Then GrabarLog "Modificar", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub PushButton1_Click()
Dim vsql As String
Dim vcp1 As String


If Me.vopcli Then vcp1 = "CodigoCliente"
If Me.vopprovee Then vcp1 = "CodigoProveedor"


If Trim(vleyenda.Text) = "" Then Exit Sub

If Val(traerDatos2("select * from asientos from numero=" + Trim(Str(vnumeroasiento)), "nrobalance", pathDBMySQL)) = 0 Then
    vsql = "update asientos set " + vcp1 + " ='" + Trim(Me.vcp1.Text) + "' where numero=" + Trim(Str(vnumeroasiento)) '+ " and  nrobalance=" + Trim(Str(vnroBalance))
Else
    vsql = "update asientos set " + vcp1 + " ='" + Trim(Me.vcp1.Text) + "' where numero=" + Trim(Str(vnumeroasiento)) + " and  nrobalance=" + Trim(Str(vTxtnrobalance))
End If

EjecutarScript (vsql)
PbAcciones_Click (4)
vleyenda.Text = ""
'rsAsientosDetalle.Update

End Sub

Private Sub PusModificarAsiento_Click()
    cargarAsientoParaModificar
End Sub

Private Sub cargarAsientoParaModificar()
On Error Resume Next
        
        frmAsientosAlta.vModificando = True
        
        cargarAsientoParaModificarDetalle
        
        cargarAsientoParaModificarEncabezado
        
If Err Then GrabarLog "CargarDetalleAsiento", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cargarAsientoParaModificarEncabezado()
With frmAsientosAlta

    .txtNumero = vnumeroasiento
    .dtpFecha = Me.vfechaAsiento
    .txtLeyenda = Me.vleyenda
    .vcliprovee = vcp1
    .lblNroInterno = vnrointerno
    .vTxtnrobalance = vTxtnrobalance
    
    '----------------------------------------------------
    If Me.vopcli Then .vCodigoCliente = vcp1
    If Me.vopprovee Then .vCodigoProveedor = vcp1
    '----------------------------------------------------

End With
End Sub
Private Sub cargarGrillaFGAsientoDetalle()
    frmAsientosAlta.Show
    Set frmAsientosAlta.FGAsientoDetalle.DataSource = rsAsientosDetalle
    frmAsientosAlta.ConfigurarGrilla
   ' Rodri: Formatear la grilla Const las columnas y las dimensiones
    
End Sub

Private Sub cargarAsientoParaModificarDetalle()
Set rsAsientosDetalle = New ADODB.Recordset
    Dim sqlAsientosDetalle As String
    
    sqlAsientosDetalle = "SELECT idAsientosDetalle, Numero, AsientosDetalle.CodigoCuenta, Cuenta, Debe, Haber FROM AsientosDetalle INNER JOIN Cuentas ON AsientosDetalle.CodigoCuenta = Cuentas.CodigoCuenta WHERE (Numero = " & vnumeroasiento & ") ORDER BY idAsientosDetalle"

    With rsAsientosDetalle
        If .State = 1 Then .Close
        
        .CursorLocation = adUseClient
        
        Call .Open(sqlAsientosDetalle, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If .State = 1 Then
            cargarGrillaFGAsientoDetalle
        End If
    
    End With

End Sub
Private Sub PusModificarFecha_Click()
Dim vsql As String

If Trim(Me.vfechaAsiento.Text) = "" Or vidAsientos = 0 Then Exit Sub

    
    vsql = "update asientos set fecha='" + strfechaMySQL(vfechaAsiento) + "' where idAsientos=" + Trim(Str(vidAsientos))
   

EjecutarScript (vsql)
PbAcciones_Click (4)
vfechaAsiento.Text = ""
vidAsientos = 0

actualizarMoviRelacionados '

'rsAsientosDetalle.Update
End Sub
Private Sub actualizarMoviRelacionados()
On Error Resume Next

Dim vsql, vtraer, vlog As String

If vnrointerno = 0 Then Exit Sub

vlog = ""
' pcuentascorrientes
 
    vsql = "update pcuentascorrientes set fecha='" + strfechaMySQL(vfechaAsiento) + "' where nrointerno=" + Trim(Str(vnrointerno))
    vtraer = "select * from pcuentascorrientes where nrointerno =" + Trim(Str(vnrointerno))
    
    If Val(traerDatos2(vtraer, "nrointerno", pathDBMySQL)) > 0 Then
        vlog = vlog + "- CtaCte Proveedores "
        EjecutarScript (vsql)
    End If
        
' pfactura

    vsql = "update pfactura set fecha='" + strfechaMySQL(vfechaAsiento) + "' where nrointerno=" + Trim(Str(vnrointerno))
    vtraer = "select * from pfactura where nrointerno =" + Trim(Str(vnrointerno))
    
    If Val(traerDatos2(vtraer, "nrointerno", pathDBMySQL)) > 0 Then
        vlog = vlog + "- Factura Compra "
        EjecutarScript (vsql)
    End If
    
    
' bancosmovimientos
    vsql = "update bancosmovimientos set fecha='" + strfechaMySQL(vfechaAsiento) + "' where nrointerno=" + Trim(Str(vnrointerno))
    vtraer = "select * from bancosmovimientos where nrointerno =" + Trim(Str(vnrointerno))
    
    If Val(traerDatos2(vtraer, "nrointerno", pathDBMySQL)) > 0 Then
        vlog = vlog + "- Movimientos de Bancos "
        EjecutarScript (vsql)
    End If
 
 ' cuentascorrientes
 
 
    vsql = "update cuentascorrientes set fecha='" + strfechaMySQL(vfechaAsiento) + "' where nrointerno=" + Trim(Str(vnrointerno))
    vtraer = "select * from cuentascorrientes where nrointerno =" + Trim(Str(vnrointerno))
    
    If Val(traerDatos2(vtraer, "nrointerno", pathDBMySQL)) > 0 Then
        vlog = vlog + "- CtaCte Cliente "
        EjecutarScript (vsql)
    End If
        
' factura

    vsql = "update factura set fecha='" + strfechaMySQL(vfechaAsiento) + "' where nrointerno=" + Trim(Str(vnrointerno))
    vtraer = "select * from factura where nrointerno =" + Trim(Str(vnrointerno))
    
    If Val(traerDatos2(vtraer, "nrointerno", pathDBMySQL)) > 0 Then
        vlog = vlog + "- Factura Venta "
        EjecutarScript (vsql)
    End If
 
 ' cheques

    vsql = "update cheques set fecha='" + strfechaMySQL(vfechaAsiento) + "' where nrointerno=" + Trim(Str(vnrointerno))
    vtraer = "select * from cheques where nrointerno =" + Trim(Str(vnrointerno))
    
    If Val(traerDatos2(vtraer, "nrointerno", pathDBMySQL)) > 0 Then
        vlog = vlog + "- Cheques "
        EjecutarScript (vsql)
    End If


MsgBox "Se han actualizado las fechas de los siguientes módulos:" + Chr(13) + vlog
vnrointerno = 0

If Err Then Exit Sub
End Sub

Private Sub PusModificarLey_Click()
Dim vsql As String

If Trim(vleyenda.Text) = "" Then Exit Sub

If Val(traerDatos2("select * from asientos from numero=" + Trim(Str(vnumeroasiento)), "nrobalance", pathDBMySQL)) = 0 Then
    vsql = "update asientos set Leyenda='" + Trim(vleyenda.Text) + "' where numero=" + Trim(Str(vnumeroasiento)) '+ " and  nrobalance=" + Trim(Str(vnroBalance))
Else
    vsql = "update asientos set Leyenda='" + Trim(vleyenda.Text) + "' where numero=" + Trim(Str(vnumeroasiento)) + " and  nrobalance=" + Trim(Str(vTxtnrobalance))
End If

EjecutarScript (vsql)
PbAcciones_Click (4)
vleyenda.Text = ""

'rsAsientosDetalle.Update
End Sub

Private Sub txtBusqueda_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        If Index < 4 Then
            txtBusqueda(Index + 1).SetFocus
        Else
            Me.PbAcciones(4).SetFocus
        End If
    
    
    End If

If Err Then GrabarLog "txtBusqueda_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub vfechaAsiento_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    PusModificarFecha.SetFocus
End If
End Sub
