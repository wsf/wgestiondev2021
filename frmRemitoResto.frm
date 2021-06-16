VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "KlexGrid.ocx"
Object = "{AFD24A52-2823-4FBD-B75D-C282C11E1D98}#1.0#0"; "IFEpson.ocx"
Begin VB.Form frmRemitoResto 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Carga de Menu"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   180
   ClientWidth     =   18570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   18570
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.TabControl TabPago 
      Height          =   7695
      Left            =   3240
      TabIndex        =   135
      Top             =   600
      Visible         =   0   'False
      Width           =   6855
      _Version        =   851968
      _ExtentX        =   12091
      _ExtentY        =   13573
      _StockProps     =   68
      Appearance      =   10
      Color           =   4
      PaintManager.DisableLunaColors=   0   'False
      PaintManager.OneNoteColors=   -1  'True
      ItemCount       =   1
      Item(0).Caption =   "Forma de Pago"
      Item(0).ControlCount=   25
      Item(0).Control(0)=   "txtNroComprobante(2)"
      Item(0).Control(1)=   "txtNroComprobante(3)"
      Item(0).Control(2)=   "txtClientes(0)"
      Item(0).Control(3)=   "pbCarga(1)"
      Item(0).Control(4)=   "txtClientes(1)"
      Item(0).Control(5)=   "txtClientes(6)"
      Item(0).Control(6)=   "txtClientes(2)"
      Item(0).Control(7)=   "txtClientes(3)"
      Item(0).Control(8)=   "txtClientes(4)"
      Item(0).Control(9)=   "txtClientes(5)"
      Item(0).Control(10)=   "lblTotalTicket(1)"
      Item(0).Control(11)=   "lblPago(0)"
      Item(0).Control(12)=   "lblPago(1)"
      Item(0).Control(13)=   "lblTotalTicket(2)"
      Item(0).Control(14)=   "lblPago(2)"
      Item(0).Control(15)=   "txtEfectivo"
      Item(0).Control(16)=   "lblPago(3)"
      Item(0).Control(17)=   "lblPago(4)"
      Item(0).Control(18)=   "lblPago(5)"
      Item(0).Control(19)=   "cmdImprimir(1)"
      Item(0).Control(20)=   "cmdImprimir(0)"
      Item(0).Control(21)=   "lblPago(6)"
      Item(0).Control(22)=   "txtEmpleado(0)"
      Item(0).Control(23)=   "pbCarga(0)"
      Item(0).Control(24)=   "txtEmpleado(1)"
      Begin XtremeSuiteControls.PushButton cmdImprimir 
         Height          =   615
         Index           =   1
         Left            =   4920
         TabIndex        =   155
         Top             =   6840
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Imprimir Ticket (F5)"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Picture         =   "frmRemitoResto.frx":0000
      End
      Begin XtremeSuiteControls.FlatEdit txtClientes 
         Height          =   315
         Index           =   0
         Left            =   360
         TabIndex        =   138
         Top             =   5640
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   139
         Tag             =   "CodigoCliente"
         Top             =   5640
         Width           =   375
         _Version        =   851968
         _ExtentX        =   661
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtClientes 
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   140
         Top             =   5640
         Width           =   2415
         _Version        =   851968
         _ExtentX        =   4260
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtClientes 
         Height          =   285
         Index           =   6
         Left            =   4320
         TabIndex        =   141
         Top             =   5640
         Width           =   2295
         _Version        =   851968
         _ExtentX        =   4048
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtClientes 
         Height          =   285
         Index           =   2
         Left            =   1800
         TabIndex        =   142
         Top             =   6000
         Width           =   2445
         _Version        =   851968
         _ExtentX        =   4313
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtClientes 
         Height          =   285
         Index           =   3
         Left            =   4320
         TabIndex        =   143
         Top             =   6000
         Width           =   2295
         _Version        =   851968
         _ExtentX        =   4048
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtClientes 
         Height          =   285
         Index           =   4
         Left            =   4320
         TabIndex        =   144
         Top             =   6360
         Width           =   2295
         _Version        =   851968
         _ExtentX        =   4048
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtClientes 
         Height          =   285
         Index           =   5
         Left            =   1800
         TabIndex        =   145
         Top             =   6360
         Width           =   2445
         _Version        =   851968
         _ExtentX        =   4313
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtNroComprobante 
         Height          =   315
         Index           =   2
         Left            =   360
         TabIndex        =   136
         Top             =   4080
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtNroComprobante 
         Height          =   315
         Index           =   3
         Left            =   1800
         TabIndex        =   137
         Top             =   4080
         Width           =   4860
         _Version        =   851968
         _ExtentX        =   8572
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtEfectivo 
         Height          =   615
         Left            =   360
         TabIndex        =   147
         Top             =   1800
         Width           =   6255
         _Version        =   851968
         _ExtentX        =   11024
         _ExtentY        =   1085
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   21.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton cmdImprimir 
         Height          =   615
         Index           =   0
         Left            =   3120
         TabIndex        =   156
         Top             =   6840
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Cancelar (Esc)"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Picture         =   "frmRemitoResto.frx":059A
      End
      Begin XtremeSuiteControls.FlatEdit txtEmpleado 
         Height          =   315
         Index           =   0
         Left            =   360
         TabIndex        =   158
         Top             =   4800
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
         MaxLength       =   3
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   0
         Left            =   1320
         TabIndex        =   159
         Tag             =   "Mozo"
         Top             =   4800
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtEmpleado 
         Height          =   315
         Index           =   1
         Left            =   1800
         TabIndex        =   160
         Top             =   4800
         Width           =   4860
         _Version        =   851968
         _ExtentX        =   8572
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.Label lblPago 
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   157
         Top             =   4560
         Width           =   6500
         _Version        =   851968
         _ExtentX        =   11465
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Empleado :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblPago 
         Height          =   375
         Index           =   4
         Left            =   1320
         TabIndex        =   154
         Top             =   4005
         Width           =   375
         _Version        =   851968
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "-"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblPago 
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   153
         Top             =   5400
         Width           =   6500
         _Version        =   851968
         _ExtentX        =   11465
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Asignar Ticket a Cliente :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblPago 
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   152
         Top             =   3720
         Width           =   6500
         _Version        =   851968
         _ExtentX        =   11465
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Nro. Ticket-Factura :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblPago 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   151
         Top             =   2760
         Width           =   6500
         _Version        =   851968
         _ExtentX        =   11465
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cambio en Efectivo :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblTotalTicket 
         Height          =   615
         Index           =   2
         Left            =   360
         TabIndex        =   150
         Top             =   2880
         Width           =   6255
         _Version        =   851968
         _ExtentX        =   11033
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "0.00"
         ForeColor       =   192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   30
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
         AutoEllipsis    =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblPago 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   149
         Top             =   1560
         Width           =   6500
         _Version        =   851968
         _ExtentX        =   11465
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Total en Efectivo :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblPago 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   148
         Top             =   480
         Width           =   6500
         _Version        =   851968
         _ExtentX        =   11465
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Total Ticket:"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblTotalTicket 
         Height          =   615
         Index           =   1
         Left            =   240
         TabIndex        =   146
         Top             =   600
         Width           =   6255
         _Version        =   851968
         _ExtentX        =   11033
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "0.00"
         ForeColor       =   49152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   30
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
         AutoEllipsis    =   -1  'True
      End
   End
   Begin MSDataGridLib.DataGrid dgArticulosGrilla 
      Height          =   6615
      Left            =   9960
      TabIndex        =   130
      Top             =   1920
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   11668
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   0
      RowHeight       =   15
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
   Begin VB.PictureBox fraMesa 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   21
      Left            =   9720
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   132
      TabStop         =   0   'False
      Top             =   7560
      Visible         =   0   'False
      Width           =   1200
      Begin XtremeSuiteControls.Label lblNroMesa 
         Height          =   255
         Index           =   21
         Left            =   0
         TabIndex        =   134
         Top             =   120
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblTotalMesa 
         Height          =   135
         Index           =   21
         Left            =   0
         TabIndex        =   133
         Top             =   600
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   238
         _StockProps     =   79
         Alignment       =   2
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtArticulos 
      Height          =   495
      Left            =   9960
      TabIndex        =   131
      Top             =   1320
      Width           =   4575
      _Version        =   851968
      _ExtentX        =   8070
      _ExtentY        =   873
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   255
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GBTotal 
      Height          =   1215
      Left            =   9960
      TabIndex        =   128
      Top             =   0
      Width           =   4575
      _Version        =   851968
      _ExtentX        =   8070
      _ExtentY        =   2143
      _StockProps     =   79
      Caption         =   "Total Ticket"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.Label lblTotalTicket 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   129
         Top             =   360
         Width           =   4335
         _Version        =   851968
         _ExtentX        =   7646
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "0.00"
         ForeColor       =   49152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   30
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
      End
   End
   Begin XtremeSuiteControls.CheckBox chkLectorCodigoBarra 
      Height          =   255
      Left            =   8280
      TabIndex        =   127
      Top             =   120
      Width           =   1815
      _Version        =   851968
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Lector Cod. Barra (F6)"
      Appearance      =   2
   End
   Begin VB.Frame fraTotales 
      Caption         =   "Totales :"
      ForeColor       =   &H00808080&
      Height          =   2715
      Left            =   15600
      TabIndex        =   0
      Top             =   2400
      Width           =   3225
      Begin VB.TextBox txtIva 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   28
         Top             =   1415
         Width           =   1575
      End
      Begin VB.CheckBox chkTotalManual 
         Caption         =   "Total Manual"
         Height          =   195
         Left            =   150
         TabIndex        =   27
         Top             =   300
         Width           =   1275
      End
      Begin VB.CommandButton cmdActualizarTotal 
         Caption         =   "Actualizar Total"
         Height          =   315
         Left            =   1440
         TabIndex        =   24
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtImpuesto 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   2025
         Width           =   1575
      End
      Begin VB.TextBox txtIva 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   8
         Top             =   805
         Width           =   1575
      End
      Begin VB.TextBox txtPDescuento 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   1720
         Width           =   675
      End
      Begin VB.TextBox txtSubtotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   500
         Width           =   1575
      End
      Begin VB.TextBox txtIva 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         ForeColor       =   &H00800000&
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   9
         Top             =   1110
         Width           =   1575
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   2330
         Width           =   1575
      End
      Begin VB.TextBox txtDescuento 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2160
         TabIndex        =   12
         Top             =   1720
         Width           =   850
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "> I.V.A. 27 %:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   25
         TabIndex        =   29
         Top             =   1445
         Width           =   1400
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "> I.V.A. 10,5 %:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   25
         TabIndex        =   7
         Top             =   835
         Width           =   1400
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "> Total  :"
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
         Index           =   14
         Left            =   25
         TabIndex        =   6
         Top             =   2360
         Width           =   1400
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "> Impuesto  :"
         Height          =   195
         Index           =   13
         Left            =   25
         TabIndex        =   5
         Top             =   2055
         Width           =   1400
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "> Subtotal :"
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
         Index           =   0
         Left            =   25
         TabIndex        =   4
         Top             =   530
         Width           =   1400
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "> I.V.A. 21 %:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   25
         TabIndex        =   3
         Top             =   1140
         Width           =   1400
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "> Descuento %  :"
         Height          =   195
         Index           =   12
         Left            =   25
         TabIndex        =   2
         Top             =   1750
         Width           =   1400
      End
   End
   Begin XtremeSuiteControls.ComboBox cboCambioDeMesa 
      Height          =   315
      Left            =   4440
      TabIndex        =   96
      Top             =   120
      Width           =   645
      _Version        =   851968
      _ExtentX        =   1147
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
   End
   Begin VB.PictureBox fraMesa 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   20
      Left            =   8355
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   7525
      Width           =   1200
      Begin XtremeSuiteControls.Label lblTotalMesa 
         Height          =   135
         Index           =   20
         Left            =   0
         TabIndex        =   117
         Top             =   675
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   238
         _StockProps     =   79
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblNroMesa 
         Height          =   255
         Index           =   20
         Left            =   0
         TabIndex        =   93
         Top             =   120
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
   End
   Begin VB.PictureBox fraMesa 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   19
      Left            =   7140
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   7525
      Width           =   1200
      Begin XtremeSuiteControls.Label lblTotalMesa 
         Height          =   135
         Index           =   19
         Left            =   0
         TabIndex        =   116
         Top             =   675
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   238
         _StockProps     =   79
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblNroMesa 
         Height          =   255
         Index           =   19
         Left            =   0
         TabIndex        =   91
         Top             =   120
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
   End
   Begin VB.PictureBox fraMesa 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   18
      Left            =   5925
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   7525
      Width           =   1200
      Begin XtremeSuiteControls.Label lblTotalMesa 
         Height          =   135
         Index           =   18
         Left            =   0
         TabIndex        =   115
         Top             =   675
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   238
         _StockProps     =   79
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblNroMesa 
         Height          =   255
         Index           =   18
         Left            =   0
         TabIndex        =   89
         Top             =   120
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
   End
   Begin VB.PictureBox fraMesa 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   17
      Left            =   4710
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   7525
      Width           =   1200
      Begin XtremeSuiteControls.Label lblTotalMesa 
         Height          =   135
         Index           =   17
         Left            =   0
         TabIndex        =   114
         Top             =   675
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   238
         _StockProps     =   79
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblNroMesa 
         Height          =   255
         Index           =   17
         Left            =   0
         TabIndex        =   87
         Top             =   120
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
   End
   Begin VB.PictureBox fraMesa 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   16
      Left            =   3480
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   7525
      Width           =   1200
      Begin XtremeSuiteControls.Label lblTotalMesa 
         Height          =   135
         Index           =   16
         Left            =   0
         TabIndex        =   113
         Top             =   675
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   238
         _StockProps     =   79
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblNroMesa 
         Height          =   255
         Index           =   16
         Left            =   0
         TabIndex        =   85
         Top             =   120
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
   End
   Begin VB.PictureBox fraMesa 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   15
      Left            =   8355
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   6500
      Width           =   1200
      Begin XtremeSuiteControls.Label lblTotalMesa 
         Height          =   135
         Index           =   15
         Left            =   0
         TabIndex        =   112
         Top             =   675
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   238
         _StockProps     =   79
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblNroMesa 
         Height          =   255
         Index           =   15
         Left            =   0
         TabIndex        =   82
         Top             =   120
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
   End
   Begin VB.PictureBox fraMesa 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   14
      Left            =   7140
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   6500
      Width           =   1200
      Begin XtremeSuiteControls.Label lblTotalMesa 
         Height          =   135
         Index           =   14
         Left            =   0
         TabIndex        =   111
         Top             =   675
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   238
         _StockProps     =   79
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblNroMesa 
         Height          =   255
         Index           =   14
         Left            =   0
         TabIndex        =   80
         Top             =   120
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
   End
   Begin VB.PictureBox fraMesa 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   13
      Left            =   5925
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   6500
      Width           =   1200
      Begin XtremeSuiteControls.Label lblTotalMesa 
         Height          =   135
         Index           =   13
         Left            =   0
         TabIndex        =   110
         Top             =   675
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   238
         _StockProps     =   79
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblNroMesa 
         Height          =   255
         Index           =   13
         Left            =   0
         TabIndex        =   78
         Top             =   120
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
   End
   Begin VB.PictureBox fraMesa 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   12
      Left            =   4710
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   6500
      Width           =   1200
      Begin XtremeSuiteControls.Label lblTotalMesa 
         Height          =   135
         Index           =   12
         Left            =   0
         TabIndex        =   109
         Top             =   675
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   238
         _StockProps     =   79
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblNroMesa 
         Height          =   255
         Index           =   12
         Left            =   0
         TabIndex        =   76
         Top             =   120
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
   End
   Begin VB.PictureBox fraMesa 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   11
      Left            =   3495
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   6500
      Width           =   1200
      Begin XtremeSuiteControls.Label lblTotalMesa 
         Height          =   135
         Index           =   11
         Left            =   0
         TabIndex        =   108
         Top             =   675
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   238
         _StockProps     =   79
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblNroMesa 
         Height          =   255
         Index           =   11
         Left            =   0
         TabIndex        =   74
         Top             =   120
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
   End
   Begin VB.PictureBox fraMesa 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   10
      Left            =   8355
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   5475
      Width           =   1200
      Begin XtremeSuiteControls.Label lblTotalMesa 
         Height          =   135
         Index           =   10
         Left            =   0
         TabIndex        =   107
         Top             =   675
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   238
         _StockProps     =   79
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblNroMesa 
         Height          =   255
         Index           =   10
         Left            =   0
         TabIndex        =   72
         Top             =   120
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
   End
   Begin VB.PictureBox fraMesa 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   9
      Left            =   7140
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   5475
      Width           =   1200
      Begin XtremeSuiteControls.Label lblTotalMesa 
         Height          =   135
         Index           =   9
         Left            =   0
         TabIndex        =   106
         Top             =   675
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   238
         _StockProps     =   79
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblNroMesa 
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   70
         Top             =   120
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
   End
   Begin VB.PictureBox fraMesa 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   8
      Left            =   5925
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   5475
      Width           =   1200
      Begin XtremeSuiteControls.Label lblTotalMesa 
         Height          =   135
         Index           =   8
         Left            =   0
         TabIndex        =   105
         Top             =   675
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   238
         _StockProps     =   79
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblNroMesa 
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   68
         Top             =   120
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
   End
   Begin VB.PictureBox fraMesa 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   7
      Left            =   4710
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   5475
      Width           =   1200
      Begin XtremeSuiteControls.Label lblTotalMesa 
         Height          =   135
         Index           =   7
         Left            =   0
         TabIndex        =   104
         Top             =   675
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   238
         _StockProps     =   79
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblNroMesa 
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   66
         Top             =   120
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
   End
   Begin VB.PictureBox fraMesa 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   6
      Left            =   3495
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   5475
      Width           =   1200
      Begin XtremeSuiteControls.Label lblTotalMesa 
         Height          =   135
         Index           =   6
         Left            =   0
         TabIndex        =   103
         Top             =   675
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   238
         _StockProps     =   79
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblNroMesa 
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   64
         Top             =   120
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
   End
   Begin VB.PictureBox fraMesa 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   5
      Left            =   8355
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1200
      Begin XtremeSuiteControls.Label lblTotalMesa 
         Height          =   135
         Index           =   5
         Left            =   0
         TabIndex        =   102
         Top             =   600
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   238
         _StockProps     =   79
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblNroMesa 
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   62
         Top             =   225
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
   End
   Begin VB.PictureBox fraMesa 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   4
      Left            =   7140
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1200
      Begin XtremeSuiteControls.Label lblTotalMesa 
         Height          =   135
         Index           =   4
         Left            =   0
         TabIndex        =   101
         Top             =   600
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   238
         _StockProps     =   79
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblNroMesa 
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   61
         Top             =   225
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
   End
   Begin VB.PictureBox fraMesa 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   3
      Left            =   5925
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1200
      Begin XtremeSuiteControls.Label lblTotalMesa 
         Height          =   135
         Index           =   3
         Left            =   0
         TabIndex        =   100
         Top             =   600
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   238
         _StockProps     =   79
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblNroMesa 
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   60
         Top             =   225
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
   End
   Begin VB.PictureBox fraMesa 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   1
      Left            =   3495
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1200
      Begin XtremeSuiteControls.Label lblTotalMesa 
         Height          =   135
         Index           =   1
         Left            =   0
         TabIndex        =   98
         Top             =   600
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   238
         _StockProps     =   79
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblNroMesa 
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   58
         Top             =   225
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
   End
   Begin VB.PictureBox fraMesa 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   960
      Index           =   2
      Left            =   4710
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   4440
      Width           =   1200
      Begin XtremeSuiteControls.Label lblTotalMesa 
         Height          =   135
         Index           =   2
         Left            =   0
         TabIndex        =   99
         Top             =   600
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   238
         _StockProps     =   79
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblNroMesa 
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   59
         Top             =   225
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
      End
   End
   Begin Grid.KlexGrid KlexDetalle 
      Height          =   3255
      Left            =   120
      TabIndex        =   52
      Top             =   600
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   5741
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
      MouseIcon       =   "frmRemitoResto.frx":0B34
   End
   Begin EPSON_Impresora_Fiscal.PrinterFiscal FiscalEpson 
      Left            =   15840
      Top             =   8640
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Frame fraCargaDetalle 
      Height          =   555
      Left            =   120
      TabIndex        =   18
      Top             =   3840
      Width           =   9585
      Begin XtremeSuiteControls.FlatEdit txtDetalle 
         Height          =   315
         Index           =   1
         Left            =   795
         TabIndex        =   45
         Top             =   165
         Width           =   4455
         _Version        =   851968
         _ExtentX        =   7858
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtDetalle 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   44
         Top             =   160
         Width           =   650
         _Version        =   851968
         _ExtentX        =   1147
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtDetalle 
         Height          =   315
         Index           =   2
         Left            =   5280
         TabIndex        =   46
         Top             =   165
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1499
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   12648447
         BackColor       =   12648447
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtDetalle 
         Height          =   315
         Index           =   3
         Left            =   6150
         TabIndex        =   47
         Top             =   165
         Width           =   795
         _Version        =   851968
         _ExtentX        =   1411
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtDetalle 
         Height          =   315
         Index           =   4
         Left            =   6960
         TabIndex        =   48
         Top             =   165
         Width           =   795
         _Version        =   851968
         _ExtentX        =   1411
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtDetalle 
         Height          =   315
         Index           =   5
         Left            =   7765
         TabIndex        =   49
         Top             =   165
         Width           =   800
         _Version        =   851968
         _ExtentX        =   1411
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtDetalle 
         Height          =   315
         Index           =   6
         Left            =   8580
         TabIndex        =   50
         Top             =   165
         Width           =   875
         _Version        =   851968
         _ExtentX        =   1543
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   12632319
         BackColor       =   12632319
         Alignment       =   1
      End
   End
   Begin XtremeSuiteControls.PushButton cmdGrabarAsiento 
      Height          =   495
      Left            =   14760
      TabIndex        =   42
      Top             =   8520
      Width           =   1215
      _Version        =   851968
      _ExtentX        =   2143
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Grabar Asiento"
      UseVisualStyle  =   -1  'True
      Picture         =   "frmRemitoResto.frx":0B50
   End
   Begin XtremeSuiteControls.PushButton PBAcciones 
      Height          =   315
      Index           =   1
      Left            =   1560
      TabIndex        =   40
      ToolTipText     =   "Depura la Grilla de Detalles"
      Top             =   120
      Width           =   1395
      _Version        =   851968
      _ExtentX        =   2469
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Vaciar Detalle"
      Appearance      =   2
      Picture         =   "frmRemitoResto.frx":0F2B
      ImageAlignment  =   8
   End
   Begin XtremeSuiteControls.PushButton cmdVerComentario 
      Height          =   510
      Left            =   14760
      TabIndex        =   43
      Top             =   7920
      Width           =   1215
      _Version        =   851968
      _ExtentX        =   2143
      _ExtentY        =   900
      _StockProps     =   79
      Caption         =   "Ver comentario"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PBAcciones 
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   39
      ToolTipText     =   "Borra el Detalle Seleccionado de la Grilla"
      Top             =   120
      Width           =   1395
      _Version        =   851968
      _ExtentX        =   2469
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Borrar Detalle"
      Appearance      =   2
      Picture         =   "frmRemitoResto.frx":1344
      ImageAlignment  =   8
   End
   Begin VB.Frame fraTipoDocumento 
      ForeColor       =   &H00808080&
      Height          =   465
      Left            =   120
      TabIndex        =   31
      Top             =   8640
      Width           =   9435
      Begin VB.OptionButton opTipoDoc 
         Caption         =   "Nota de Débito"
         Height          =   225
         Index           =   5
         Left            =   7680
         MaskColor       =   &H8000000F&
         TabIndex        =   38
         Top             =   120
         Width           =   1545
      End
      Begin VB.OptionButton opTipoDoc 
         BackColor       =   &H80000000&
         Caption         =   "Remito"
         Height          =   315
         Index           =   4
         Left            =   1290
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   120
         Width           =   1275
      End
      Begin VB.OptionButton opTipoDoc 
         BackColor       =   &H80000000&
         Caption         =   "Factura"
         Height          =   315
         Index           =   0
         Left            =   30
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   120
         Width           =   1215
      End
      Begin VB.OptionButton opTipoDoc 
         BackColor       =   &H80000000&
         Caption         =   "Presupuesto"
         Height          =   315
         Index           =   1
         Left            =   2610
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   120
         Width           =   1395
      End
      Begin VB.OptionButton opTipoDoc 
         BackColor       =   &H80000000&
         Caption         =   "Nota de Crédito"
         Height          =   315
         Index           =   2
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   120
         Width           =   1515
      End
      Begin VB.OptionButton opTipoDoc 
         BackColor       =   &H80000000&
         Caption         =   "Documento"
         Height          =   315
         Index           =   3
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   120
         Value           =   -1  'True
         Width           =   1395
      End
      Begin VB.CommandButton cmdNotaCredito 
         Caption         =   "+"
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
         Left            =   7170
         TabIndex        =   32
         Top             =   150
         Width           =   435
      End
   End
   Begin VB.Frame fraDocNumero 
      Height          =   495
      Left            =   14760
      TabIndex        =   22
      Top             =   1800
      Width           =   3855
      Begin VB.Label lblTipoDocumento 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   41
         Top             =   120
         Width           =   2625
      End
      Begin VB.Label lblNroDocumento 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2640
         TabIndex        =   23
         Top             =   150
         Width           =   1095
      End
   End
   Begin VB.Frame fraPrecio 
      Height          =   585
      Left            =   8760
      TabIndex        =   21
      Top             =   8520
      Width           =   6615
      Begin VB.ComboBox cbolista 
         Height          =   315
         ItemData        =   "frmRemitoResto.frx":1762
         Left            =   1380
         List            =   "frmRemitoResto.frx":1764
         TabIndex        =   25
         Text            =   "1"
         Top             =   175
         Width           =   615
      End
      Begin VB.Label lbllista 
         AutoSize        =   -1  'True
         Caption         =   "> Lista de Precio:"
         Height          =   195
         Left            =   60
         TabIndex        =   26
         Top             =   220
         Width           =   1230
      End
   End
   Begin VB.Frame fraDocAbrir 
      Height          =   585
      Left            =   15720
      TabIndex        =   19
      Top             =   7680
      Width           =   2805
      Begin MSComctlLib.Toolbar BarraCliente 
         Height          =   330
         Left            =   0
         TabIndex        =   20
         Top             =   120
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   582
         ButtonWidth     =   1588
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         DisabledImageList=   "ImageList1"
         HotImageList    =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "Cliente"
               Object.ToolTipText     =   "Cliente"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Nuevo"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Buscar"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   8580
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483633
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   -2147483644
            UseMaskColor    =   0   'False
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   12
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRemitoResto.frx":1766
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRemitoResto.frx":1878
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRemitoResto.frx":198A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRemitoResto.frx":1A9C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRemitoResto.frx":1BAE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRemitoResto.frx":1CC0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRemitoResto.frx":1DD2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRemitoResto.frx":1EE4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRemitoResto.frx":1FF6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRemitoResto.frx":2108
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRemitoResto.frx":221A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRemitoResto.frx":232C
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSAdodcLib.Adodc bfactura 
      Height          =   330
      Left            =   15840
      Top             =   7320
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
      Left            =   15840
      Top             =   7320
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
   Begin MSAdodcLib.Adodc barticulo 
      Height          =   330
      Left            =   15960
      Top             =   7320
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
   Begin VB.Frame fraConfig 
      Caption         =   "Configuración :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1155
      Left            =   17280
      TabIndex        =   14
      Top             =   4560
      Width           =   675
      Begin VB.CheckBox bienes 
         Alignment       =   1  'Right Justify
         Caption         =   "Bienes de Capital :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   330
         TabIndex        =   16
         Top             =   990
         UseMaskColor    =   -1  'True
         Width           =   1635
      End
      Begin VB.ComboBox tprecio 
         Height          =   315
         ItemData        =   "frmRemitoResto.frx":243E
         Left            =   600
         List            =   "frmRemitoResto.frx":244B
         TabIndex        =   15
         Text            =   "Pesos ($)"
         Top             =   570
         Width           =   1245
      End
      Begin VB.Label Label2 
         Caption         =   "Tomar precio en:"
         Height          =   225
         Left            =   600
         TabIndex        =   17
         Top             =   330
         Width           =   1425
      End
   End
   Begin XtremeSuiteControls.DateTimePicker dtpFecha 
      Height          =   300
      Left            =   16680
      TabIndex        =   51
      Top             =   8520
      Width           =   1815
      _Version        =   851968
      _ExtentX        =   3201
      _ExtentY        =   529
      _StockProps     =   68
      Format          =   1
      CurrentDate     =   40312.4711689815
   End
   Begin XtremeSuiteControls.ListBox ListBoxMarkup 
      Height          =   4095
      Left            =   60
      TabIndex        =   83
      Top             =   4440
      Width           =   3375
      _Version        =   851968
      _ExtentX        =   5953
      _ExtentY        =   7223
      _StockProps     =   77
      BackColor       =   -2147483643
      UseVisualStyle  =   -1  'True
      SelectionBackColor=   15255731
      SelectionForeColor=   0
      EnableMarkup    =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PBAcciones 
      Height          =   315
      Index           =   3
      Left            =   5160
      TabIndex        =   94
      Top             =   120
      Width           =   1395
      _Version        =   851968
      _ExtentX        =   2469
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Guardar (F2)"
      Appearance      =   2
      Picture         =   "frmRemitoResto.frx":2472
      ImageAlignment  =   8
   End
   Begin XtremeSuiteControls.PushButton PBAcciones 
      Height          =   315
      Index           =   2
      Left            =   3000
      TabIndex        =   95
      Top             =   120
      Width           =   1395
      _Version        =   851968
      _ExtentX        =   2469
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Mover a ..."
      Appearance      =   2
      Picture         =   "frmRemitoResto.frx":2A0C
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton PBAcciones 
      Height          =   315
      Index           =   4
      Left            =   6600
      TabIndex        =   97
      Top             =   120
      Width           =   1635
      _Version        =   851968
      _ExtentX        =   2884
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Cerrar Mesa (F5)"
      Appearance      =   2
      Picture         =   "frmRemitoResto.frx":2E1D
      ImageAlignment  =   8
   End
   Begin MSDataGridLib.DataGrid dgArticulos 
      Height          =   1695
      Left            =   14760
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   2990
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   15
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
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1695
      Left            =   14760
      TabIndex        =   118
      Top             =   5400
      Width           =   2175
      _Version        =   851968
      _ExtentX        =   3836
      _ExtentY        =   2990
      _StockProps     =   79
      Caption         =   "GroupBox1"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton PBDocAFactura 
         Height          =   525
         Index           =   0
         Left            =   1200
         TabIndex        =   119
         Top             =   840
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3096
         _ExtentY        =   926
         _StockProps     =   79
         Caption         =   "Doc. A Factura"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
         Picture         =   "frmRemitoResto.frx":33B7
      End
      Begin XtremeSuiteControls.GroupBox GBDocAFactura 
         Height          =   1095
         Index           =   0
         Left            =   0
         TabIndex        =   120
         Top             =   240
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   1931
         _StockProps     =   79
         Caption         =   "IVA"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton RBIva 
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   121
            Top             =   600
            Width           =   795
            _Version        =   851968
            _ExtentX        =   1411
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Sumar"
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RBIva 
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   122
            Top             =   240
            Width           =   795
            _Version        =   851968
            _ExtentX        =   1411
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Restar"
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.PushButton PBDocAFactura 
         Height          =   525
         Index           =   1
         Left            =   3000
         TabIndex        =   123
         Top             =   840
         Width           =   1305
         _Version        =   851968
         _ExtentX        =   2302
         _ExtentY        =   926
         _StockProps     =   79
         Caption         =   "Vista Previa"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
         Picture         =   "frmRemitoResto.frx":39B9
      End
      Begin XtremeSuiteControls.GroupBox GBDocAFactura 
         Height          =   495
         Index           =   1
         Left            =   1200
         TabIndex        =   124
         Top             =   240
         Width           =   2535
         _Version        =   851968
         _ExtentX        =   4471
         _ExtentY        =   873
         _StockProps     =   79
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.CheckBox chkBorrarDocOriginal 
            Height          =   255
            Left            =   240
            TabIndex        =   125
            Top             =   180
            Width           =   2175
            _Version        =   851968
            _ExtentX        =   3836
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Borrar Documento Original"
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
         End
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pasar Documento a Factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         TabIndex        =   126
         Top             =   0
         Width           =   3450
      End
   End
End
Attribute VB_Name = "frmRemitoResto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCerrarMesa As Boolean
Dim vIdFactura As Long
Dim vLeyendaAsiento As String, vTotalAsiento As Double
Dim vtotal_real, vtotal_global  As Double
Dim vF5 As Integer
Public vvpdolar, vganancia As Double
Dim venvase As Boolean
Public vnrofactnc As String
Public vGrabaModo, vTipoDocumento, cargando As Integer, vnrocomprobante As Long
Dim ban As String
'----------------------------------------
Dim vRemitoControl As Long
Dim vCantidadControl As Integer
'----------------------------------------
Dim vOpenGrilla() As Boolean
Dim checksum() As Boolean
Dim rsArticulosGrilla As ADODB.Recordset
Dim vHabilitaDocAFactura As Boolean
Dim vClienteDefault As Boolean
Dim vIdTempMesa As Integer, vnroremito As Long, vNroRemitoTemp As Long
Dim rsArticulos As ADODB.Recordset
Private Enum MedioPago
    efectivoPesos = 1
    efectivoDolar = 2
    tarjeta = 3
    cheque = 4
    Deposito = 5
    NotaC = 8
    ContadoCredito = 11
    AjusteCredito = 12
End Enum
Private Sub GuardarMesa(vCerrar As Boolean, Optional vCuit As String)
On Error Resume Next

    Dim codClienteCobro As String
    
    If vIdTempMesa = 0 Then
        MsgBox "No Tiene Asignada una mesa!!!", vbExclamation, "Mensaje ..."
        Exit Sub
    End If
    
    Call GuardarDoc
    Call CargarFormatoMesas
    Call CargarMostrador
    
    If vCerrarMesa = True Then
        'Adrian (Manda para hacer el Pago)
        codClienteCobro = EsNulo(txtClientes(0).Text)
    
        With frmCobros
            .txtNroComprobante.Text = lblNroDocumento
            .txtTipoComp.Text = lblTipoDocumento.Caption
            .total = Val(txtTotal.Text)
            .pendiente = Val(txtTotal.Text)
            .remito = Val(vnroremito)
            .esComprobanteAutomatico = False
            .esFacturacion = True
            .codCliente = txtClientes(0).Text
    
            If opTipoDoc(2).Value Then
                'Antes de guardar tengo que pedir los datos de las facturas
                frmNroFactNC.Show
            Else
                'GuardarDoc
            End If
    
            Call .BuscarDatosOperacionesCliente(codClienteCobro, vnroremito)
            .txtImporteEfectivoPesos.SetFocus
    
            Load frmCobros
        
            .HabilitarControles (True)

        End With
    
    End If
    
If Err Then GrabarLog "GuardarMesa", Err.Number & " " & Err.Description, Me.Caption
End Sub
Public Function CargarDatosCliente(vValor As String, vPorDefecto As Boolean) As Boolean
    On Error Resume Next
    
    Dim rsCliente As New ADODB.Recordset, sqlCliente As String

    If vPorDefecto = True Then
        sqlCliente = "SELECT * FROM Clientes WHERE (PorDefecto = 'S')"
    Else
        sqlCliente = "SELECT * FROM Clientes WHERE ((Nombre = '" & Trim(vValor) & "') OR (Codigo = '" + Trim(vValor) + "'))"
    End If
    
    With rsCliente
        .CursorLocation = adUseClient
        Call .Open(sqlCliente, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not .EOF = True Then
            'Entro en el caso que no Sea El Cliente Por Defecto
            If Not vPorDefecto = True Then
                Select Case ValidarCliente(.Fields("Codigo").Value)
            
                    Case "CreditoMax"
                        Call Habilitar(Not True)
                        Exit Function
                
                    Case "Estado"
                        MsgBox "El Estado del Cliente No permite que pueda Facturarle!!", vbExclamation, "Mensaje ..."
                        Call Habilitar(Not True)
                        Exit Function
                
                    Case Else
                
                End Select
            End If
            
            Call Habilitar(True)
            txtClientes(0).Text = EsNulo(.Fields("Codigo").Value)
            txtClientes(1).Text = EsNulo(.Fields("Nombre").Value)
            txtClientes(2).Text = EsNulo(.Fields("Direccion").Value)
            txtClientes(3).Text = EsNulo(.Fields("Localidad").Value)
            txtClientes(4).Text = EsNulo(.Fields("Telefono").Value)
            txtClientes(5).Text = TraerDato("TipoIva", "idTipoIva =  '" & EsNulo(.Fields("idTipoIva").Value) & "'", "TipoIva")
            txtClientes(6).Text = Replace(EsNulo(.Fields("Cuit").Value), "-", "")

            If Not vPorDefecto = True Then
                If Not .Fields("u_pago").Value = "" Then
                    gupago = .Fields("u_pago").Value
                Else
                    gupago = "No encontrado"
                End If

                If Not EsNulo(.Fields("u_venta").Value) = "" Then
                    guventa = .Fields("u_venta").Value
                Else
                    guventa = "No encontrado"
                End If
    
                gsaldo = Format(.Fields("saldo").Value, "#######0.00")
                gcredito = Format(.Fields("CreditoMax").Value, "#######0.00")
            End If
            
            cbolista.Text = Val(.Fields("idListas").Value)
        
            'Ver info de cliente
            'If ConfigRemito(0) = True Then frmClienteInfo.foco
                       
            'Cargar numero de comprobante al inicio
            If ConfigRemito(1) = True Then NroComprobante
        
            CargarDatosCliente = True
        
            'FormatoGrillaDetalle (1)
        
        End If
        
    End With

    sqlCliente = ""
    
    If rsCliente.State = 1 Then
        rsCliente.Close
        Set rsCliente = Nothing
    End If
    
If Err Then GrabarLog "BuscarCliente", Err.Number & " " & Err.Description, Me.Name
End Function
Private Function ValidarCliente(vCodigoCliente As String) As String
On Error Resume Next

    If vCodigoCliente = "" Then
        ValidarCliente = Not True
        MsgBox "Debe ingresar un cliente !!!!", vbExclamation, "Mensaje ..."
        Exit Function
    End If

    Dim vSaldoCliente As Double, vCreditoMax As Double, i As Integer
    
    ValidarCliente = ""

    vSaldoCliente = Val(TraerDato("SaldoClientesSimple", "Codigo = '" & Trim(vCodigoCliente) & "'", "SaldoCliente"))
    vCreditoMax = Val(TraerDato("Clientes", "Codigo = '" & Trim(vCodigoCliente) & "'", "CreditoMax"))

    'Controlo que El Estado lo deje facturar
    If TraerDato("Estados", "idEstados = '" & TraerDato("Clientes", "Codigo = '" & Trim(vCodigoCliente) & "'", "idEstados") & "'", "SePuedeFacturar") = "N" Then
        
        For i = 0 To txtClientes.Count - 1
            txtClientes(i).Text = ""
            txtClientes(i).Tag = ""
        Next

        ValidarCliente = "Estado"
        
        Exit Function
    End If

    If (vSaldoCliente > vCreditoMax) And Not (vCreditoMax = 0) Then
        If Not MsgBox("El Saldo Actual del Cliente Supera el Limite de Crédito ¿ Permitir Movimiento de todas maneras ?", vbExclamation + vbYesNo, "Mensaje ...") = vbYes Then
                        
            ValidarCliente = "CreditoMax"
            
            For i = 0 To txtClientes.Count - 1
                txtClientes(i).Text = ""
                txtClientes(i).Tag = ""
            Next

            Exit Function
        End If
                    
    End If
        
If Err Then GrabarLog "ValidarCliente", Err.Number & " " & Err.Description, Me.Caption
End Function
Private Sub BuscaDoc()
    On Error Resume Next
    
    frmBuscarFactura.Show

    If Err Then GrabarLog "BuscaDoc", Err.Number & " " & Err.Description, Me.Name
End Sub
Function CalcularIva(vTipoIva As String, vvtotal As Double) As Double
    On Error Resume Next
    Dim ivatotal As Double
    

    Select Case Trim(txtClientes(4).Text)

        Case "Responsable Inscripto", "Resp.Inscripto"
            ivatotal = vvtotal * Val(vTipoIva) / 100

        Case "Responsable Monotributo"
            ivatotal = vvtotal * Val(vTipoIva) / 100

        Case "Consumidor Final"
            ivatotal = vvtotal * Val(vTipoIva) / 100
    End Select

    ivatotal = vvtotal * Val(vTipoIva) / 100
    CalcularIva = ivatotal

    If Err Then GrabarLog "caliva", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub CalcularTotales()
    On Error Resume Next
    
    Dim vTotalParcial, vTotal105, vTotal210, vTotal270, vdescuento, vImpuesto, vnegro, vIva As Double
    Dim i As Integer
    
    LimpiarTotales

    vTotalParcial = 0
    vTotal105 = 0
    vTotal210 = 0
    vTotal270 = 0
    vdescuento = 0
    vnegro = 0

    vLeyendaAsiento = ""
    
    With KlexDetalle
        
        If chkTotalManual.Value = 0 Then
            For i = 1 To .Rows - 1
                
                If Not Trim(.TextMatrix(i, 4)) = "" Then
                    
                    Call SeleccionarColor(.TextMatrix(i, 25), i)
                    
                    vtotal_global = 0
                    
                    vTotalParcial = vTotalParcial + Val(KlexDetalle.TextMatrix(i, 11))

                    If opTipoDoc(0).Value = True Then 'PANIC seleecionar Tipos de DOC CON IVA
                        
                        If txtClientes(4).Text = "Iva Responsable Inscripto" Or txtClientes(4).Text = "Iva Resp.Inscripto" Or txtClientes(4).Text = "Responsable Monotributo" Then
                            If .TextMatrix(i, 9) = "10.50" Then vTotal105 = vTotal105 + Val(.TextMatrix(i, 7) * 0.105)
                            If .TextMatrix(i, 9) = "21.00" Then vTotal210 = vTotal210 + Val(.TextMatrix(i, 7) * 0.21)
                            If .TextMatrix(i, 9) = "27.00" Then vTotal270 = vTotal270 + Val(.TextMatrix(i, 7) * 0.27)
                            
                            txtTotal.Text = Val(txtSubtotal.Text) + Val(txtIva(0).Text) + Val(txtIva(1).Text) + Val(txtIva(2).Text) - Val(txtDescuento.Text) + Val(txtImpuesto.Text)
                        Else
                            vTotal105 = 0
                            vTotal210 = 0
                            vTotal270 = 0
                        End If
                    Else
                        vTotal105 = 0
                        vTotal210 = 0
                        vTotal270 = 0
                    End If
                
                    'If Val(.Recordset("Tiva").Value) = 0 Then vnegro = vnegro + .Recordset("total").Value
    
                    vLeyendaAsiento = vLeyendaAsiento & Trim(.TextMatrix(i, 6)) & " - "
                    
                End If
            Next
        
        Else
            txtTotal.Text = Val(txtSubtotal.Text) + Val(txtIva(0).Text) + Val(txtIva(1).Text) + Val(txtIva(2).Text) - Val(txtDescuento.Text) + Val(txtImpuesto.Text)
            vTotalParcial = Val(txtTotal.Text)
        End If
        
    End With
    
    vLeyendaAsiento = Trim(vLeyendaAsiento)
    vLeyendaAsiento = Trim(Mid(vLeyendaAsiento, 1, Len(vLeyendaAsiento) - 1))
    
    vtotal_global = vTotalParcial

    If opTipoDoc(0).Value = True Then 'PANIC seleecionar Tipos de DOC CON IVA
        If txtClientes(4).Text = "Iva Responsable Inscripto" Then
            
            txtSubtotal.Text = vTotalParcial
    
            txtIva(0).Text = vTotal105
            txtIva(1).Text = vTotal210
            txtIva(2).Text = vTotal270
    
        Else
            txtSubtotal.Text = vTotalParcial + Val(Me.txtIva(0).Text) + Val(Me.txtIva(1).Text) + Val(Me.txtIva(2).Text) + vnegro
        End If
    Else
        txtSubtotal.Text = vTotalParcial + Val(Me.txtIva(0).Text) + Val(Me.txtIva(1).Text) + Val(Me.txtIva(2).Text) + vnegro
    End If
    
    vtotal_real = Val(txtSubtotal.Text + Val(Me.txtIva(0).Text) + Val(Me.txtIva(1).Text) + Val(Me.txtIva(2).Text) + vnegro)
    
    vdescuento = Str((vTotalParcial + vnegro) * Val(txtPDescuento.Text) / 100)
    txtDescuento.Text = vdescuento
    
    'vTotalParcial = vTotalParcial - vdescuento + txtImpuesto

    vImpuesto = Val(txtImpuesto.Text) * (vTotalParcial + vnegro) / 100
    
    If opTipoDoc(0).Value = True Then
        If txtClientes(4).Text = "Iva Responsable Inscripto" Then
            txtTotal.Text = vtotal_real + Val(vImpuesto) - Val(txtDescuento.Text)
        Else
            txtTotal.Text = Val(vTotalParcial) + Val(vnegro) + Val(vImpuesto) - Val(txtDescuento.Text)
        End If
    Else
        txtTotal.Text = Val(vTotalParcial) + Val(vnegro) + Val(vImpuesto) - Val(txtDescuento.Text)
    End If
    
    vTotalAsiento = Val(txtTotal.Text)

    lblTotalMesa(vIdTempMesa).Caption = "$ " & txtTotal.Text
    lblTotalTicket(0).Caption = "$ " & txtTotal.Text
    lblTotalTicket(1).Caption = "$ " & txtTotal.Text
    'DecorarTalles
    
    
    If Err Then GrabarLog "CalcularTotales", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub SeleccionarColor(vColor As String, vFila As Integer)
On Error Resume Next

    With KlexDetalle
        .Col = 25
        .Row = vFila
        
        Select Case Left(vColor, 1)
                    
            Case "N"
                .CellBackColor = vbRed
                
            Case "B"
                .CellBackColor = vbGreen
                
            Case ""
                .CellBackColor = vbWhite
                    
        End Select

    End With

If Err Then GrabarLog "SeleccionarColor", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub CargarComentario()
On Error Resume Next
    
    If Trim(txtClientes(0).Text) = "" Then
        MsgBox "Debe cargar un cliente previamente", vbExclamation, "Mensaje..."
        Exit Sub
    End If
    
    'With frmComentario
    '    .txtCliente.Text = Trim(txtClientes(0).Text)
    '    .txtCliente_Keypress 13
    '    .TabComentarios.Tab = 0
    'End With

If Err Then GrabarLog "CargarComentario", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub ElegirTipoPrecio()
    On Error Resume Next
        
    If Trim(tprecio.Text) = "Preguntar" Then
        If MsgBox("¿ Desea seleccionar el precio en Pesos ($) ?", vbYesNo) = vbYes Then
            txtDetalle(2).Text = Val(txtDetalle(2).Text)
            tprecio.Text = "Pesos ($)"
        Else
            txtDetalle(2).Text = inulo(vvpdolar) * gdolar
            tprecio.Text = "Dolar (u$s)"
        End If
            
    End If
        
    If Trim(tprecio.Text) = "Dolar (u$s)" Then txtDetalle(2) = vvpdolar
    If Trim(tprecio.Text) = "Pesos ($)" Then txtDetalle(2).Text = Val(txtDetalle(2).Text)
    
    If Err Then GrabarLog "ElegirTipoMoneda", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Function CargarFDetalle(vRemitoDetalle As Long) As Boolean
    On Error Resume Next

    Dim rsCargaFDetalle As New ADODB.Recordset, sqlCargaFDetalle As String
    Dim j As Integer
    
    sqlCargaFDetalle = "SELECT * FROM FDetalle WHERE (remito = " & vRemitoDetalle & ") ORDER BY idFDetalle ASC"
    
    With rsCargaFDetalle
        .CursorLocation = adUseClient
        
        Call .Open(sqlCargaFDetalle, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        FormatoGrillaDetalle (.RecordCount)
            
        If Not .EOF = True Then
            .MoveFirst
            
            j = 1
            
            Do Until .EOF = True
                
                KlexDetalle.TextMatrix(j, 1) = EsNulo(.Fields("idFDetalle").Value)
                KlexDetalle.TextMatrix(j, 2) = EsNulo(.Fields("Fecha").Value)
                KlexDetalle.TextMatrix(j, 3) = EsNulo(.Fields("Remito").Value)
                KlexDetalle.TextMatrix(j, 4) = "[" & EsNulo(.Fields("Codigo").Value) & "]"
                KlexDetalle.TextMatrix(j, 5) = EsNulo(.Fields("Cantidad").Value)
                KlexDetalle.TextMatrix(j, 6) = EsNulo(.Fields("Detalle").Value)
                
                
                If opTipoDoc(1).Value = True Or txtClientes(4).Text = "Consumidor Final" Or txtClientes(4).Text = "Responsable Monotributo" Then
                    
                    KlexDetalle.TextMatrix(j, 5) = Format(Val(.Fields("Precio").Value), "######0.00") '(precio)
                    KlexDetalle.TextMatrix(j, 11) = "0"
                
                Else
                    
                    If bienes.Value = 1 Then
                        KlexDetalle.TextMatrix(j, 5) = Format(Val(txtDetalle(2).Text) - (Val(txtDetalle(2).Text) * 9.5 / 100), "######0.00") ' (precio)
                        KlexDetalle.TextMatrix(j, 11) = "10.5"
                    Else
                        Dim vIdPorcentaje As String, vPorcentajeIva As String
                        
                        vIdPorcentaje = TraerDato("Articulos", "Codigo = '" & .Fields("Codigo").Value & "'", "idPorcentajeIva")
                        vPorcentajeIva = TraerDato("PorcentajeIva", "idPorcentajeIva = '" & vIdPorcentaje & "'", "Porcentaje")
                        
                        KlexDetalle.TextMatrix(j, 7) = Format(Val(.Fields("Precio").Value), "######0.00")
                        KlexDetalle.TextMatrix(j, 8) = Format(Val(.Fields("Descuento").Value), "######0.00")
                        KlexDetalle.TextMatrix(j, 9) = vPorcentajeIva
                        KlexDetalle.TextMatrix(j, 10) = Format(Val(.Fields("Impuesto").Value), "######0.00")
                        KlexDetalle.TextMatrix(j, 11) = Format(Val(.Fields("Total").Value), "######0.00")
                    
                    End If
                
                End If

                'KlexDetalle.TextMatrix(j, 6) = EsNulo(.Fields("Descuento").Value)
                'KlexDetalle.TextMatrix(j, 8) = EsNulo(.Fields("Total").Value)
    
                'If cboVenta = "Contado" Then
                '    KlexDetalle.TextMatrix(j, 9) = .Fields("total_cdo").Value
                'Else
                '    KlexDetalle.TextMatrix(j, 10) = .Fields("total_ctacte").Value
                'End If
    
                KlexDetalle.TextMatrix(j, 13) = EsNulo(.Fields("Envase").Value)
                KlexDetalle.TextMatrix(j, 15) = EsNulo(.Fields("Pago").Value)
                KlexDetalle.TextMatrix(j, 16) = EsNulo(.Fields("Resta").Value)
    
                'If cboVenta.Text = "Contado" Then
                '    KlexDetalle.TextMatrix(j, 14) = "SI"                     'Pagado
                '    KlexDetalle.TextMatrix(j, 15) = .Fields("Pago").Value    'Pago
                '    KlexDetalle.TextMatrix(j, 16) = 0                        'Resta
                'Else
                '    KlexDetalle.TextMatrix(j, 14) = "NO"                     'Pagado
                '    KlexDetalle.TextMatrix(j, 15) = 0                        'Pago
                '    KlexDetalle.TextMatrix(j, 16) = .Fields("Resta").Value   'Resta
                'End If
    
                KlexDetalle.TextMatrix(j, 17) = EsNulo(.Fields("TotalIva").Value)    'Totaliva
                KlexDetalle.TextMatrix(j, 18) = EsNulo(.Fields("ganancia").Value)    'Ganancia
                KlexDetalle.TextMatrix(j, 19) = EsNulo(.Fields("Sueldo").Value)      'Sueldo
                KlexDetalle.TextMatrix(j, 20) = EsNulo(.Fields("repartidor").Value)  'Repartidor
                KlexDetalle.TextMatrix(j, 21) = EsNulo(.Fields("Confirmado").Value)
                
                'Se tendria que borrar
                KlexDetalle.TextMatrix(j, 22) = EsNulo(.Fields("IdFDetalle").Value)
                
                '2010-07-23 Juan
                'KlexDetalle.Rows = KlexDetalle.Rows + 1
                'KlexDetalle.Row = KlexDetalle.Row + 1
        
                vRemitoControl = Val(.Fields("Remito").Value)
                vCantidadControl = vCantidadControl + 1
    
                .MoveNext
                j = j + 1
            Loop

            DecorarTalles
            
            CalcularTotales

            CargarFDetalle = True
        Else
        
            CargarFDetalle = Not True
        
        End If
            
    End With
        
    sqlCargaFDetalle = ""
    
    If rsCargaFDetalle.State = 1 Then
        rsCargaFDetalle.Close
        Set rsCargaFDetalle = Nothing
    End If
    
    If Err Then GrabarLog "CargarFDetalle (" & vRemitoDetalle & ")", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub CargarFactura(vRemitoDetalle As Long)
    On Error Resume Next

    With bfactura
        txtClientes(0).Text = .Recordset("codigo").Value
        txtClientes(1).Text = EsNulo(.Recordset("nombre").Value)
        txtClientes(2).Text = EsNulo(.Recordset("domicilio").Value)
        txtClientes(3).Text = EsNulo(.Recordset("localidad").Value)
        txtClientes(4).Text = EsNulo(.Recordset("Telefono").Value)
        txtClientes(5).Text = EsNulo(.Recordset("iva").Value)
        txtClientes(6).Text = EsNulo(.Recordset("cuit").Value)
        
'        txtNroInterno.Text = EsNulo(.Recordset("NroInterno").Value)
        
        vnroremito = .Recordset("remito").Value
        txtSubtotal.Text = .Recordset("subtotal").Value
        txtIva(1).Text = .Recordset("tiva").Value
        txtIva(0).Text = Format(.Recordset("tiva2").Value, "###########0.00")
        txtTotal.Text = .Recordset("total").Value
        txtDescuento.Text = .Recordset("descuento").Value
        txtImpuesto.Text = .Recordset("Impuesto").Value
        
        dtpFecha.Value = .Recordset("fecha").Value
    
        CargarTipoDocumento (.Recordset("tipo").Value)
    
    End With
    
    If Err Then GrabarLog "CargarFactura (" & vRemitoDetalle & ")", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub CargarRemito(vRemitoModif As Long)
    On Error Resume Next
    
    With bfactura
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM Factura WHERE (remito = " & vRemitoModif & ")"
        .Refresh

        If Not .Recordset.EOF Then
            CargarFactura (vRemitoModif)
            CargarFDetalle (vRemitoModif)
    
            vnrocomprobante = EsNulo(.Recordset("NComprobante").Value)
            lblNroDocumento.Caption = EsNulo(.Recordset("NComprobante").Value)

        End If
    
    End With
    
    If Err Then GrabarLog "CargarRemito (" & vRemitoModif & ")", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub HabilitarDocAFactura(vHabilita As Boolean)
On Error Resume Next

    vHabilitaDocAFactura = vHabilita
    
    PBDocAFactura(0).Enabled = vHabilita
    PBDocAFactura(1).Enabled = vHabilita
    
    RBIva(0).Enabled = vHabilita
    RBIva(1).Enabled = vHabilita
    
    chkBorrarDocOriginal.Enabled = vHabilita
    
    GBDocAFactura(0).Enabled = vHabilita
    GBDocAFactura(1).Enabled = vHabilita

If Err Then GrabarLog "HabilitarDocAFactura", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CargarTipoDocumento(vtipo)
    On Error Resume Next

    Select Case vtipo

        Case "Fact A"
            opTipoDoc(0).Value = True

        Case "Fact B"
            opTipoDoc(0).Value = True

        Case "Presupuesto"
            opTipoDoc(1).Value = True

        Case "Nota C"
            opTipoDoc(2).Value = True

        Case "Documento"
            opTipoDoc(3).Value = True

        Case "Remito"
            opTipoDoc(4).Value = True
    
        Case "Nota D"
            opTipoDoc(5).Value = True
            
    End Select

    If Err Then GrabarLog "cargartipodoc (" & vtipo & ")", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cboCambioDeMesa_GotFocus()
On Error Resume Next

    Call CargarComboNew("Mesas", "idMesas", cboCambioDeMesa, True)

If Err Then GrabarLog "cboCambioDeMesa_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cboLista_GotFocus()
    On Error Resume Next
    
    CargarCombo "Listas", "Lista", cbolista, False
    
    If Err Then GrabarLog "cboLista_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub chkTotalManual_Click()
On Error Resume Next


If Err Then GrabarLog "chkTotalManual_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Public Sub cmdActualizarTotal_Click()
    On Error Resume Next
    
    If chkTotalManual.Value = 0 Then CalcularTotales

    If Err Then GrabarLog "cmdActualizarTotal_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdImprimir_Click(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 0
           
            
        Case 1
            Call CerrarMesayPagar
    
    End Select

    
    TabPago.Visible = False
    Call Habilitar(True)


If Err Then GrabarLog "cmdImprimir_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cmdVerComentario_Click()
On Error Resume Next

    CargarComentario

If Err Then GrabarLog "cmdVerComentario_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub ConfirmarDetalle()
    On Error Resume Next
    
    Dim vdife1 As Double, i As Integer, vConceptoCaja
  
    Dim rsFDetalle As New ADODB.Recordset, sqlFDetalle As String
            
    sqlFDetalle = "SELECT * FROM FDetalle WHERE (Remito = " & Val(vnroremito) & ") ORDER BY idFDetalle ASC"
    
    With rsFDetalle
        Call .Open(sqlFDetalle, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not .EOF = True Then .MoveFirst
        
            For i = 1 To KlexDetalle.Rows - 1
                
                If Not Trim(KlexDetalle.TextMatrix(i, 1)) = "" Then
                    .Filter = "idFDetalle = " & Trim(KlexDetalle.TextMatrix(i, 1)) & ""
                    'Se Borro, Algo malo Paso
                    If .EOF = True Then .AddNew
                Else
                    .AddNew
                End If
            
                .Fields("Fecha").Value = strfechaMySQL(dtpFecha.Value) ' KlexDetalle.TextMatrix(i, 0)
                .Fields("Remito").Value = vnroremito
                .Fields("Cantidad").Value = Val(KlexDetalle.TextMatrix(i, 5))
                .Fields("Codigo").Value = Replace(Replace(KlexDetalle.TextMatrix(i, 4), "[", ""), "]", "")
                .Fields("Detalle").Value = EsNulo(KlexDetalle.TextMatrix(i, 6))
                .Fields("Precio").Value = Val(KlexDetalle.TextMatrix(i, 7))         '(precio)
                .Fields("TIva").Value = Val(KlexDetalle.TextMatrix(i, 9))           '(tiva)
                '.fields("devolucion").Value = KlexDetalle.TextMatrix(i, 7)         '(devolucion)
                .Fields("Total").Value = Val(KlexDetalle.TextMatrix(i, 11))         '(total)
   
                If TipoDocumento = "Fact A" And txtClientes(4).Text = "Iva Responsable Inscripto" Then
                    .Fields("TotalIva").Value = Val(KlexDetalle.TextMatrix(i, 17))  '(totaliva)
                Else
                    .Fields("TotalIva").Value = Val(KlexDetalle.TextMatrix(i, 11))
                End If
                 
                .Fields("confirmado").Value = "S"
                .Fields("Repartidor").Value = Trim(txtEmpleado(0).Text)

                ' bdetalle.Recordset("totaliva") = Talles.TextMatrix(Talles.Row, 8)
            
                If Not KlexDetalle.TextMatrix(i, 21) = "S" Then
                    .Update
                    KlexDetalle.TextMatrix(i, 1) = .Fields("idFDetalle").Value

                    If TipoDocumento = "Nota C" Then
                        Call GuardarEnStock("Remito-Nuevo", EsNulo(.Fields("Codigo").Value), strfechaMySQL(dtpFecha.Value), -Val(KlexDetalle.TextMatrix(i, 5)), "Devolucion de Mercaderia", KlexDetalle.TextMatrix(i, 1), 0)
                    Else
                        If Not TipoDocumento = "Presupuesto" Then
                            If Not KlexDetalle.TextMatrix(i, 24) = "S" Then
                                Call GuardarEnStock("Remito-Nuevo", EsNulo(.Fields("Codigo").Value), strfechaMySQL(dtpFecha.Value), Val(KlexDetalle.TextMatrix(i, 5)), "Salida de Mercaderia", KlexDetalle.TextMatrix(i, 1), 0)
                            Else
                                Call GuardarEnStock("Remito-Nuevo", EsNulo(KlexDetalle.TextMatrix(i, 4)), strfechaMySQL(dtpFecha.Value), Val(KlexDetalle.TextMatrix(i, 5)), "Actualizacion de Mercaderia", KlexDetalle.TextMatrix(i, 1), 0)
                            End If
                        End If
                    End If
            
                Else
                    Call GuardarEnStock("Remito-Modificar", EsNulo(.Fields("Codigo").Value), strfechaMySQL(dtpFecha.Value), Val(KlexDetalle.TextMatrix(i, 5)), "Salida de Mercaderia", KlexDetalle.TextMatrix(i, 1), 0)
                    .MoveNext
                End If

            Next
    
        
    End With

    sqlFDetalle = ""
    
    If rsFDetalle.State = 1 Then
        rsFDetalle.Close
        Set rsFDetalle = Nothing
    End If
    
If Err Then
    GrabarLog "ConfirmarDetalle", Left(Err.Number & " " & Err.Description, 99), Me.Name
    MsgBox "Revise si el documento fue guardado correctamente", vbCritical, "Cuidado"
Else
    checksum(2) = True
End If
End Sub
Private Sub CtaCte()
    On Error Resume Next
    
    With frmCtaCteC
        .Show
        '.txtCliente.Text = Trim(v(0).Text)
        '.txtCliente_Keypress (13)
    End With
    
    If Err Then GrabarLog "CtaCte", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdNotaCredito_Click()
On Error Resume Next

    If opTipoDoc(2).Value Then
        'Antes de guardar tengo que pedir los datos de las facturas
        frmNroFactNC.Show
    End If

If Err Then GrabarLog "cmdNotaCredito_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub dgArticulos_DblClick()
On Error Resume Next

    With rsArticulos
        If Not .EOF = True And Not .BOF = True Then
            txtDetalle(1).Text = .Fields("Codigo").Value
            Call txtDetalle_KeyPress(1, 13)
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

Private Sub dgArticulosGrilla_DblClick()
On Error Resume Next
        
    With rsArticulosGrilla
        If Not .EOF = True And Not .BOF = True Then
            GuardarRenglon ("Grilla")
        End If
    End With

If Err Then GrabarLog "dgArticulosGrilla_DblClick", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub lblTotalMesa_DblClick(Index As Integer)
On Error Resume Next

    Call CargarDetalleMesa("Mesas", lblNroMesa(Index).Tag)

If Err Then GrabarLog "lblNroMesa_DblClick", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub ListBoxMarkup_DblClick()
On Error Resume Next

    With ListBoxMarkup
        Call CargarDetalleMesa("Listado", .List(.ListIndex))
    End With

If Err Then GrabarLog "ListBoxMarkup_DblClick", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub ListBoxMarkup_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        With ListBoxMarkup
            Call CargarDetalleMesa("Listado", .List(.ListIndex))
        End With
    End If

If Err Then GrabarLog "ListBoxMarkup_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub CargarDetalleMesa(vVieneTipo As String, vValor As String)
On Error Resume Next

    Dim vInicial As Integer, vNroMesa As Integer
    
    Select Case vVieneTipo
    
        Case "Mesas"
            vNroMesa = Val(vValor)
        
        Case "Listado"
            'Analizo que mesa
            vInicial = Val(InStr(1, vValor, "Grid.Column='1'Text='", vbTextCompare))
            vNroMesa = Val(Mid(vValor, vInicial + 21, 1))
    
    End Select

    'Cargo los detalles en base a que mesa es
    vNroRemitoTemp = Val(TraerDato("TempMesas", "idMesas = " & vNroMesa & "", "Remito"))
    
    If vNroRemitoTemp = 0 Then
        vNroRemitoTemp = UltimoRemito("Factura")
        vnroremito = vNroRemitoTemp
    End If
    
    vIdTempMesa = vNroMesa
    
    Call SeleccionarMesa(Trim(vNroMesa))
    
If Err Then GrabarLog "CargarDetalleMesa", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub fraMesa_DblClick(Index As Integer)
On Error Resume Next
    
    Call CargarDetalleMesa("Mesas", lblNroMesa(Index).Tag)

If Err Then GrabarLog "picMesa_DblClick", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub lblNroMesa_DblClick(Index As Integer)
On Error Resume Next

    Call CargarDetalleMesa("Mesas", lblNroMesa(Index).Tag)

If Err Then GrabarLog "lblNroMesa_DblClick", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtArticulos_Change()
On Error Resume Next

    Call CargarArticulos(False, txtArticulos.Text)

If Err Then GrabarLog "txtArticulos_Change", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtArticulos_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        With rsArticulosGrilla
            If Not .EOF = True And Not .BOF = True Then
                GuardarRenglon ("Grilla")
            End If
        End With
    End If

If Err Then GrabarLog "txtArticulos_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtEfectivo_Change()
On Error Resume Next

    lblTotalTicket(2).Caption = Val(Format(lblTotalTicket(1).Caption, "####0.00")) - Val(Format(txtEfectivo.Text, "####0.00"))

If Err Then GrabarLog "txtEfectivo_Change", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtDetalle_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
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
    
        If KeyCode = 13 And Not Trim(txtDetalle(Index).Text) = "" Then
            dgArticulos_DblClick
        End If
    End If
    
If Err Then GrabarLog "f_KeyUp", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub KlexDetalle_DblClick()
On Error Resume Next

    With KlexDetalle
        .Col = 25
        Select Case .TextMatrix(.Row, 25)
        
            Case "Blanco"
                .CellBackColor = vbRed
                .TextMatrix(.Row, 25) = "Negro"

            Case "Negro"
                .TextMatrix(.Row, 25) = "Blanco"
                .CellBackColor = vbGreen
                

            Case ""
                .TextMatrix(.Row, 25) = "Negro"
                .CellBackColor = vbRed
        
        End Select
        
    End With

If Err Then GrabarLog "KlexDetalle_DblClick", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub KlexDetalle_LeaveCell()
On Error Resume Next

    With KlexDetalle
    
        Select Case .Col
        
            Case 0, 4
        
            Case 5 'Cantidad
                '.TextMatrix(.Row, 5) = .TextMatrix(.Row, 5) * .TextMatrix(.Row, 7)
                
        
            Case 7 'Precio
                '.TextMatrix(.Row, 11) = .TextMatrix(.Row, 5) * .TextMatrix(.Row, 7)
                
            
            Case 8
                '.TextMatrix(.Row, 11) = DescuentoImpuesto
             
            
            Case 9
            
            Case 10
            
            Case 11
            
            Case 25
            
        End Select


    End With

If Err Then GrabarLog "KlexDetalle_LeaveCell", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Function DescuentoImpuesto() As Double
On Error Resume Next

    Dim vdescuento As Double, vImpuesto As Double
    
    Dim vAuxiliar As Double
    
    With KlexDetalle
        vAuxiliar = .TextMatrix(.Row, 5) * .TextMatrix(.Row, 7)
        
        If Not .TextMatrix(.Row, 8) = "" Then
            vdescuento = vAuxiliar - (vAuxiliar * Val(.TextMatrix(.Row, 8)) / 100)
                    
                    
        Else
        
        End If
        

    End With
    
    DescuentoImpuesto = vdescuento + vImpuesto

    
If Err Then GrabarLog "DescuentoImpuesto", Err.Number & " " & Err.Description, Me.Caption
End Function
Private Sub PbAcciones_Click(Index As Integer)
On Error Resume Next
    
    Select Case Index
    
        Case 0
            BorrarDetalle
        
        Case 1
             VaciarDetalle
        
        'Mover a Otra Mesa
        Case 2
            MoverDetalle
            
        'Guardar Mesa
        Case 3
            If Not vIdTempMesa = 21 Then
                vF5 = 0
                vCerrarMesa = False
                Call GuardarMesa(vCerrarMesa)
                fraMesa_DblClick (21)
            Else
                'es mostrador no puede guardar temporalmente
            End If
            
        'Cerra Mesa
        Case 4
            Call Habilitar(False)
            
            With TabPago
               .Visible = True
                .Left = (Me.Width - .Width) / 2
                .Top = (Me.Height - .Height) / 2
                
                'Cargo el Numero de Ticket a Imprimir
                txtNroComprobante(2).Text = FiscalEpson.AnswerField_4
                txtNroComprobante(3).Text = FiscalEpson.AnswerField_3
                txtEfectivo.SetFocus
            End With
            
    End Select

If Err Then GrabarLog "PBBorrar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CerrarMesayPagar()
On Error Resume Next

    vF5 = 1
    vCerrarMesa = True
    
    Call GuardarMesa(vCerrarMesa, txtClientes(6).Text)
    vCerrarMesa = False
    fraMesa_DblClick (21)

If Err Then GrabarLog "CerrarMesayPagar", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub BorrarDetalle()
On Error Resume Next

    If MsgBox("Esta seguro que desea borrar este registro?", vbInformation + vbYesNo, "Mensaje ...") = vbYes Then
        Call BorrarBase("Stock WHERE (idFDetalle = " & KlexDetalle.TextMatrix(KlexDetalle.Row, 1) & ")", pathDBMySQL)
        Call BorrarBase("FDetalle WHERE (idFDetalle = " & KlexDetalle.TextMatrix(KlexDetalle.Row, 1) & ")", pathDBMySQL)
        
        KlexDetalle.RemoveItem KlexDetalle.RowSel
    
        If vCantidadControl >= 1 Then
            vCantidadControl = vCantidadControl - 1
        End If
    
        CalcularTotales
    End If
        
If Err Then GrabarLog "BorrarDetalle", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub VaciarDetalle()
On Error Resume Next

    Dim i As Integer
    
    If MsgBox("Esta seguro que desea Borrar Todos Los Registros ?", vbInformation + vbYesNo, "Mensaje ...") = vbYes Then
    
        For i = 1 To KlexDetalle.Rows - 1
            Call BorrarBase("Stock WHERE (idFDetalle = " & KlexDetalle.TextMatrix(KlexDetalle.Row, 1) & ")", pathDBMySQL)
            Call BorrarBase("FDetalle WHERE (idFDetalle = " & KlexDetalle.TextMatrix(KlexDetalle.Row, 1) & ")", pathDBMySQL)
            KlexDetalle.RemoveItem KlexDetalle.RowSel
        Next i
    
        KlexDetalle.Rows = 2
        FormatoGrillaDetalle (1)
        vCantidadControl = 0
        CalcularTotales

    End If
    
If Err Then GrabarLog "VaciarDetalle", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub MoverDetalle()
On Error Resume Next

    Dim vNroRemitoFDetalle As Long, vNroRemitoNuevo As Long, vIDFDetalle As Long
    
    With KlexDetalle
        '0 - Controlo que seleccione un Detalle
        '1 - Controlo que la mesa este o no abierta (Si no Genero Datos Nuevos
        '2 - Cambio el Nro de Remito del Detalle
        
        'Controlo que tenga el IdFDetalle
        If .TextMatrix(.Row, 1) = "" Then
            MsgBox "Para Mover el Detalle primero debera guardar la Mesa", vbInformation, "Mensaje ..."
            Exit Sub
        End If
            
        'Nro de Remito Actual
        vNroRemitoFDetalle = .TextMatrix(.Row, 3)
        
        'Nro de Remito Nuevo
        vNroRemitoNuevo = TraerDato("TempMesas", "idMesas = " & Val(cboCambioDeMesa.Text) & "", "Remito")
        
        'Id de Detalle Actual
        vIDFDetalle = .TextMatrix(.Row, 1)
        
        
        If vNroRemitoNuevo = vNroRemitoFDetalle Then
            MsgBox "Movimiento Incorrecto", vbExclamation, "Mensaje ..."
            Exit Sub
        Else
            'Entro aca solamente en el caso que la Mesa NO este Abierta
            If vNroRemitoNuevo = 0 Then
                
                'Traigo un Nro de Remito Nuevo
                vNroRemitoNuevo = UltimoRemito("Factura")
                NroComprobante
                Call GuardarFactura(Val(vNroRemitoNuevo))
                
                'Inserto en las Mesas Temporal para que me abra una mensa
                Call EjecutarScript("INSERT INTO TempMesas (idMesas, idMozos, Remito) VALUES (" & Val(cboCambioDeMesa.Text) & ", " & Trim(txtEmpleado(0).Text) & ", " & vNroRemitoNuevo & " )")
            End If
    
            'Cambio el Nro de Remito del FDetalle para relacionarlo con la mesa
            Call EjecutarScript("UPDATE FDetalle SET Remito=" & vNroRemitoNuevo & " WHERE idFDetalle = " & vIDFDetalle & " ")
            .RemoveItem .RowSel
            
            Call SeleccionarMesa(Val(cboCambioDeMesa.Text))
            
        End If
    
    End With
    
    'MsgBox "Movimiento realizado con exito", vbInformation, "Mensaje"
    'Actualizo Todo
    Call CargarFormatoMesas
            
        
If Err Then GrabarLog "MoverDetalle", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub pbCarga_Click(Index As Integer)
On Error Resume Next

    vVuelveBusqueda = Me.Name
    vVieneBusqueda = pbCarga(Index).Tag

    Select Case Index
        
        Case 0 To 10
            frmBusqueda.Show
            txtDetalle(0).SetFocus
            Me.ZOrder (1)
            
    End Select
            
If Err Then GrabarLog "pbCarga_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub PBDocAFactura_Click(Index As Integer)
On Error Resume Next

    Dim i As Integer, j As Integer
    
    '0- Validar
        'A- Controlar Todos los Detalles Con IVA (10,21,27)
        'B- Controlar Condicion de Iva del Cliente (Monotributo, Responsable Inscripto, Exento)
        'C- Otros (No Implementado)
    
    '1- Cambiar el Tipo de Documento
    '2- Cambiar los Detalles (Sumar el Iva o Restar el Iva)
    '3- ReCalcular Totales
    '4- Borrar el Documento Original (chkBorrarDocOriginal)
        
    '0
    If ValidarDocAFactura = False Then
        MsgBox "No estan bien Algunos Parametros del Documento o del Cliente", vbExclamation, "Mensaje ..."
        Exit Sub
    End If
    
    '1
    opTipoDoc(0).Value = True
    
    '2
    Call CambiarIvaEnDetalles
    
If Err Then GrabarLog "PBDocAFactura_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Function ValidarDocAFactura() As Boolean
On Error Resume Next

    Dim i As Integer, j As Integer
    '0
    
    'A
    With KlexDetalle
        For i = 1 To .Rows - 2
            
            If Val(.TextMatrix(i, 11)) = 10.5 Or Val(.TextMatrix(i, 11)) = 21 Or .TextMatrix(i, 11) = 27 Then
                ValidarDocAFactura = True
            Else
                ValidarDocAFactura = False
                Exit Function
            End If
        
        Next
    
    End With

    'B
    If txtClientes(4).Text = "Responsable Inscripto" Or txtClientes(4).Text = "Resp. Inscripto" Or txtClientes(4).Text = "Exento" Or txtClientes(4).Text = "Responsable Monotributo" Then
        ValidarDocAFactura = True
    End If
    
If Err Then GrabarLog "ValidarDocAFactura", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub CambiarIvaEnDetalles()
On Error Resume Next

    Dim i As Integer, j As Integer
    
    '2
    With KlexDetalle
        For i = 1 To .Rows - 2
            
            'Cantidad
            .TextMatrix(i, 2) = .TextMatrix(i, 2)
                 
            'Tiva
            .TextMatrix(i, 11) = .TextMatrix(i, 11)
            
            If RBIva(0).Value = True Then
                'Resta el Iva al Detalle
                
                'Precio
                .TextMatrix(i, 5) = .TextMatrix(i, 5) / Val(1 & "." & Val(Replace(.TextMatrix(i, 11), ".", "")))
                
                'Total
                .TextMatrix(i, 8) = Val(.TextMatrix(i, 2)) * Val(.TextMatrix(i, 5))
                
                'Total_CtaCte
                .TextMatrix(i, 10) = ""
                
                'Pago
                .TextMatrix(i, 15) = ""
                
                'Resta
                .TextMatrix(i, 16) = ""
                
                'TotalIva
                .TextMatrix(i, 17) = ""
            
            Else
            
                'Suma el Iva al Detalle
                .TextMatrix(i, 2) = ""      'Cantidad
                .TextMatrix(i, 5) = ""      'Precio
                .TextMatrix(i, 8) = ""      'Total
                .TextMatrix(i, 10) = ""     'Total_CtaCte
                .TextMatrix(i, 11) = ""     'T-IVA (tiva)
                .TextMatrix(i, 15) = ""     'Pago
                .TextMatrix(i, 16) = ""     'Resta
                .TextMatrix(i, 17) = ""     'TotalIva
            End If
            
        Next
    
    End With

If Err Then GrabarLog "CambiarIvaEnDetalles", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdGrabarAsiento_Click()
On Error Resume Next

    If vConfigGral.vIncluyeContabilidad = True Then
        With frmAsientosAlta
            .Show
            .ZOrder (0)
            .txtCuentaVieneDe.Text = Me.Caption
            .txtImporteVieneDe.Text = txtTotal.Text
        End With
    Else
        MsgBox "No Incluye el Modulo de Contabilidad...", vbInformation, "Mensaje ..."
    End If
    
If Err Then GrabarLog "cmdGrabarAsiento", Err.Number & " " & Err.Description, Me.Name
End Sub


Private Sub txtImpuesto_Change()
On Error Resume Next

    txtImpuesto.Text = Format(txtImpuesto.Text, "########0.00")

If Err Then GrabarLog "txtImpuesto_Change", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtPDescuento_Change()
    On Error Resume Next
    Dim vauxi As Double
    
    vauxi = Val(txtSubtotal) + Val(txtIva(0).Text) + Val(txtIva(1).Text) + Val(txtIva(2).Text)
    vauxi = (vauxi * Val(txtPDescuento.Text) / 100)
    txtDescuento.Text = vauxi

    If Err Then GrabarLog "txtPDescuento_Change", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtImpuesto_KeyPress(KeyAscii As Integer)
    On Error Resume Next

    If KeyAscii = 13 Then
        If (chkTotalManual = 0) Then
            'CalcularTotales
        Else
             txtTotal.SetFocus
        End If
    End If

    If Err Then GrabarLog "txtImpuesto_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub

Public Sub txtDetalle_Change(Index As Integer)
    On Error Resume Next
    Dim descuento, impuesto As Double

    If Index = 1 Then
        If Not chkLectorCodigoBarra.Value = xtpChecked Then
            Call MostrarCoincidencias("Articulos", txtDetalle(Index).Text)
            vOpenGrilla(0) = True
        End If
    Else
        
        If (ConfigRemito(5) = False) And (Val(cbolista.Text) = 0) Then
            MsgBox "Debe cargar un número de lista para poder facturar ", vbInformation, "Mensaje ..."
            cbolista.BackColor = vbRed
            cbolista.SetFocus
            Exit Sub
        End If
    
        cbolista.BackColor = vbWhite
    
        descuento = Val(txtDetalle(3).Text) * Val(txtDetalle(2).Text) / 100
        impuesto = Val(txtDetalle(5).Text) * Val(txtDetalle(2).Text) / 100

        If (Val(txtDetalle(0).Text) * Val(txtDetalle(2).Text)) > 0 Then
            txtDetalle(6).Text = Val(txtDetalle(0).Text) * Val(txtDetalle(2).Text) - descuento + impuesto
        End If
    End If
    
    If Err Then GrabarLog "txtDetalle_Change", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub MostrarCoincidencias(vTipoBusqueda As String, vBusqueda As String)
On Error Resume Next
    
    Select Case vTipoBusqueda
    
        Case "Articulos"
            Dim sqlArticulos As String, sqlTipoDetalle As String
    
            Set rsArticulos = New ADODB.Recordset
    
            If Trim(txtDetalle(1).Text) = "" Then
                sqlArticulos = "SELECT * FROM Articulos WHERE 1=2"
            Else
                sqlArticulos = "SELECT * FROM Articulos WHERE (Codigo LIKE '%" & Trim(vBusqueda) & "%') OR (Descrip LIKE '%" & Trim(vBusqueda) & "%')"
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

        Case "Clientes"
    
        Case "Empleados"

            
    End Select

If Err Then GrabarLog "MostrarCoincidencias", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub FormatoGrilla(vtipo As String)
On Error Resume Next
    
    Dim i As Integer
    
    Select Case vtipo
    
        Case "Articulos"
    
            With dgArticulos
                'Lo Paso al Frente
                .ZOrder (0)
        
                'Lo Ubico justo debajo de donde escribo
                .Top = fraCargaDetalle.Top + fraCargaDetalle.Height
        
                .Left = txtDetalle(1).Left
                .Width = txtDetalle(1).Width
        
                .HeadLines = 1.2
        
                For i = 0 To .Columns.Count - 1
        
                    Select Case i
            
                        Case 4
                            .Columns(i).Width = txtDetalle(1).Width - 750
                        Case Else
                            .Columns(i).Width = 0
                    End Select
                Next

            End With
    
        Case "ArticulosGrilla"
            
            With dgArticulosGrilla
                .HeadLines = 1.2
                .RowHeight = 350
                
                For i = 0 To .Columns.Count - 1
        
                    Select Case i
            
                        Case 4
                            .Columns(i).Width = txtDetalle(1).Width - 1500
                        
                        Case 22
                            .Columns(i).Width = 1000

                        Case Else
                            .Columns(i).Width = 0
                    End Select
                Next

            End With
            
        Case "Empleados"

    
    End Select
    
If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtDetalle_GotFocus(Index As Integer)
    On Error Resume Next
    
    With bdetalle
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM fdetalle WHERE (remito = " & Val(vnroremito) & ") ORDER BY idFDetalle ASC"
        .Refresh
        
        If Not .Recordset.EOF = True Then .Recordset.MoveLast
    End With
    
    If txtDetalle(4).Text = "" Then txtDetalle(4).Text = "21"

    If Err Then GrabarLog "f_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub txtDetalle_KeyPress(Index As Integer, _
                      KeyAscii As Integer)

    On Error Resume Next

    
    If KeyAscii = 13 Then
        Select Case Index

            Case 1
            
                If chkLectorCodigoBarra.Value = 0 Then
                    If Not vOpenGrilla(0) = True Then Pasar (Index)
                
                    With barticulo
                        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
                    
                        If vOpenGrilla(0) = True Then
                            With rsArticulos
                                If Not (.EOF = True) And Not (.BOF = True) Then
                                    barticulo.RecordSource = "SELECT Articulos.*, PorcentajeIva.Porcentaje FROM Articulos LEFT JOIN PorcentajeIva ON Articulos.idPorcentajeIva=PorcentajeIva.idPorcentajeIva WHERE (codigo =  '" & Trim(.Fields("Codigo").Value) & "')"
                                    '2010-07-24 Juan
                                    'barticulo.RecordSource = "SELECT * FROM Articulos "
                                    barticulo.Refresh
                                End If
                            End With
                        Else
                            .RecordSource = "SELECT Articulos.*, PorcentajeIva.Porcentaje FROM Articulos LEFT JOIN PorcentajeIva ON Articulos.idPorcentajeIva=PorcentajeIva.idPorcentajeIva WHERE (codigo =  '" & Trim(txtDetalle(1).Text) & "')"
                            .Refresh
                        End If
                    
                
                        If .Recordset.EOF = True Then
                            .RecordSource = "SELECT Articulos.*, PorcentajeIva.Porcentaje FROM Articulos LEFT JOIN PorcentajeIva ON Articulos.idPorcentajeIva=PorcentajeIva.idPorcentajeIva WHERE (Descrip LIKE '%" & Trim(txtDetalle(1).Text) & "%') OR (codigo LIKE '%" & Trim(txtDetalle(1).Text) & "%')"
                            .Refresh
                
                            If Not .Recordset.EOF Then
                                txtDetalle(1).Tag = .Recordset("codigo").Value
                                vganancia = TraerDato("Articulos_Ganancia", "(CodEmp = '" & Trim(txtEmpleado(0).Text) & "') AND (CodCli = '" & Trim(txtClientes(0).Text) & "') AND (CodRub = '" & barticulo.Recordset("rubro").Value & "')", "Porcentaje")
                                venvase = .Recordset("Envase").Value
                                If vganancia = 0 Then
                                    vganancia = Val(Format(.Recordset("Ganancia").Value, "#######0.00"))
                                End If
                                MostrarDetalle
                                ElegirTipoPrecio
                            Else
                                MsgBox "No Existe el articulo seleccionado", vbExclamation, "Mensaje ..."
                                txtDetalle(1).Text = ""
                                txtDetalle(1).Tag = ""
                            End If
                    
                        Else
                            MostrarDetalle
                            'ElegirTipoPrecio
                        End If
                
                    End With
                    dgArticulos.Visible = False
                    Pasar (Index)
                Else
                    With barticulo
                        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
                        .RecordSource = "SELECT Articulos.*, PorcentajeIva.Porcentaje FROM Articulos LEFT JOIN PorcentajeIva ON Articulos.idPorcentajeIva=PorcentajeIva.idPorcentajeIva WHERE (CodigoBarra = '" & Trim(txtDetalle(1).Text) & "')"
                        .Refresh
                        
                        If Not .Recordset.EOF = True Then
                                txtDetalle(1).Tag = .Recordset("codigo").Value
                                vganancia = TraerDato("Articulos_Ganancia", "(CodEmp = '" & Trim(txtEmpleado(0).Text) & "') AND (CodCli = '" & Trim(txtClientes(0).Text) & "') AND (CodRub = '" & barticulo.Recordset("rubro").Value & "')", "Porcentaje")
                                venvase = EsNulo(.Recordset("Envase").Value)
                                If vganancia = 0 Then
                                    vganancia = Val(Format(.Recordset("Ganancia").Value, "#######0.00"))
                                End If
                                MostrarDetalle
                            Pasar (5)
                        Else
                            MsgBox "No Existe el articulo seleccionado", vbExclamation, "Mensaje ..."
                            txtDetalle(1).Text = ""
                            txtDetalle(1).Tag = ""
                        End If
                    
                    End With
                
                End If
            Case Else
                Pasar (Index)
        End Select

    End If

    If Err Then GrabarLog "f_keypress (" & Index & "-" & KeyAscii & ")", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, _
                       Shift As Integer)
    On Error Resume Next


    If KeyCode = 27 And TabPago.Visible = True Then
        cmdImprimir_Click (0)
    End If
    
    If KeyCode = vbKeyF1 Then
        
    End If
    
    'Guardar
    If KeyCode = vbKeyF2 Then
        PbAcciones_Click (3)
    End If

    If KeyCode = vbKeyF3 Then
        BorrarDetalle
    End If
    
    If KeyCode = vbKeyF4 Then
        VaciarDetalle
    End If
    
    'Imprimir - Cerrar Mesa
    If KeyCode = vbKeyF5 Then
        If Not TabPago.Visible = True Then
            PbAcciones_Click (4)
        Else
            cmdImprimir_Click (1)
        End If
    
    End If
    
    'Cambiar Codigo Barra (Activo / No Activo)
    If KeyCode = vbKeyF6 Then
        chkLectorCodigoBarra.Value = Not CBool(chkLectorCodigoBarra.Value)
        If Trim(txtDetalle(0).Text) = "" Then
            txtDetalle(0).SetFocus
        Else
            txtDetalle(1).SetFocus
        End If
    
    End If
    
    If KeyCode = vbKeyF7 Then
        MoverDetalle
    End If
    
    If KeyCode = vbKeyF8 Then
        
    End If
    
    If KeyCode = vbKeyF9 Then
        txtArticulos.SetFocus
        txtArticulos.SelStart = 0
        txtArticulos.SelLength = Len(txtArticulos.Text)
    End If
    If KeyCode = vbKeyF10 Then
        
    End If
    If KeyCode = vbKeyF11 Then
        
    End If
    If KeyCode = vbKeyF12 Then
        
    End If
    
    'MsgBox KeyCode
    
If Err Then GrabarLog "Form_KeyUp", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub Form_Load()
    On Error Resume Next
    
    Dim i As Integer
    ReDim vOpenGrilla(0)

    With Me
        .Show
        .Width = 14745
        .Height = 9000
        .Top = 0
        .Left = 0
        .KeyPreview = True
    End With
    
    'Mozos=Usuarios
    Call LimpiarCampos
    Call CargarFormatoMesas
    Call CargarArticulos(True, "")
    Call CargarMostrador
    Call CargarMozoDeUsuario(vConfigGral.vUser)
    
    If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CargarMostrador()
On Error Resume Next

    Dim rsMostrador As New ADODB.Recordset, sqlMostrador As String

    sqlMostrador = "SELECT * FROM Mesas WHERE (idMozo = 0)"

    With rsMostrador
        If .State = 1 Then .Close
        
        Call .Open(sqlMostrador, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .State = 0 Then
            
            If Not .EOF = True Then
                vIdTempMesa = .Fields("idMesas").Value
                
                vnroremito = TraerDato("TempMesas", "idMesas = 21", "Remito")
                vNroRemitoTemp = vnroremito
                If vnroremito = 0 Then
                    vnroremito = UltimoRemito("Factura")
                    Call EjecutarScript("INSERT INTO TempMesas(idMesas, idMozos, Remito) VALUES (" & vIdTempMesa & "," & Val(0) & ", " & vNroRemitoTemp & ")")
                Else
                    'Por Ahora no hago nada, ya que se encuenta abierto el remito
                End If
                
                Call fraMesa_DblClick(vIdTempMesa)
            End If
        
        End If
    
    End With
    
    sqlMostrador = ""

    If rsMostrador.State = 1 Then
        rsMostrador.Close
        Set rsMostrador = Nothing
    End If
    
If Err Then GrabarLog "CargarMozoDeUsuario", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub CargarArticulos(vOpen As Boolean, varticulo As String)
On Error Resume Next

    If vOpen = True Then
    
        Set rsArticulosGrilla = New ADODB.Recordset
        
        Dim sqlArticulos As String

        sqlArticulos = "SELECT * FROM Articulos ORDER BY Descrip ASC"

        With rsArticulosGrilla
            If .State = 1 Then .Close
            .CursorLocation = adUseClient
            If .State = 1 Then .Close
        
            Call .Open(sqlArticulos, ConnDDBB, adOpenStatic, adLockReadOnly)
            
            Set dgArticulosGrilla.DataSource = rsArticulosGrilla
            
            FormatoGrilla ("ArticulosGrilla")
            
          End With
    Else
        With rsArticulosGrilla
            If .EOF = True Or .BOF = True Then .MoveFirst
            .Fields.Refresh
            .Find ("Descrip LIKE '%" & Trim(varticulo) & "%'")
    
        End With
    End If
    
  
    
    'sqlArticulos = ""

    'If rsArticulos.State = 1 Then
    '    rsArticulos.Close
    '    Set rsArticulos = Nothing
    'End If
    
If Err Then GrabarLog "CargarArticulos", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub CargarMozoDeUsuario(vUsuarioAMozo As String)
On Error Resume Next

    Dim rsMozos As New ADODB.Recordset, sqlMozos As String

    sqlMozos = "SELECT * FROM Mozos WHERE (Mozo = '" & vUsuarioAMozo & "')"

    With rsMozos
        If .State = 1 Then .Close
        
        Call .Open(sqlMozos, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .State = 0 Then
            If Not .EOF = True Then
                txtEmpleado(0).Text = EsNulo(.Fields("idMozos").Value)
                txtEmpleado(1).Text = EsNulo(.Fields("Mozo").Value)
            End If
        End If
    
    End With
    
    sqlMozos = ""

    If rsMozos.State = 1 Then
        rsMozos.Close
        Set rsMozos = Nothing
    End If
    
If Err Then GrabarLog "CargarMozoDeUsuario", Err.Number & " " & Err.Description, Me.Caption
End Sub
Public Sub Habilitar(vHabilita As Boolean)
On Error Resume Next
    
    Dim i As Integer
    
    'vOpenGrilla(0) = False
    With Me
        .fraTipoDocumento.Enabled = vHabilita
        For i = 0 To txtClientes.Count - 1
            txtClientes(i).Enabled = Not vHabilita
        Next
        
        '.FraAccionesDoc.Enabled = vHabilita
        .fraCargaDetalle.Enabled = vHabilita
        .fraConfig.Enabled = vHabilita
        .fraDocNumero.Enabled = vHabilita
        .fraPrecio.Enabled = vHabilita
        '.PBEnvaces.Enabled = vHabilita
        .fraTotales.Enabled = vHabilita
        .KlexDetalle.Enabled = vHabilita
        .PbAcciones(0).Enabled = vHabilita
        .PbAcciones(1).Enabled = vHabilita
        .PbAcciones(2).Enabled = vHabilita
        .PbAcciones(3).Enabled = vHabilita
        .PbAcciones(4).Enabled = vHabilita
        .cboCambioDeMesa.Enabled = vHabilita
        .dtpFecha.Enabled = vHabilita
        .cmdGrabarAsiento.Enabled = vHabilita
        .cmdVerComentario.Enabled = vHabilita
    End With

If Err Then GrabarLog "Habilitar", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub CargarFormatoMesas()
On Error Resume Next
    
    Dim rsMesas As New ADODB.Recordset, sqlMesas As String, vTotalMesa As Double, i As Integer

    'With XtremeSuiteControls.Icons
        '.LoadBitmap App.Path & "\Iconos\mesaclose.png", 10, xtpImageHot
        '.LoadBitmap App.Path & "\Iconos\Help.bmp", 12, xtpImageNormal
        '.LoadBitmap App.Path & "\Iconos\Love.png", 13, xtpImageNormal
    'End With
   
    sqlMesas = "SELECT M.idMesas as idMesa, idMozo, Mesa, Habilitada, Comentario, idTempMesas, TM.idMesas, TM.idMozos, Remito FROM Mesas M LEFT JOIN TempMesas TM ON M.idMesas=TM.idMesas ORDER BY M.idMesas;"
   
    With rsMesas
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        Call .Open(sqlMesas, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If .State = 1 Then If Not .EOF = True Then .MoveFirst
    
        ListBoxMarkup.Clear
        
        For i = 1 To Val(.RecordCount)
            vTotalMesa = 0
            
            'Controlo que la mesa no este abierta
            If Not EsNulo(.Fields("idTempMesas").Value) = "" Then
                fraMesa(i).Picture = LoadPicture(App.Path & "\Imagenes\MesaOpen.jpg")
                vTotalMesa = GenerarDato("SELECT Sum(Total) as Total FROM FDetalle WHERE Remito =  " & .Fields("Remito").Value & "", "Total")
                lblTotalMesa(i).Caption = "$ " & Val(Format(TraerDato("Factura", "remito = " & .Fields("Remito").Value & "", "Total"), "#####0.00"))
            Else
                fraMesa(i).Picture = LoadPicture(App.Path & "\Imagenes\MesaClose.jpg")
            End If
            
            fraMesa(i).Tag = .Fields("idMesa").Value
            lblNroMesa(i).Tag = .Fields("idMesa").Value
            
            lblTotalMesa(i).Caption = "$ " & Val(Format(vTotalMesa, "#####0.00"))
            
            lblNroMesa(i).Caption = Val(i)
            
            If Not .Fields("idTempMesas").Value = "" Then
                
                ListBoxMarkup.AddItem "<Border BorderThickness='2' BorderBrush='Red' Margin='0, 2, 0, 2' Padding='2'><StackPanel Orientation='Horizontal'>" & _
                    "<Grid><Grid.ColumnDefinitions><ColumnDefinition Width='Auto'/><ColumnDefinition Width='*'/></Grid.ColumnDefinitions>" & _
                    "<Grid.RowDefinitions><RowDefinition/><RowDefinition/></Grid.RowDefinitions>" & _
                    "<TextBlock TextAlignment='Left' FontWeight='Bold' Foreground='Navy'" & _
                    "Text='Nro Mesa:'/>" & _
                    "<TextBlock TextAlignment='Left' Grid.Row='1' FontWeight='Bold' Foreground='Navy'" & _
                    "Text='Importe Parcial:'/>" & _
                    "<TextBlock Margin='6, 0, 0, 0' Grid.Column='1'" & _
                    "Text='" & Val(i) & "'/>" & _
                    "<TextBlock Margin='6, 0, 0, 0'  Grid.Column='1' Grid.Row='1'" & _
                    "Text='$" & Format(vTotalMesa, "00.00") & "' />" & _
                    "</Grid></StackPanel></Border>"
                
            Else
            
                ListBoxMarkup.AddItem "<Border BorderThickness='2' BorderBrush='DodgerBlue' Margin='0, 2, 0, 2' Padding='2'><StackPanel Orientation='Horizontal'>" & _
                    "<Grid><Grid.ColumnDefinitions><ColumnDefinition Width='Auto'/><ColumnDefinition Width='*'/></Grid.ColumnDefinitions>" & _
                    "<Grid.RowDefinitions><RowDefinition/><RowDefinition/></Grid.RowDefinitions>" & _
                    "<TextBlock TextAlignment='Left' FontWeight='Bold' Foreground='Navy'" & _
                    "Text='Nro Mesa:'/>" & _
                    "<TextBlock TextAlignment='Left' Grid.Row='1' FontWeight='Bold' Foreground='Navy'" & _
                    "Text='Importe Parcial:'/>" & _
                    "<TextBlock Margin='6, 0, 0, 0' Grid.Column='1'" & _
                    "Text='" & Val(i) & "'/>" & _
                    "<TextBlock Margin='6, 0, 0, 0'  Grid.Column='1' Grid.Row='1'" & _
                    "Text='$" & "0.00' />" & _
                    "</Grid></StackPanel></Border>"
            
            End If

            .MoveNext
        Next

    End With


    With frmRemitoResto
    
        If Not Trim(fraMesa(1).Tag) = "" Then
        
            'vSinAsignar = False
        
            '.vMesa = picMesa(0).Caption
            '.vNroMesa = picMesa(0).Tag
        
        
            '.cboMozo.Tag = TraerDato("Mesas", "idMesas = " & .vNroMesa & "", "idMozo")
            '.cboMozo.Text = TraerDato("Mozos", "idMozos = " & .cboMozo.Tag & "", "Mozo")
        
            '.cboMesa.BackColor = lblTotal(Index).BackColor
            '.cboMozo.BackColor = lblTotal(Index).BackColor
            '.f(1).BackColor = lblTotal(Index).BackColor
        
            'If Not TraerDato("Reservas", "(idMesas = " & fraMesa(Index).Tag & ") AND (dia = '" & strfechaMySQL(Date) & "')", "idReservas") = "" Then
        
                '.lblReserva.Caption = " MESA RESERVADA "
                '.lblReserva.BackColor = lblTotal(Index).BackColor
        
            'Else
            
                '.lblReserva.Caption = ""
                '.lblReserva.BackStyle = vbTransparent
                '.lblReserva.Enabled = False
        
            'End If
        
            'SeleccionarMesa (vNroMesa)
        
        Else
            'MsgBox "La Mesa no se encuentra Habilitada!!!", vbInformation, "Mensaje ..."
            'vSinAsignar = True
            'Exit Sub
        End If
    
    End With
     
    sqlMesas = ""
    
    If rsMesas.State = 1 Then
        rsMesas.Close
        Set rsMesas = Nothing
    End If
        
If Err Then GrabarLog "CargarMesas", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub SeleccionarMesa(vNroMesa As Long)
On Error Resume Next

    Dim rsFactura As New ADODB.Recordset
    Dim sqlFactura As String, vRemitoMesa As Long
    
    vRemitoMesa = Val(TraerDato("TempMesas", "idMesas = " & vNroMesa & "", "Remito"))
    
    txtEmpleado(0).Text = TraerDato("Mesas", "idMesas = " & vNroMesa & "", "idMozo")
    txtEmpleado(1).Text = TraerDato("Mozos", "idMozos = " & txtEmpleado(0).Text & "", "Mozo")
    
    Me.Caption = "Documentos de Ventas:  [" & TraerDato("Mesas", "idMesas = " & vNroMesa & "", "Comentario") & "]"
    
    If Not vRemitoMesa = 0 Then
        sqlFactura = "SELECT * FROM Factura WHERE (Remito = " & vRemitoMesa & ")"
    
        With rsFactura
            If .State = 1 Then .Close
            Call .Open(sqlFactura, ConnDDBB, adOpenStatic, adLockReadOnly)
        
            MousePointer = vbHourglass
            vGrabaModo = 1
            CargarRemito (.Fields("remito").Value)
            Habilitar (True)
            
            MousePointer = vbDefault
            
            txtDetalle(0).SetFocus
            
        End With
    
    Else
        Limpiar
        LimpiarCampos
            
    End If
    
    sqlFactura = ""
    
    If rsFactura.State = 1 Then
        rsFactura.Close
        Set rsFactura = Nothing
    End If
    
If Err Then GrabarLog "SeleccionarMesa", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    'Unload frmClienteInfo
    'Unload frmCargarReparto
    vGrabaModo = 0
    Call BorrarBase("TempMesas WHERE (idMesas = 21)", pathDBMySQL)
    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

    If Err Then GrabarLog "Form_unload", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub GuardarCondicion() ' para el tema de factura
    On Error Resume Next

    Select Case vGrabaModo ' esta variable contiene 1 si se está modificando una factura
    
        Case 1
            'bfactura.Refresh
            'bfactura.Recordset.EditMode
            ' modifica iva venta
        
            ' ----- cristian
            ' Arollback(Index, 1) = "update"
            'bfactura.Refresh
            'bfactura.Recordset.MoveFirst
            With bfactura
                If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
                .RecordSource = "SELECT * FROM Factura WHERE (remito = " & Trim(vnroremito) & ")"
                .Refresh
            
                If .Recordset.EOF = True Then
                    MsgBox "Error al querer modificar la factura seleccionada", vbInformation
                    GrabarLog "GrabarCondicion", Err.Number & " " & Err.Description, Me.Name
                    Exit Sub
                End If
            
            End With
            '---------------------------------------
            ' bivaventa.Refresh
            ' bivaventa.Recordset.Find("clave = " + v(6))

            'If bivaventa.Recordset.EOF Then Exit Sub
            
            '----------------------------------------
            
            'grabaivaventa
        
        Case Else
    
            ' ----- cristian
            ' Arollback(Index, 1) = "delete"
        
            ' actualiza iva venta
            
            '-----------------01/08/2007
            'bivaventa.Refresh
            'bivaventa.Recordset.AddNew
            '-----------------
            'grabaivaventa
            With bfactura
                If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
                .RecordSource = "SELECT * FROM Factura WHERE 1=2"
                .Refresh
                .Recordset.AddNew
            End With
    End Select

    If Err Then GrabarLog "GuardarCondicion", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Function GuardarFactura(vNroRemitoFactura As Long) As Boolean
    On Error Resume Next

    With bfactura
        If vGrabaModo = 1 Then
            If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
            .RecordSource = "SELECT * FROM Factura WHERE (remito = " & vNroRemitoFactura & ")"
            .Refresh

            If Not .Recordset.EOF = True Then
                .Recordset.MoveFirst
            Else
                '2010-07-28 - Juan
                .Recordset.AddNew
                .Recordset("remito").Value = Val(vNroRemitoFactura)
            End If
        Else
            If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
            .RecordSource = "SELECT * FROM Factura"
            .Refresh
            
            If TraerDato("TempMesas", "idMesas = " & vIdTempMesa & "", "Remito") = 0 Then
                .Recordset.AddNew
            Else
                .Recordset.Filter = "remito = " & vNroRemitoFactura & ""
                If .Recordset.EOF = True Then
                    .Recordset.AddNew
                    .Recordset("remito").Value = Val(vNroRemitoFactura)
                End If
            End If
        End If
        
        .Recordset("Fecha").Value = dtpFecha.Value
        .Recordset("Hora").Value = Mid(Now(), 12, 8)
        .Recordset("Codigo").Value = EsNulo(Trim(txtClientes(0).Text))
        .Recordset("Nombre").Value = EsNulo(txtClientes(1).Text)
        .Recordset("Domicilio").Value = EsNulo(txtClientes(2).Text)
        .Recordset("Localidad").Value = EsNulo(txtClientes(3).Text)
        .Recordset("Telefono").Value = EsNulo(txtClientes(4).Text)
        .Recordset("Iva").Value = EsNulo(txtClientes(5).Text)
        .Recordset("Cuit").Value = EsNulo(txtClientes(6).Text)
                
        If Not vGrabaModo = 1 Then .Recordset("remito").Value = Val(vNroRemitoFactura)
        
        .Recordset("SubTotal").Value = Val(txtSubtotal.Text)
        .Recordset("tiva").Value = Val(txtIva(1).Text)
        .Recordset("tiva2").Value = Val(txtIva(0).Text)
        .Recordset("total").Value = Val(txtTotal.Text)
           
         .Recordset("descuento").Value = Val(txtDescuento.Text)
         .Recordset("Impuesto").Value = Val(txtImpuesto.Text)
         
         .Recordset("Ncomprobante").Value = vnrocomprobante
         .Recordset("Tipo").Value = TipoDocumento
         .Recordset("Cod_Repartidor").Value = EsNulo(txtEmpleado(0).Text)
         .Recordset("repartidor").Value = EsNulo(txtEmpleado(1).Text)
         .Recordset("Comentario").Value = Trim(lblTipoDocumento.Caption) & " " & Trim(lblNroDocumento.Caption)

         .Recordset.Update

        If TraerDato("TempMesas", "idMesas = " & vIdTempMesa & "", "Remito") = "" Then
            Call EjecutarScript("INSERT INTO TempMesas (idMesas, idMozos, Remito) VALUES (" & vIdTempMesa & ", " & Trim(txtEmpleado(0).Text) & ", " & vnroremito & " )")
        Else
            'Call EjecutarScript("UPDATE TempMesas SET idMesas = " & vIdTempMesa & ", idMozos = " & Trim(txtMozo(0).Text) & ", Remito = " & vNroRemito & " )")
        End If
        
        vTipoDocumento = TipoDocumento
        
        vIdFactura = Val(.Recordset("id").Value)
    End With
    
    If vCerrarMesa = True Then GuardarIva
    
    If Err Then
        MsgBox "La factura no fue guardada correctamente.", vbCritical, "Error..."
        GrabarLog "GuardarFactura", Err.Number & " " & Err.Description, Me.Name
        checksum(1) = False
    Else
        checksum(1) = True
        GuardarFactura = True
    End If

End Function
Private Sub UltimoNroInterno()
    On Error Resume Next
    
    Dim rsNroInterno As New ADODB.Recordset, sqlNroInterno As String
    
    sqlNroInterno = "SELECT MAX(NroInterno) as NroInterno FROM Factura"
    
    With rsNroInterno
        Call .Open(sqlNroInterno, ConnDDBB, adOpenStatic, adLockReadOnly)

        If Not .EOF = True Then
            'txtNroInterno.Text = Val(EsNulo(.Fields("NroInterno").Value)) + 1
        Else
            'txtNroInterno.Text = 1
        End If
        
        'Set .Recordset = Nothing
    End With
    
    sqlNroInterno = ""
    
    If rsNroInterno.State = 1 Then
        rsNroInterno.Close
        Set rsNroInterno = Nothing
    End If
    
    If Err Then GrabarLog "UltimoNroInterno", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Function Guardar() As Boolean
    On Error Resume Next
        
    If KlexDetalle.Rows = 2 And chkTotalManual.Value = 0 And KlexDetalle.TextMatrix(1, 4) = "" Then
        MsgBox "Debe cargar detalles para poder GRABAR el Documento", vbExclamation, "Mensaje ..."
        Guardar = False
        Exit Function
    End If
    
    'If Trim(txtClientes(0).Text) = "" Then
    '    MsgBox "Debe Ingresar un Cliente", vbExclamation, "Mensaje ..."
    '    Guardar = False
    '    Exit Function
    'End If
    
    If chkTotalManual.Value = 0 Then
        CalcularTotales ' calcular los totales nuevamente
    End If
    
    MsgBox "Detalle de Mesa Guardado", vbInformation, "Mensaje ..."
    
    'Controlo que venga un Nro de Remito
    GuardarFactura (vnroremito)
    
    'Guardo la operacion en la cta cte
    If vCerrarMesa = True Then WCtaCte (vnroremito)
    
    If vCerrarMesa = True Then Call BorrarBase("TempMesas WHERE Remito = " & vnroremito & "", pathDBMySQL)

    If chkTotalManual.Value = 0 Then ConfirmarDetalle
        
    If vCerrarMesa = True Then ImprimirTicket (vnroremito)
    
    If vCerrarMesa = True Then NuevoCliente
    
    cargando = 0
    
    If vCerrarMesa = True Then vCerrarMesa = False
    
    If Err Then
        MsgBox "La factura no fue cargada correctamente. Revisar las operaciones!", vbCritical, "Error..."
        GrabarLog "Guardar", Err.Number & " " & Err.Description, Me.Name
        vGrabaModo = 0
        checksum(3) = False
    Else
        Guardar = True
        checksum(3) = True
    End If

End Function
Private Sub GuardarCliente()
    On Error Resume Next
    
    Dim rsNuevoCliente As New ADODB.Recordset, sqlNuevoCliente As String, vCodigoNuevo As String
    
    sqlNuevoCliente = "SELECT * FROM Clientes WHERE 1=2"
    
    With rsNuevoCliente
        .CursorLocation = adUseClient
        Call .Open(sqlNuevoCliente, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        .AddNew

        vCodigoNuevo = Val(GenerarDato("SELECT MAX(Clientes.Codigo) AS UltimoCodigo FROM Clientes", "UltimoCodigo")) + 1
        
        vCodigoNuevo = String(4 - Len(vCodigoNuevo), "0") & vCodigoNuevo
         
        .Fields("codigo") = vCodigoNuevo
        .Fields("Nombre") = txtClientes(0).Text
        .Fields("Direccion") = txtClientes(1).Text
        .Fields("Localidad") = txtClientes(2).Text
        .Fields("Telefono") = txtClientes(3).Text
        .Fields("Iva") = txtClientes(4).Text
        .Fields("Cuit") = txtClientes(5).Text
        .Fields("Codigo_Num").Value = Val(.Fields("Codigo").Value)
        .Fields("Pasivo").Value = "NO"
        .Fields("Fecha_Alta").Value = Date
        .Fields("Comentario").Value = "Desde Remito"
        
        .Update
    
    End With
    
    MsgBox "Los datos del cliente fueron guardados"

    If Err Then GrabarLog "GuardarCliente", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub GuardarDoc()
    On Error Resume Next

    'If Trim(txtClientes(0).Text) = "" Then
    '    MsgBox "Tiene campos obligatorios vacios, complete la factura y vuelva a intentarlo", vbInformation, "Mensaje"
    '    Exit Sub
    'End If
    
    If Trim(txtEmpleado(0).Text) = "" Then
        MsgBox "Tiene campos obligatorios vacios, complete la factura y vuelva a intentarlo", vbInformation, "Mensaje"
        Exit Sub
    End If
    
    Dim vPeriodoFactura As String
    
    vPeriodoFactura = Year(dtpFecha.Value) & Mid(dtpFecha.Value, 4, 2)
    
    If vPeriodoFactura = TraerDato("IvaVentaCerrado", "Periodo = '" & vPeriodoFactura & "'", "Periodo") Then
        MsgBox "La Factura pertenece a un periodo de Iva Venta Ya Cerrado!!!", vbInformation, "Mensaje ..."
        Exit Sub
    End If
    
    With txtEmpleado(0)
        If ConfigRemito(4) = True And Trim(.Text) = "" Then
            .BackColor = vbRed
            .SetFocus
            Exit Sub
        End If
    End With

    ReDim checksum(4)
    
    txtEmpleado(0).BackColor = vbWhite
    
    MousePointer = vbHourglass
    
    checksum(0) = True
    
    If Guardar = True Then
        If vCerrarMesa = True Then
                'MousePointer = vbDefault

        Else
        
        End If
    End If
    
    If vCerrarMesa = True Then
        If vConfigGral.vIncluyeContabilidad = True Then
            With frmAsientosAlta
                .Show
                .txtCuentaVieneDe.Text = Me.Caption
                .txtCuentaVieneDe.Tag = txtClientes(0).Text
                .txtLeyenda.Text = vLeyendaAsiento
                .dtpFecha.Value = dtpFecha.Value
                .txtImporteVieneDe.Text = vTotalAsiento
        
                vTotalAsiento = 0
    
                .vVieneTabla = "Factura"
                .vVieneIdNombre = "id"
                .vVieneIdValor = vIdFactura
                
                ZOrder (1)
            End With
        End If
    End If
    
    If Err Then
        MsgBox "Error! Revisar las operaciones... : " & Trim(Err.Description), vbCritical
        GrabarLog "GuardarDoc", Err.Number & " " & Err.Description, Me.Name
    End If

    MousePointer = vbDefault

End Sub
Private Sub Imprimir()
    On Error Resume Next

    If Trim(txtClientes(0).Text) = "" Then
        MsgBox "Debe ingresar un cliente", vbExclamation, "Mensaje ..."
        Exit Sub
    End If
    
    ReDim checksum(4)
    
    Call GuardarMesa(True, txtDetalle(6).Text)
    
    Call ImprimirTicket(vnroremito)

If Err Then GrabarLog "Imprimir", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtIva_Change(Index As Integer)
    On Error Resume Next
    
    'txtIva(Index).Text = Format(txtIva(Index).Text, "#######0.00")

    If Err Then GrabarLog "txtIva_Change", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtIva_KeyPress(Index As Integer, _
                         KeyAscii As Integer)
    On Error Resume Next

    If KeyAscii = 13 Then
        If Index = 0 Then txtIva(1).SetFocus
        If Index = 1 Then txtIva(2).SetFocus
        If Index = 2 Then txtPDescuento.SetFocus
    End If

    If Err Then GrabarLog "txtIva_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtIva_LostFocus(Index As Integer)
    On Error Resume Next
    
    txtIva(Index).Text = Format(txtIva(Index).Text, "#####0.00")

    If Err Then GrabarLog "iva_LostFocus", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub LimpiarBase()
    On Error Resume Next
    
    KlexDetalle.Enabled = False

    LimpiarFDetalle
    
    KlexDetalle.Enabled = True

    If Err Then GrabarLog "limpiabase", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Limpiar()

    Dim i As Integer

    For i = 0 To txtDetalle.Count - 1
        txtDetalle(i).Text = ""
        txtDetalle(i).Tag = ""
    Next
    
    If fraCargaDetalle.Enabled = True Then
        txtDetalle(0).SetFocus
    Else
        txtClientes(0).SetFocus
    End If
    
End Sub
Public Sub LimpiarCampos()
    On Error Resume Next
    Dim i As Integer

    dtpFecha.Value = Date

    For i = 0 To txtClientes.Count - 1
        txtClientes(i).Text = ""
        txtClientes(i).Tag = ""
    Next

    vOpenGrilla(0) = False
    txtSubtotal.Text = ""
    txtIva(0).Text = ""
    txtIva(1).Text = ""
    txtIva(2).Text = ""
    txtTotal.Text = ""
    txtDescuento.Text = ""
    txtImpuesto.Text = ""
    lblTotalTicket(0).Caption = "0.00"
    lblTotalTicket(1).Caption = "0.00"
    txtEfectivo.Text = ""
    LimpiarFDetalle
        
    vCantidadControl = 0
    vRemitoControl = 0
    vGrabaModo = 0
    

    KlexDetalle.Rows = 2
    FormatoGrillaDetalle (1)
    
    lblNroDocumento.Caption = ""
    lblTipoDocumento.Caption = ""
    
    fraTipoDocumento.Enabled = True
    
    Call NroComprobante
    Call HabilitarDocAFactura(False)
    Call CargarDatosCliente("", True)
    
    With bdetalle
        If Not .ConnectionString = "" Then
            If Not .Recordset.EOF Then
                .Refresh
                .Recordset.MoveLast
            End If
        End If
    End With

    chkTotalManual.Value = 0
    vnrofactnc = 0
    
    
    If Err Then
        GrabarLog "LimpiarCampos", Err.Number & " " & Err.Description, Me.Name
    Else
        checksum(4) = True
    End If
End Sub
Private Sub LimpiarFDetalle()
On Error Resume Next
    
    Call BorrarBase("fdetalle WHERE confirmado = 'N'", pathDBMySQL)
    
If Err Then GrabarLog "LimpiarFDetalle", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub MostrarDetalle()
    On Error Resume Next
    
    'Cargo Codigo, Detalle, Precio, PCosto, TipoIVA
    With barticulo
        txtDetalle(1).Text = Trim(.Recordset("descrip").Value)
        txtDetalle(1).Tag = Trim(.Recordset("codigo").Value)
        
        Select Case Trim(txtClientes(4).Text)
            
            Case "Iva Responsable Inscripto"
                If opTipoDoc(0).Value = True Then
                    txtDetalle(2).Text = Val(.Recordset("Pventa" & Val(cbolista.Text)).Value)
                Else
                    'Es Responsable pero va otro documento
                    
                    
                    txtDetalle(2).Text = .Recordset("Pventa" & Val(cbolista.Text)) * Val("1." & Replace(.Recordset("Porcentaje").Value, ".", ""))
                End If
                        
            Case "Responsable Monotributo", "Consumidor Final", "Exento"
                txtDetalle(2).Text = Val(.Recordset("Pventa" & Val(cbolista.Text)).Value)  'Val(.Recordset("Pventa" & Val(cbolista.Text)).Value) * Val("1." & Replace(.Recordset("TipoIva").Value, ".", ""))
            
            Case Else
                txtDetalle(2).Text = Val(.Recordset("Pventa" & Val(cbolista.Text)).Value)
        End Select
        
        'If Not IsNull(.Recordset("TipoIva").Value) = True Then
        '    f(2).Text = .Recordset("Pventa" & Val(cboLista.Text)) * Val("1." & Replace(.Recordset("TipoIva").Value, ".", ""))
        '    vpespecial = False
        'Else
        '    f(2).Text = .Recordset("Pventa" & Val(cboLista.Text)).Value
        '    MsgBox "Cuidado .... el Articulo seleccionado no tiene un valor asignado de IVA", vbExclamation, "Mensaje ..."
        'End If

    
        txtDetalle(4).Text = TraerDato("PorcentajeIva", "idPorcentajeIva =  '" & .Recordset("idPorcentajeIva").Value & "'", "Porcentaje")
    End With
    
    ElegirTipoPrecio

    If Err Then GrabarLog "MostrarDetalle", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub NroComprobante()
    On Error Resume Next
    
    vnrocomprobante = Val(GenerarDato("SELECT MAX(NComprobante) as NComp FROM Factura WHERE Tipo = '" & TipoDocumento & "'", "NComp")) + 1
    lblNroDocumento.Caption = vnrocomprobante
    lblTipoDocumento.Caption = "Nro. " & TipoDocumento

    If Err Then GrabarLog "NroComprobante", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub NuevoCliente()
    On Error Resume Next
    
    RecargarForm
  
    If Err Then GrabarLog "NuevoCliente", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub NuevoDoc()
    On Error Resume Next
    
    MousePointer = vbHourglass
    
    ban = "1"
    Form_Load
    LimpiarCampos
    Limpiar
    txtClientes(0).SetFocus
       
    If Err Then
        MousePointer = vbDefault
        'MsgBox "Error! Revisar las operaciones", vbCritical, "Mensaje ..."
        If Err Then GrabarLog "NuevoDoc", Err.Number & " " & Err.Description, Me.Name
    
        Exit Sub
    End If

    MousePointer = vbDefault
End Sub
Private Sub PagarArticulo(ByRef rsFDetalle As Recordset, i As Integer)
On Error Resume Next
    
    Dim vrubro As String
    
    With barticulo
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM Articulos WHERE (codigo = '" & Trim(rsFDetalle.Fields("Codigo").Value) & "')"
        .Refresh
    
        If Not .Recordset.EOF = True Then
            If IsNull(.Recordset("idRubros").Value) = True Then
                vrubro = ""
            Else
                vrubro = .Recordset("idRubros").Value
            End If
        Else
            vrubro = ""
        End If
    End With
    
    Dim rsArticulosGanancia As New ADODB.Recordset, sqlArticulosGanancia As String
    
    sqlArticulosGanancia = "SELECT * FROM Articulos_Ganancia WHERE (CodEmp = '" & Trim(txtEmpleado(0).Text) & "') AND (CodCli = '" & Trim(txtClientes(0).Text) & "') AND (CodRub = '" & Trim(vrubro) & "')"
    
    ' ------- busco la ganancia que tiene el artículo ----------------------
    With rsArticulosGanancia
        Call .Open(sqlArticulosGanancia, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .EOF = True Then ' en el caso q no tengo asignado rubro, cliente, empleado porcentaje
            If Not barticulo.Recordset.EOF = True Then
                vganancia = Val(Format(barticulo.Recordset("ganancia_vendedor").Value, "#######0.00"))
            Else
                vganancia = 0
            End If
        Else ' en el caso que tenga asignado porcentaje
            vganancia = Val(Format(.Fields("porcentaje").Value, "#######0.00"))
        End If
    End With
    
    sqlArticulosGanancia = ""
    
    If rsArticulosGanancia.State = 1 Then
        rsArticulosGanancia.Close
        Set rsArticulosGanancia = Nothing
    End If
    
    '-----------------------------------------------------------------------
    
    With rsFDetalle
        .Fields("Ganancia").Value = vganancia
    
        'esta linea va dentro del if
        .Fields("Sueldo") = (vganancia * Val(KlexDetalle.TextMatrix(i, 5)) * Val(KlexDetalle.TextMatrix(i, 7))) / 100

        
        'Juan : 2010-07-19
        
        'If (Me.chkTotalContado.Value) Then
            
            .Fields("total_cdo").Value = Val(KlexDetalle.TextMatrix(i, 11))
            .Fields("Pagado").Value = "SI"
            .Fields("Pago").Value = Format(Val(KlexDetalle.TextMatrix(i, 5)) * Val(KlexDetalle.TextMatrix(i, 7)), "######0.00")
            .Fields("resta").Value = "0.00"
        'Else
         '   If Not vGrabaModo = 1 Then
                '.Fields("total_ctacte").Value = Val(KlexDetalle.TextMatrix(i, 11))
                '.Fields("Pagado").Value = "NO"
                '.Fields("resta").Value = Val(KlexDetalle.TextMatrix(i, 11))
         '   Else
                'SI MODIFICA
         '   End If
    
        'End If
    
    End With
    
If Err Then GrabarLog "PagarArticulo", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Pasar(Index As Integer)
    On Error Resume Next

    If Index >= 5 Then
        'If Val(f(6).Text) <= 0 Then
        '    MsgBox "La cantidad y el precio deben ser valores positivos !", vbCritical, "Error..."
        '   Exit Sub
        'End If
        
        With barticulo
            If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
            .RecordSource = "SELECT Articulos.*, PorcentajeIva.Porcentaje FROM Articulos LEFT JOIN PorcentajeIva ON Articulos.idPorcentajeIva=PorcentajeIva.idPorcentajeIva WHERE (codigo = '" & Trim(.Recordset.Fields("Codigo").Value) & "')"
            .Refresh
    
            If Not .Recordset.EOF = True Then
                If opTipoDoc(0).Value = True Then
                    If txtClientes(4).Text = "Iva Responsable Inscripto" Or txtClientes(4).Text = "Resp.Inscripto" Or txtClientes(4).Text = "Responsable Monotributo" Then
                        Select Case .Recordset.Fields("idPorcentajeIVA").Value
                               
                         Case 1
                            txtIva(0).Text = Val(txtIva(0).Text) + Val(txtDetalle(6).Text) * 0.105
                            
                         Case 2
                            txtIva(1).Text = Val(txtIva(1).Text) + Val(txtDetalle(6).Text) * 0.21
                    
                         Case 3
                            txtIva(2).Text = Val(txtIva(2).Text) + Val(txtDetalle(6).Text) * 0.27
                    
                        End Select
                
                        End If
                    End If
                End If
        
                fraTipoDocumento.Enabled = Not True
                GuardarRenglon ("")
           
         End With
    Else

        txtDetalle(Index + 1).SetFocus

    End If

    If Err Then GrabarLog "Pasar (" & Index & ")", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub RecargarForm()
    On Error Resume Next
    
    Me.Visible = Not True
    ban = "1"
    Form_Load
    vGrabaModo = 0
    Me.Visible = True
    txtClientes(0).SetFocus

    If Err Then GrabarLog "RecargarForm", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub GuardarRenglon(vViene As String)
    On Error Resume Next
    
    Dim j As Integer

    With KlexDetalle
        If .Rows <= 2 And .TextMatrix(.Rows - 1, 4) = "" Then
            FormatoGrillaDetalle (1)
        Else
            .Rows = .Rows + 1
           'FormatoGrillaDetalle (.Rows)
        End If
        
        j = .Rows - 1
        
        .TextMatrix(j, 1) = ""
        .TextMatrix(j, 2) = dtpFecha.Value
        .TextMatrix(j, 3) = Val(vnroremito)
        
        If vViene = "" Then
        
            'Si el valor tiene todos 0000 y no tiene  un string lo transforma en numero
            .TextMatrix(j, 4) = "[" & Trim(txtDetalle(1).Tag) & "]"
        
            .TextMatrix(j, 5) = Val(txtDetalle(0).Text)
            .TextMatrix(j, 6) = Trim(txtDetalle(1).Text)
            .TextMatrix(j, 7) = EsNulo(txtDetalle(2).Text)  'P. Venta
            .TextMatrix(j, 8) = EsNulo(txtDetalle(3).Text)  'Descuento
            .TextMatrix(j, 9) = EsNulo(txtDetalle(4).Text)  'Tipo Iva
            .TextMatrix(j, 10) = EsNulo(txtDetalle(5).Text)  'Impuesto
                
            If opTipoDoc(1).Value = True Or txtClientes(4).Text = "Consumidor Final" Or txtClientes(4).Text = "Responsable Monotributo" Then
                .TextMatrix(j, 5) = Format(Val(txtDetalle(2).Text), "######0.00") '(precio)
                .TextMatrix(j, 11) = "0"
            Else
                If bienes.Value = 1 Then
                    .TextMatrix(j, 5) = Format(Val(txtDetalle(2).Text) - (Val(txtDetalle(2).Text) * 9.5 / 100), "######0.00") ' (precio)
                    .TextMatrix(j, 11) = "10.5" '(tiva)
                Else
                    .TextMatrix(j, 11) = Val(txtDetalle(0).Text) * Val(txtDetalle(2).Text)
                    '.TextMatrix(j, 11) = TraerDato("Articulos", "Codigo = '" & TxtDetalle(1).Tag & "'", "idPorcentajeIva")
                    '.TextMatrix(j, 5) = Format(Val(f(2).Text), "######0.00") 'Format(Val(f(2).Text) - (Val(f(2).Text) * 17.3553 / 100), "######0.00") (precio)
                    '.TextMatrix(j, 11) = "21" '(tiva)
                End If
            End If
        
        Else
            .TextMatrix(j, 4) = "[" & Trim(rsArticulosGrilla.Fields("Codigo").Value) & "]"
        
            .TextMatrix(j, 5) = Val(1)
            .TextMatrix(j, 6) = Trim(rsArticulosGrilla.Fields("Descrip").Value)
            
            'Ver Lista de Precio
            .TextMatrix(j, 7) = EsNulo(rsArticulosGrilla.Fields("PVenta1").Value)
            .TextMatrix(j, 8) = EsNulo("")
            
            'Panic: Ver tipo Iva
            .TextMatrix(j, 9) = EsNulo("21")
            .TextMatrix(j, 10) = EsNulo("")
        
            'Panic: Ver Lista de Precio
            .TextMatrix(j, 11) = Val(1) * Val(rsArticulosGrilla.Fields("PVenta1").Value)
        End If
        
        Limpiar
        
        vRemitoControl = Val(vnroremito)
        vCantidadControl = vCantidadControl + 1

        CalcularTotales
    
    End With
    
    If Err Then GrabarLog "GuardarRenglon", Left(Err.Number & " " & Err.Description, 99), Me.Name
End Sub
Private Sub txtPDescuento_KeyPress(KeyAscii As Integer)
On Error Resume Next
    
    If KeyAscii = 13 Then
        txtSubtotal.SetFocus
        txtIva(0).SetFocus
        txtIva(1).SetFocus
        txtIva(2).SetFocus
        txtDescuento.SetFocus
        txtImpuesto.SetFocus
        
    End If

If Err Then GrabarLog "txtPDescuento_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtSubtotal_change()
    On Error Resume Next

    If opTipoDoc(0).Value = True Then
        txtTotal.Text = Val(txtSubtotal.Text) + Val(txtIva(0).Text) + Val(txtIva(1).Text) + Val(txtIva(2).Text) - Val(txtDescuento.Text) + Val(txtImpuesto.Text)
    Else
        txtTotal.Text = Val(txtSubtotal.Text)
    End If

    If Err Then GrabarLog "txtSubtotal_change", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtSubtotal_KeyPress(KeyAscii As Integer)
    On Error Resume Next

    If KeyAscii = 13 Then
        txtIva(0).SetFocus
        'chkTotalManual.Value = 1
    End If

    If Err Then GrabarLog "subtotal_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtDescuento_Change()
    On Error Resume Next
    
    txtDescuento.Text = Format(txtDescuento.Text, "#######0.00")

    If Err Then GrabarLog "tdescuento_Change", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub opTipoDoc_Click(Index As Integer)
    On Error Resume Next
    
    If Index = 0 Then KlexDetalle.BackColorFixed = &H8080FF
    If Index = 1 Then KlexDetalle.BackColorFixed = &HC0C000
    If Index = 2 Then KlexDetalle.BackColorFixed = &HC0C000
    If Index = 3 Then KlexDetalle.BackColorFixed = &HFF00&
    If Index = 4 Then KlexDetalle.BackColorFixed = &HFFFF&
    
    NroComprobante
    'LimpiarTotales
    CalcularTotales

    If Err Then GrabarLog "tipodoc_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub LimpiarTotales()
    On Error Resume Next

    If chkTotalManual.Value = 0 Then
        txtSubtotal.Text = ""
        
        txtPDescuento.Text = ""
        txtImpuesto.Text = ""
        txtDescuento.Text = ""
        txtTotal.Text = ""
        lblTotalTicket(0).Caption = "0.00"
        lblTotalTicket(1).Caption = "0.00"
    End If
    
    If Err Then GrabarLog "LimpiarTotales", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub BarraCliente_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next

    Select Case Button.Index

        Case 1
            GuardarCliente
            Call txtClientes_keypress(0, 13)

        Case 2
            NuevoCliente
            

        Case 3
            
            BuscaDoc
    
    End Select

    If Err Then GrabarLog "Toolbar1_ButtonClick (" & Button & ")", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtTotal_Change()
    On Error Resume Next
    
    txtTotal.Text = Format(txtTotal.Text, "#######0.00")

    If Err Then GrabarLog "txtTotal_Change", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtTotal_GotFocus()
    On Error Resume Next
    
    txtTotal.Text = Val(txtSubtotal.Text) + Val(txtIva(0).Text) + Val(txtIva(1).Text) + Val(txtIva(2).Text) - Val(txtDescuento.Text) + Val(txtImpuesto.Text)

    If Err Then GrabarLog "txtTotal_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub tprecio_Click()
    On Error Resume Next
    
    txtDetalle(0).SetFocus

    If Err Then GrabarLog "tprecio_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub UltimaVenta()
    On Error Resume Next

    If Not Val(txtTotal.Text) = "0" Then
        EjecutarScript ("UPDATE Clientes SET U_Venta = '" & strfechaMySQL(dtpFecha.Value) & "' WHERE Codigo = '" & txtClientes(0).Text & "'")
    Else
        Exit Sub
    End If
    
    If Err Then GrabarLog "UltimaVenta", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub txtClientes_keypress(Index As Integer, _
                      KeyAscii As Integer)
    On Error Resume Next

    If KeyAscii = 13 Then
    
        If Index = 7 Then
            txtDetalle(0).SetFocus
            NroComprobante
            Exit Sub
        End If
       
        If Index = 0 Then
            If CargarDatosCliente(txtClientes(0).Text, False) = True Then
                vnroremito = UltimoRemito("Factura")
                If Not vGrabaModo = 1 Then NroComprobante

                If txtEmpleado(0).Text = "" Then
                    pbCarga(0).SetFocus
                Else
                    txtDetalle(0).SetFocus
                End If
            End If
        Else
            If Index >= 5 Then
                txtDetalle(0).SetFocus
            Else
                If Index > 6 Then Index = 0
                txtClientes(Index + 1).SetFocus
            End If
        End If
    
    End If
    
    If Err Then GrabarLog "v_keypress (" & Index & "-" & KeyAscii & ")", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub lblNroDocumento_Click()
    On Error Resume Next
    
    lblNroDocumento.Caption = Trim(Str(InputBox("Ingresar nro. de Factura: ", "Nro. Factura...")))

    If Not Val(lblNroDocumento.Caption) = 0 Then vnrocomprobante = Val(lblNroDocumento.Caption)

    If Err Then GrabarLog "vnro_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Function TipoDocumento() As String
    On Error Resume Next
    Dim i As Integer

    For i = 0 To 5

        If opTipoDoc(i).Value = True Then

            Select Case i

                Case 0

                    If txtClientes(4).Text = "Iva Responsable Inscripto" Then
                        TipoDocumento = "Fact A"
                        Exit For
                    Else
                        'Este es el original - LO ATO CON ALAMBRE!!!!
                        TipoDocumento = "Fact B"
                        Exit For
                    End If

                Case 1
                    TipoDocumento = "Presupuesto"
                    Exit For

                Case 2
                    TipoDocumento = "Nota C"
                    Exit For

                Case 3
                    TipoDocumento = "Documento"
                    Exit For

                Case 4
                    TipoDocumento = "Remito"
                    Exit For
                
                Case 5
                    TipoDocumento = "Nota D"
                    Exit For
            
            End Select

        End If

    Next

    If Err Then GrabarLog "TipoDocumento", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub ControlRemito()
    On Error Resume Next
    
    Dim rsControl As New ADODB.Recordset, sqlControl As String
    
    sqlControl = "SELECT * FROM Factura INNER JOIN FDetalle ON Factura.remito = FDetalle.remito WHERE (Factura.remito = " & vRemitoControl & ")"
    
    With rsControl
        Call .Open(sqlControl, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If .RecordCount <> vCantidadControl Then
    
            'MsgBox "Error cuando se graba el documento de " & v(0).Text & ""
            
        Else
        
            'Todo bien
        
        End If
    
    End With
    
    sqlControl = ""
    
    rsControl.Close
    Set rsControl = Nothing

    If Err Then GrabarLog "ControlRemito", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub GuardarIva()
    On Error Resume Next
    
    Dim vTipoDocumentoIva As String
    
    vTipoDocumentoIva = TipoDocumento
    
    If vTipoDocumentoIva = "Fact A" Or vTipoDocumento = "Fact B" Or vTipoDocumentoIva = "Nota C" Then
        
        If vGrabaModo = 0 Then
            Call EjecutarScript("INSERT INTO IvaFacturaVenta (remito, Iva105, Iva210, Iva270) VALUES (" & vnroremito & ", " & Val(txtIva(0).Text) & ", " & Val(txtIva(1).Text) & ", " & Val(txtIva(2).Text) & ");")
        Else
            'Panic 'Controlar si es Doc A Fact
            'Call EjecutarScript("INSERT INTO IvaFacturaVenta (remito, Iva105, Iva210, Iva270) VALUES (" & vRemito & ", " & Val(txtIva(0).Text) & ", " & Val(txtIva(1).Text) & ", " & Val(txtIva(2).Text) & ");")
            Call EjecutarScript("UPDATE IvaFacturaVenta SET Iva105 = " & Val(txtIva(0).Text) & ", Iva210 = " & Val(txtIva(1).Text) & ", Iva270 = " & Val(txtIva(2).Text) & " WHERE (remito = " & vnroremito & ")")
        End If
    
    End If
    
    If Err Then GrabarLog "GuardarIva", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub DecorarTalles()
Dim i, j, k As Integer

On Error Resume Next

    With KlexDetalle
    
        k = 0
        For i = 0 To .Rows - 1

            '2010-07-23 Juan
            '.TextMatrix(.Row, 8) = Format(.TextMatrix(.Row, 8), "#####0.00")
            '.TextMatrix(.Row, 5) = Format(.TextMatrix(.Row, 5), "#####0.00")

            .Row = i
    
            For j = 1 To 11
                .Col = j
                If k = 1 Then
                    .CellBackColor = &HFFC0C0
                End If
            Next
            If k = 0 Then
                k = 1
            Else
                k = 0
            End If
        Next
    
    End With

If Err Then GrabarLog "DecorarTalles", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub WCaja(vConceptoCaja, vFechaCaja As Date, vImporteCaja As Double, vCodCliente As String)
    On Error Resume Next
    
    Dim rsCaja As New ADODB.Recordset
    Dim sqlCaja As String
    
    With rsCaja
        'If Not grabamodo = 1 Then
        
            sqlCaja = "SELECT * FROM caja"
            .Open sqlCaja, ConnDDBB, adOpenDynamic, adLockPessimistic
        
            .AddNew
            .Fields("remito").Value = vnroremito
        
        'Else
            
        '    sqlCaja = "SELECT * FROM caja WHERE (remito = " & Trim(vremito) & ")"
        '    .Open sqlCaja, ConnDDBB, adOpenDynamic, adLockPessimistic
        '
        '    If .EOF = True Then
        '        .AddNew
        '        .Fields("Remito").Value = vremito
        '    End If
            
        'End If
        
        .Fields("fecha").Value = strfechaMySQL(vFechaCaja)
        .Fields("Importe").Value = Val(vImporteCaja)
        
        .Fields("CodigoCliente").Value = vCodCliente
        
        .Fields("Usuario").Value = vConfigGral.vUser
        .Fields("CodigoConcepto").Value = vConceptoCaja
        .Fields("comentario").Value = ""
            
        .Fields("NroCheque") = Null
        .Fields("FechaDeposito") = Null
        .Fields("FechaConfeccion") = Null
        .Fields("idCajas") = Null
        
        .Update
    
    End With
    
    sqlCaja = ""
    
    rsCaja.Close
    Set rsCaja = Nothing
   
If Err Then GrabarLog "WCAJA", Left(Err.Number & " " & Err.Description, 99), Me.Name
End Sub
Public Sub WCtaCte(vnroremito As Long)
    On Error Resume Next
    
    Dim rsCtaCteC As New ADODB.Recordset, sqlCtaCteC As String
    
    sqlCtaCteC = "SELECT * FROM CuentasCorrientes WHERE (remito = " & vnroremito & ")"
   
    With rsCtaCteC
        .CursorLocation = adUseClient
        
        Call .Open(sqlCtaCteC, ConnDDBB, adOpenStatic, adLockPessimistic)
       
        If .EOF = True Then
            .AddNew
            .Fields("remito").Value = Trim(vnroremito)
            .Fields("comentario").Value = "Nro. " & TipoDocumento & " " & Trim(vnrocomprobante)
        End If
        
        .Fields("Fecha").Value = strfechaMySQL(dtpFecha.Value)
        .Fields("Fechainput").Value = strfechaMySQL(dtpFecha.Value)
        .Fields("Codigo").Value = txtClientes(0).Text
        .Fields("Nombre").Value = txtClientes(1).Text
       
        .Fields("AnoMes").Value = Right(.Fields("Fecha").Value, 4) & Mid(.Fields("Fecha").Value, 4, 2)
    
        If TipoDocumento = "Nota C" Then   ' si es una nota de credito
            
            .Fields("credito").Value = Val(Format(txtTotal.Text, "#####0.00"))
            .Fields("debito").Value = 0
            .Fields("saldo") = 0 'bcliente.Recordset("saldo") - bfactura_temp.Recordset("Total")
            
        Else
            If (TipoDocumento = "Documento") Or (TipoDocumento = "Fact A") Or (TipoDocumento = "Fact B") Then
                .Fields("debito") = Val(Format(txtTotal.Text, "#####0.00"))
                .Fields("credito") = 0
                .Fields("saldo") = Val(Format(txtTotal.Text, "#####0.00")) 'bclientes.Recordset("saldo") + bfactura_temp.Recordset("Total")
                    
            End If
        
        End If
        
        .Update
        
    End With

If Err Then GrabarLog "wcorrientes", Err.Number & " " & Err.Description, Me.Name
End Sub

Public Sub PagarCtaCte(vnroremito As Long, importe As Double, idMedioPago As Integer)
    On Error Resume Next
    
    Dim rsCtaCteC As New ADODB.Recordset, sqlCtaCteC As String
    
    sqlCtaCteC = "SELECT * FROM CuentasCorrientes WHERE (remito = " & vnroremito & ")"
     
    With rsCtaCteC
        .CursorLocation = adUseClient
               
        Call .Open(sqlCtaCteC, ConnDDBB, adOpenStatic, adLockPessimistic)
        Dim SaldoAnterior, debito, credito As Double
        Do While Not .EOF
            If IsNull(.Fields("debito")) Then
                debito = 0
            Else
                debito = .Fields("debito")
            End If
            
            If IsNull(.Fields("credito")) Then
                credito = 0
            Else
                credito = .Fields("credito")
            End If
                        
            SaldoAnterior = SaldoAnterior + debito - credito
            
            .MoveNext
        Loop
        
        If .RecordCount > 0 Then
            .MoveLast
        End If
        
        'If .EOF = True Then
            .AddNew
            .Fields("remito").Value = Trim(vnroremito)
            .Fields("comentario").Value = "Nro. " & TipoDocumento & " " & Trim(vnrocomprobante)
        'End If
        
        .Fields("Fecha").Value = strfechaMySQL(dtpFecha.Value)
        .Fields("Fechainput").Value = strfechaMySQL(dtpFecha.Value)
        
        'Buscar el cliente segun el remito
        .Fields("Codigo").Value = ObtenerCodigoClienteDesdeCtaCte(vnroremito)
        .Fields("Nombre").Value = ObtenerNombreClienteDesdeCtaCte(vnroremito)
       
        .Fields("anomes").Value = Right(.Fields("Fecha").Value, 4) & Mid(.Fields("Fecha").Value, 4, 2)
    
        .Fields("idMedioPago") = idMedioPago

        If (TipoDocumento = "Documento") Or (TipoDocumento = "Fact A") Or (TipoDocumento = "Fact B") Then
            
            .Fields("debito") = 0
            .Fields("credito") = importe
            .Fields("saldo") = SaldoAnterior - .Fields("credito") 'bclientes.Recordset("saldo") + bfactura_temp.Recordset("Total")
                    
        End If
        
        .Update
        
    End With

If Err Then GrabarLog "wcorrientes", Err.Number & " " & Err.Description, Me.Name
End Sub


'Metodo obsoleto
Private Function TipoVenta(vTipoVenta As String, vnroremito As Long) As Boolean
On Error Resume Next

    Select Case Trim(vTipoVenta)
        
        Case "Cuenta Corriente"
            WCtaCte (vnroremito)
        
        Case "Cheques"
            WCtaCte (vnroremito)
            'wcheques
        
        Case "Contado"
            'WCaja(
        
        Case "Credito"
            'wcredito
    
    End Select
    
If Err Then GrabarLog "TipoVenta", Err.Number & " " & Err.Description, Me.Name
End Function
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
        .ColWidth(0) = 125
        
        'Aca Pego el IdFDetalle-Entonces se si modifico o NO
        .TextMatrix(0, 1) = "idFDetalle"
        .ColWidth(1) = 0
        
        .TextMatrix(0, 2) = "Fecha"
        .ColWidth(3) = 0
        
        .TextMatrix(0, 3) = "Remito"
        .ColWidth(2) = 0
        
        .TextMatrix(0, 4) = "Codigo"
        .ColWidth(4) = 0
        
        .TextMatrix(0, 5) = "Cant."
        .ColWidth(5) = 750
        .ColDisplayFormat(5) = "#0.00"
        
        .TextMatrix(0, 6) = "Detalle"
        .ColWidth(6) = 3800
        
        .TextMatrix(0, 7) = "P. Venta"
        .ColWidth(7) = 850
        .ColDisplayFormat(7) = "#0.00"
        
        .TextMatrix(0, 8) = "% Desc."
        .ColWidth(8) = 850
        .ColDisplayFormat(8) = "#0.00"
                
        .TextMatrix(0, 9) = "% Iva"
        .ColWidth(9) = 850
        .ColDisplayFormat(9) = "#0.00"
        
        .TextMatrix(0, 10) = "% Imp."
        .ColWidth(10) = 850
        .ColDisplayFormat(10) = "#0.00"

        .TextMatrix(0, 11) = "$ Total"
        .ColWidth(11) = 900
        .ColDisplayFormat(11) = "#0.00"
        
        .ColWidth(25) = 200
        .TextMatrix(0, 25) = ""
        
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
Private Sub ImprimirTicket(vNroRemitoTicket As Long)
On Error Resume Next
    
    If vNroRemitoTicket = 0 Then
        MsgBox "El ticket no se puede imprimir, debido a errores internos!!!", vbExclamation, "Mensaje ..."
        Exit Sub
    End If
    
    Dim rsFactura As New ADODB.Recordset, sqlFactura As String
    Dim vnombre() As String, vdireccion() As String, vDocumento() As String
    Dim vTipoIva As String, vAliasIvaC As String, vAliasIvaE As String
    Dim vRespuesta As Boolean
    
    sqlFactura = "SELECT Factura.Codigo as CodCli, FDetalle.Codigo as CodArt, Factura.*, FDetalle.* FROM Factura INNER JOIN FDetalle ON Factura.Remito=FDetalle.Remito WHERE (Factura.Remito = " & vnroremito & ")"
       
    With rsFactura
        .CursorLocation = adUseClient
        Call .Open(sqlFactura, ConnDDBB, adOpenStatic, adLockPessimistic)
            
        If Not .State = 1 And .EOF = True Then
            MsgBox "No se podra Imprimir el Ticket por el siguiente Eror:  " & Err.Description
            Exit Sub
        End If
        
        FiscalEpson.BaudRate = 9600
        FiscalEpson.PortNumber = 1
        FiscalEpson.MessagesOn = True
        
        ReDim vnombre(1)
        ReDim vdireccion(2)
        ReDim vDocumento(1)
        
        vTipoIva = VerTipoIva("No", .Fields("CodCli").Value)
        vAliasIvaC = VerTipoIva("", .Fields("CodCli").Value)
        
        
        vAliasIvaE = TraerDato("Tipoiva", "TipoIva = '" & vDatosEmpresa.CondicionIva & "'", "AliasAfip")
        
        vnombre(0) = EsNuloGuion(Mid(.Fields("Nombre").Value, 1, 40))
        vnombre(1) = EsNuloGuion(Mid(.Fields("Nombre").Value, 41, 80))
        
        vdireccion(0) = EsNuloGuion(Mid(.Fields("Domicilio").Value, 1, 40))
        vdireccion(1) = EsNuloGuion(Mid(.Fields("Domicilio").Value, 41, 80))
        vdireccion(2) = EsNuloGuion(Mid(.Fields("Domicilio").Value, 81, 120))
        
        vDocumento(0) = EsNuloGuion(VerDocumento(vTipoIva, "T", .Fields("CodCli").Value))
        vDocumento(1) = EsNuloGuion(VerDocumento(vTipoIva, "N", .Fields("CodCli").Value))
        
        If vAliasIvaC = "I" Or vAliasIvaC = "M" Then
        
            'Panic: Faltan los IF
            If FiscalEpson.Status = True Then
                        
                Select Case vAliasIvaC
            
                    Case "I"
                        vRespuesta = FiscalEpson.OpenInvoice("T", "C", "A", "1", "P", "12", vAliasIvaE, vTipoIva, vnombre(0), vnombre(1), vDocumento(0), vDocumento(1), "", vdireccion(0), vdireccion(1), vdireccion(2), "", "", "G")
                
                    Case "M"
                        'No se como Hacerlo Actuar - Sugerencia Martin/Ramiro
                        vRespuesta = FiscalEpson.OpenInvoice("T", "C", "B", "1", "P", "12", vAliasIvaE, vTipoIva, vnombre(0), vnombre(1), vDocumento(0), vDocumento(1), "", vdireccion(0), vdireccion(1), vdireccion(2), "", "", "G")
                
                    Case "F"
                        'vRespuesta = FiscalEpson.OpenInvoice("T", "C", "B", "1", "P", "12", vAliasIvaE, vAliasAfip, vnombre(0), vnombre(1), vDocumento(0), vDocumento(1), "", vdireccion(0), vdireccion(1), vdireccion(2), "", "", "G")
                    
                End Select
            
                If vRespuesta = True Then
                    If vRespuesta = True Then vRespuesta = FiscalEpson.SendExtraDescription("")
                    Do Until .EOF = True
                        
                        vRespuesta = FiscalEpson.SendInvoiceItem(Mid(.Fields("Detalle").Value, 1, 20), FormatoNumeros(3, .Fields("Cantidad").Value), FormatoNumeros(2, .Fields("Precio").Value), VerIvaArticulo(.Fields("CodArt").Value), "M", "", "", Mid(.Fields("Detalle").Value, 21, 50), Mid(.Fields("Detalle").Value, 51, 80), Mid(.Fields("Detalle").Value, 81, 110), "0")
                        'vRespuesta = FiscalEpson.SendTicketItem(Mid(.Fields("Detalle").Value, 1, 20), FormatoNumeros(3, .Fields("Cantidad").Value), FormatoNumeros(2, .Fields("Precio").Value), VerIvaArticulo(.Fields("CodArt").Value), "M", "0", "0")
                    
                        .MoveNext
                    Loop
                
                    'Call FiscalEpson.SendInvoicePayment("Descuento por pagoA", "200.00", "D")
                    'Call FiscalEpson.SendInvoicePayment("Recargo por x motit", "150", "R")
        
                    vRespuesta = FiscalEpson.GetInvoiceSubtotal("P", "SUB TOT")
                    vRespuesta = FiscalEpson.SendInvoicePayment("Su Pago", FormatoNumeros(3, Val(txtEfectivo.Text)), "T")
                    
                    If vAliasIvaC = "I" Then
                        vRespuesta = FiscalEpson.CloseInvoice("T", "A", "Total")
                    Else
                        vRespuesta = FiscalEpson.CloseInvoice("T", "B", "Total")
                    End If
                
                End If
                
            End If
        Else
            If FiscalEpson.Status = True Then
                vRespuesta = FiscalEpson.OpenTicket("G")
            
                If vRespuesta Then
                    If vRespuesta = True Then vRespuesta = FiscalEpson.SendExtraDescription("")
                    
                    Do Until .EOF = True
                
                        vRespuesta = FiscalEpson.SendTicketItem(Mid(.Fields("Detalle").Value, 1, 20), FormatoNumeros(3, (.Fields("Cantidad").Value)), FormatoNumeros(2, .Fields("Precio").Value), VerIvaArticulo(.Fields("CodArt").Value), "M", "0", "0")
            
                        .MoveNext
                    Loop
                End If
        
                If vRespuesta Then vRespuesta = FiscalEpson.GetTicketSubtotal("P", "SUB TOT")
                If vRespuesta Then vRespuesta = FiscalEpson.SendTicketPayment("Su Pago", FormatoNumeros(3, Val(txtEfectivo.Text)), "T")
                If vRespuesta Then vRespuesta = FiscalEpson.CloseTicket()
            End If
        End If
        
    End With
    
If Err Then GrabarLog "ImprimirTicket", Err.Number & " " & Err.Description, Me.Caption
End Sub
Function FormatoNumeros(vCantidad, vValor) As String
On Error Resume Next

    Dim vLongitud As Integer, vLongitud2 As Integer, vTraePunto As Integer, i As Integer
    
    FormatoNumeros = ""

    vLongitud = Len(vValor)
    
    vTraePunto = Val(InStr(1, vValor, "."))
    
    If vTraePunto > 0 Then
        
        vLongitud2 = Len(Mid(vValor, vTraePunto + 1, Val(vLongitud - vTraePunto)))
        vValor = Trim(Mid(vValor, vTraePunto + 1, Val(vLongitud - vTraePunto)))
        
        If vValor < 999 Then
            Select Case vLongitud2
                Case 0
                    vValor = vValor & "00"
                Case 1
                    vValor = "0" & vValor & "0"
                Case 2
                    vValor = "0" & vValor
            End Select
        End If
    
    Else
        For i = 0 To Val(vCantidad - vLongitud)
            vValor = vValor & "0"
        Next
    End If
    
    FormatoNumeros = vValor
    
If Err Then GrabarLog "FormatoNumeros", Err.Number & " " & Err.Description, Me.Caption
End Function
Private Function VerIvaArticulo(vCodigoArticulo As String) As String
On Error Resume Next

    Dim rsPorcentajeIva As New ADODB.Recordset, sqlPorcentajeIva As String
    
    sqlPorcentajeIva = "SELECT * FROM Articulos A INNER JOIN PorcentajeIva P ON A.IdPorcentajeIva=P.idPorcentajeIva WHERE (A.Codigo = '" & vCodigoArticulo & "')"
    
    With rsPorcentajeIva
        Call .Open(sqlPorcentajeIva, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If (.State = 1) And Not (.EOF = True) Then
            
            VerIvaArticulo = Replace(.Fields("Porcentaje").Value, ".", "")
        
            VerIvaArticulo = VerIvaArticulo & "0"
        Else
            VerIvaArticulo = ""
        End If
    
    End With

    sqlPorcentajeIva = ""
    
    If rsPorcentajeIva.State = 1 Then
        rsPorcentajeIva.Close
        Set rsPorcentajeIva = Nothing
    End If
    
If Err Then GrabarLog "VerIvaArticulo", Err.Number & " " & Err.Description, Me.Caption
End Function
Private Function VerDocumento(ByRef vTipoIva, vTipoDocumento As String, vCodCliente As String) As String
On Error Resume Next
    
    Select Case vTipoIva
    
        Case "Cons. Final", "Consumidor Final"
            Select Case vTipoDocumento
            
                Case "T"
                    VerDocumento = TraerDato("Clientes", "Codigo = '" & vCodCliente & "'", "TipoDocumento")
                Case "N"
                    VerDocumento = TraerDato("Clientes", "Codigo = '" & vCodCliente & "'", "NroDocumento")
                
            End Select
            If vTipoDocumento = "T" Then
                
            Else
                
            End If
            
        Case "Exento"
            If vTipoDocumento = "T" Then
                VerDocumento = TraerDato("TipoIva", "TipoIva = '" & vTipoIva & "'", "AliasAfip")
            Else
                VerDocumento = TraerDato("Clientes", "Codigo = '" & vCodCliente & "'", "Cuit")
            End If
            
        Case "Iva Responsable Inscripto"
            If vTipoDocumento = "T" Then
                VerDocumento = TraerDato("TipoIva", "TipoIva = '" & vTipoIva & "'", "AliasAfip")
            Else
                VerDocumento = Replace(TraerDato("Clientes", "Codigo = '" & vCodCliente & "'", "Cuit"), "-", "")
            End If
    
    End Select
    
If Err Then GrabarLog "VerDocumento", Err.Number & " " & Err.Description, Me.Caption
End Function
Private Function VerTipoIva(ByRef vAliasAfip As String, vCodigoCliente As String) As String
On Error Resume Next

    Dim rsTipoIvaCliente As New ADODB.Recordset, sqlTipoIvaCliente As String
    
    sqlTipoIvaCliente = "SELECT * FROM Clientes CL INNER JOIN TipoIva TI ON CL.idTipoIva=TI.idTipoIva WHERE Codigo = '" & vCodigoCliente & "'"
    
    VerTipoIva = ""
    
    With rsTipoIvaCliente
        Call .Open(sqlTipoIvaCliente, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .State = 0 Then
            If Not .EOF = True Then
                If Not vAliasAfip = "" Then
                    VerTipoIva = .Fields("TipoIva").Value
                Else
                    VerTipoIva = .Fields("AliasAfip").Value
                End If
                
            End If
        End If

    End With
    
    sqlTipoIvaCliente = ""
    
    If rsTipoIvaCliente.State = 1 Then
        rsTipoIvaCliente.Close
        Set rsTipoIvaCliente = Nothing
    End If

If Err Then GrabarLog "VerTipoIva", Err.Number & " " & Err.Description, Me.Caption
End Function
