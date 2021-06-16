VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmDetalleFactura 
   Caption         =   "Detalle de la Factura:"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   13695
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8850
   ScaleWidth      =   13695
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.GroupBox BGImprimir 
      Height          =   1050
      Left            =   0
      TabIndex        =   27
      Top             =   8040
      Width           =   15135
      _Version        =   851968
      _ExtentX        =   26696
      _ExtentY        =   1852
      _StockProps     =   79
      Caption         =   "Impresion de Datos"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton PBImprimir 
         Height          =   495
         Index           =   0
         Left            =   360
         TabIndex        =   28
         Top             =   360
         Width           =   2400
         _Version        =   851968
         _ExtentX        =   4233
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Imprimir Todos los Detalles"
         UseVisualStyle  =   -1  'True
         Picture         =   "detalle_factura.frx":0000
      End
      Begin XtremeSuiteControls.PushButton PBImprimir 
         Height          =   495
         Index           =   1
         Left            =   2760
         TabIndex        =   29
         Top             =   360
         Width           =   2400
         _Version        =   851968
         _ExtentX        =   4233
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Imprimir Detalles Impagos"
         UseVisualStyle  =   -1  'True
         Picture         =   "detalle_factura.frx":0460
      End
      Begin XtremeSuiteControls.PushButton PBImprimir 
         Height          =   495
         Index           =   2
         Left            =   5160
         TabIndex        =   30
         Top             =   360
         Width           =   2400
         _Version        =   851968
         _ExtentX        =   4233
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Imprimir Detalles Pagos"
         UseVisualStyle  =   -1  'True
         Picture         =   "detalle_factura.frx":092D
      End
   End
   Begin XtremeSuiteControls.GroupBox GBDetalle 
      Height          =   1185
      Left            =   0
      TabIndex        =   4
      Top             =   9480
      Visible         =   0   'False
      Width           =   15135
      _Version        =   851968
      _ExtentX        =   26696
      _ExtentY        =   2090
      _StockProps     =   79
      Caption         =   "Modificar Detalle"
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox TxtDetalle 
         Alignment       =   2  'Center
         DataField       =   "Codigo"
         DataSource      =   "bfdetalle"
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxtDetalle 
         Alignment       =   2  'Center
         DataField       =   "Cantidad"
         DataSource      =   "bfdetalle"
         Height          =   315
         Index           =   1
         Left            =   1920
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtDetalle 
         Alignment       =   2  'Center
         DataField       =   "Detalle"
         DataSource      =   "bfdetalle"
         Height          =   315
         Index           =   2
         Left            =   2760
         TabIndex        =   20
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox TxtDetalle 
         Alignment       =   2  'Center
         DataField       =   "Precio"
         DataSource      =   "bfdetalle"
         Height          =   315
         Index           =   3
         Left            =   5880
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxtDetalle 
         Alignment       =   2  'Center
         DataField       =   "Descuento"
         DataSource      =   "bfdetalle"
         Height          =   315
         Index           =   4
         Left            =   6600
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox TxtDetalle 
         Alignment       =   2  'Center
         DataField       =   "Impuesto"
         DataSource      =   "bfdetalle"
         Height          =   315
         Index           =   5
         Left            =   7560
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TxtDetalle 
         Alignment       =   2  'Center
         DataField       =   "Total"
         DataSource      =   "bfdetalle"
         Height          =   315
         Index           =   6
         Left            =   8520
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxtDetalle 
         Alignment       =   2  'Center
         DataField       =   "total_cdo"
         DataSource      =   "bfdetalle"
         Height          =   315
         Index           =   7
         Left            =   9240
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtDetalle 
         Alignment       =   2  'Center
         DataField       =   "Total_ctacte"
         DataSource      =   "bfdetalle"
         Height          =   315
         Index           =   8
         Left            =   10080
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TxtDetalle 
         Alignment       =   2  'Center
         DataField       =   "confirmado"
         DataSource      =   "bfdetalle"
         Height          =   315
         Index           =   9
         Left            =   11160
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TxtDetalle 
         Alignment       =   2  'Center
         DataField       =   "Pagado"
         DataSource      =   "bfdetalle"
         Height          =   315
         Index           =   10
         Left            =   12240
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox TxtDetalle 
         Alignment       =   2  'Center
         DataField       =   "pago"
         DataSource      =   "bfdetalle"
         Height          =   315
         Index           =   11
         Left            =   12840
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox TxtDetalle 
         Alignment       =   2  'Center
         DataField       =   "resta"
         DataSource      =   "bfdetalle"
         Height          =   315
         Index           =   12
         Left            =   13320
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox TxtDetalle 
         Alignment       =   2  'Center
         DataField       =   "Totaliva"
         DataSource      =   "bfdetalle"
         Height          =   315
         Index           =   13
         Left            =   13800
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxtDetalle 
         Alignment       =   2  'Center
         DataField       =   "pespecial"
         DataSource      =   "bfdetalle"
         Height          =   315
         Index           =   17
         Left            =   7800
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TxtDetalle 
         Alignment       =   2  'Center
         DataField       =   "repartidor"
         DataSource      =   "bfdetalle"
         Height          =   315
         Index           =   16
         Left            =   5400
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TxtDetalle 
         Alignment       =   2  'Center
         DataField       =   "sueldo"
         DataSource      =   "bfdetalle"
         Height          =   315
         Index           =   15
         Left            =   3000
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TxtDetalle 
         Alignment       =   2  'Center
         DataField       =   "ganancia"
         DataSource      =   "bfdetalle"
         Height          =   315
         Index           =   14
         Left            =   960
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblDetalles 
         AutoSize        =   -1  'True
         Caption         =   "P. Especial:"
         Height          =   195
         Index           =   3
         Left            =   6840
         TabIndex        =   26
         Top             =   780
         Width           =   840
      End
      Begin VB.Label lblDetalles 
         AutoSize        =   -1  'True
         Caption         =   "Empleado :"
         Height          =   195
         Index           =   2
         Left            =   4440
         TabIndex        =   25
         Top             =   765
         Width           =   795
      End
      Begin VB.Label lblDetalles 
         AutoSize        =   -1  'True
         Caption         =   "Sueldo:"
         Height          =   195
         Index           =   1
         Left            =   2400
         TabIndex        =   24
         Top             =   765
         Width           =   540
      End
      Begin VB.Label lblDetalles 
         AutoSize        =   -1  'True
         Caption         =   "Ganancia:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   765
         Width           =   735
      End
   End
   Begin MSDataGridLib.DataGrid DgFacturas 
      Bindings        =   "detalle_factura.frx":0DF9
      Height          =   4425
      Left            =   90
      Negotiate       =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   7805
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483624
      HeadLines       =   2
      RowHeight       =   19
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   25
      BeginProperty Column00 
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
      BeginProperty Column01 
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
      BeginProperty Column02 
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
      BeginProperty Column03 
         DataField       =   "Cantidad"
         Caption         =   "Cantidad"
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
         DataField       =   "Detalle"
         Caption         =   "Detalle"
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
         DataField       =   "Precio"
         Caption         =   "Precio"
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
      BeginProperty Column07 
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
      BeginProperty Column08 
         DataField       =   "Total"
         Caption         =   "Total"
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
         DataField       =   "total_cdo"
         Caption         =   "total_cdo"
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
         DataField       =   "Total_ctacte"
         Caption         =   "Total_ctacte"
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
         DataField       =   "tiva"
         Caption         =   "tiva"
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
         DataField       =   "confirmado"
         Caption         =   "confirmado"
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
      BeginProperty Column14 
         DataField       =   "Envase"
         Caption         =   "Envase"
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
         DataField       =   "Pagado"
         Caption         =   "Pagado"
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
         DataField       =   "pago"
         Caption         =   "pago"
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
         DataField       =   "resta"
         Caption         =   "resta"
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
         DataField       =   "Totaliva"
         Caption         =   "Totaliva"
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
         DataField       =   "ganancia"
         Caption         =   "ganancia"
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
         DataField       =   "sueldo"
         Caption         =   "sueldo"
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
         DataField       =   "repartidor"
         Caption         =   "repartidor"
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
         DataField       =   "id_ctacte"
         Caption         =   "id_ctacte"
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
         DataField       =   "pespecial"
         Caption         =   "pespecial"
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
         DataField       =   "remito_ant"
         Caption         =   "remito_ant"
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
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   6045.166
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column24 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc bfdetalle 
      Height          =   330
      Left            =   15360
      Top             =   8160
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
   Begin MSDataGridLib.DataGrid DgFdetalle 
      Bindings        =   "detalle_factura.frx":0E10
      Height          =   2565
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   0
      Top             =   5400
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   4524
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   16777215
      HeadLines       =   2
      RowHeight       =   19
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   25
      BeginProperty Column00 
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
      BeginProperty Column01 
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
      BeginProperty Column02 
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
      BeginProperty Column03 
         DataField       =   "Cantidad"
         Caption         =   "Cantidad"
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
         DataField       =   "Detalle"
         Caption         =   "Detalle"
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
         DataField       =   "Precio"
         Caption         =   "Precio"
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
      BeginProperty Column07 
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
      BeginProperty Column08 
         DataField       =   "Total"
         Caption         =   "Total"
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
         DataField       =   "total_cdo"
         Caption         =   "total_cdo"
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
         DataField       =   "Total_ctacte"
         Caption         =   "Total_ctacte"
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
         DataField       =   "tiva"
         Caption         =   "tiva"
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
         DataField       =   "confirmado"
         Caption         =   "confirmado"
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
      BeginProperty Column14 
         DataField       =   "Envase"
         Caption         =   "Envase"
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
         DataField       =   "Pagado"
         Caption         =   "Pagado"
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
         DataField       =   "pago"
         Caption         =   "pago"
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
         DataField       =   "resta"
         Caption         =   "resta"
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
         DataField       =   "Totaliva"
         Caption         =   "Totaliva"
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
         DataField       =   "ganancia"
         Caption         =   "ganancia"
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
         DataField       =   "sueldo"
         Caption         =   "sueldo"
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
         DataField       =   "repartidor"
         Caption         =   "repartidor"
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
         DataField       =   "id_ctacte"
         Caption         =   "id_ctacte"
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
         DataField       =   "pespecial"
         Caption         =   "pespecial"
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
         DataField       =   "remito_ant"
         Caption         =   "remito_ant"
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
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   3060.284
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   540.284
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   480.189
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column24 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc bfactura 
      Height          =   330
      Left            =   15360
      Top             =   7800
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
   Begin VB.Label Label2 
      Caption         =   "Detalles de los documentos de venta seleccionado:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   435
      Left            =   30
      TabIndex        =   3
      Top             =   60
      Width           =   15135
   End
   Begin VB.Label Label1 
      Caption         =   "Seleccione un articulo haciendo doble click. Luego puede modificar los datos en la planilla a continuación:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   30
      TabIndex        =   2
      Top             =   5040
      Width           =   15135
   End
End
Attribute VB_Name = "frmDetalleFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vCodigoCliente As String
Private Sub DgFacturas_Click()
On Error Resume Next

    With bfdetalle
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM fdetalle WHERE (remito = " & bfactura.Recordset("remito") & ") ORDER BY fecha ASC, idFDetalle ASC"
        .Refresh
        
        If Not .Recordset.EOF = True Then .Recordset.MoveLast
    End With
    
If Err Then GrabarLog "DgFacturas_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub DgFacturas_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next

    OrdenarDataGrid ColIndex, bfactura.Recordset, DgFacturas

If Err Then GrabarLog "DgFacturas_HeadClick", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub DgFdetalle_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next

  OrdenarDataGrid ColIndex, bfdetalle.Recordset, Dgfdetalle

If Err Then GrabarLog "DgFdetalle_HeadClick", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub Form_Load()
On Error Resume Next
    
    WindowState = vbMaximized
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub PBImprimir_Click(Index As Integer)
On Error Resume Next
    
    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox "Prepare la Impresora !!!!", vbInformation, "Mensaje ..."
    
    With Mantenimiento.rsResumenFacturaCliente
        If .State = 1 Then .Close
        
        Select Case Index
        
            Case 0
                .Source = "SHAPE {SELECT * FROM Factura WHERE Codigo = '" & bfactura.Recordset("CodCli").Value & "'} AS ResumenFacturaCliente APPEND ({SELECT * FROM FDetalle} AS ResumenDetalleCliente RELATE 'Remito' TO 'Remito') AS ResumenDetalleCliente"
            
            Case 1
                .Source = "SHAPE {SELECT * FROM Factura WHERE Codigo = '" & bfactura.Recordset("CodCli").Value & "'} AS ResumenFacturaCliente APPEND ({SELECT * FROM FDetalle WHERE Resta > 0} AS ResumenDetalleCliente RELATE 'Remito' TO 'Remito') AS ResumenDetalleCliente"
            
            Case 2
                .Source = "SHAPE {SELECT * FROM Factura WHERE Codigo = '" & bfactura.Recordset("CodCli").Value & "'} AS ResumenFacturaCliente APPEND ({SELECT * FROM FDetalle WHERE Resta = 0} AS ResumenDetalleCliente RELATE 'Remito' TO 'Remito') AS ResumenDetalleCliente"
        
        End Select

        If .State = 0 Then .Open
        .Close
        .Open
        
        If .RecordCount = 0 Then
        
            Exit Sub
        End If
    End With

    With drDetalleCliente
        Select Case Index
        
            Case 0
                .Sections("TituloEmpresa").Controls("lblTitulo").Caption = "Listado de Cuentas Corrientes de Clientes con detalles"
            
            Case 1
                .Sections("TituloEmpresa").Controls("lblTitulo").Caption = "Listado de Cuentas Corrientes de Clientes con detalles IMPAGOS"
            
            Case 2
                .Sections("TituloEmpresa").Controls("lblTitulo").Caption = "Listado de Cuentas Corrientes de Clientes con detalles PAGOS"
        
        End Select
        

        .Sections("TituloEmpresa").Controls("txtcliente").Caption = bfactura.Recordset("Codigo").Value & " " & bfactura.Recordset("Nombre").Value
        .Sections("TituloEmpresa").Controls("txtLocalidad").Caption = bfactura.Recordset("Localidad").Value
        .Sections("TituloEmpresa").Controls("txtDireccion").Caption = bfactura.Recordset("Direccion").Value
        .Refresh
        .Show
    End With
    
If Err Then GrabarLog "PBImprimir_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
