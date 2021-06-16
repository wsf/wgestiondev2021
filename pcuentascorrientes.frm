VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "Copia de KlexGrid.ocx"
Begin VB.Form frmCtaCteP 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cuentas Corrientes de Proveedores"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   195
   ClientWidth     =   10080
   Icon            =   "pcuentascorrientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicInferior 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   60
      Picture         =   "pcuentascorrientes.frx":000C
      ScaleHeight     =   585
      ScaleWidth      =   10005
      TabIndex        =   58
      Top             =   6180
      Width           =   10005
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         Height          =   495
         Index           =   1
         Left            =   5460
         Picture         =   "pcuentascorrientes.frx":50BF
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   825
      End
      Begin VB.CommandButton cmdBorrar 
         Caption         =   "Borrar"
         Height          =   495
         Left            =   6270
         Picture         =   "pcuentascorrientes.frx":54C0
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   765
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         Height          =   495
         Index           =   1
         Left            =   7020
         Picture         =   "pcuentascorrientes.frx":58C5
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   765
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Pagos asignados"
         Height          =   495
         Index           =   2
         Left            =   7770
         Picture         =   "pcuentascorrientes.frx":5C9A
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   495
         Left            =   9210
         Picture         =   "pcuentascorrientes.frx":6072
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.Label lblWGESTION2010 
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
         TabIndex        =   65
         Top             =   180
         Width           =   1770
      End
      Begin VB.Label lblWGESTION2010 
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
         Left            =   150
         TabIndex        =   59
         Top             =   180
         Width           =   1770
      End
   End
   Begin XtremeSuiteControls.PushButton PusFiltrarDatos 
      Height          =   405
      Left            =   60
      TabIndex        =   57
      Top             =   1830
      Width           =   10005
      _Version        =   851968
      _ExtentX        =   17648
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Filtrar Datos"
      Appearance      =   6
      Picture         =   "pcuentascorrientes.frx":6445
   End
   Begin VB.CommandButton cmdFiltrar 
      Caption         =   "Filtrar"
      Height          =   375
      Left            =   5100
      Picture         =   "pcuentascorrientes.frx":69DF
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   5940
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1275
   End
   Begin XtremeSuiteControls.PushButton cmdVerDetalle 
      Height          =   495
      Index           =   0
      Left            =   3000
      TabIndex        =   54
      Top             =   7200
      Visible         =   0   'False
      Width           =   750
      _Version        =   851968
      _ExtentX        =   1323
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Ver Detalle"
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      ImageAlignment  =   4
      TextImageRelation=   4
   End
   Begin MSAdodcLib.Adodc bPCuentasCorrientes 
      Height          =   330
      Left            =   60
      Top             =   6540
      Visible         =   0   'False
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
      Caption         =   "bPCuentasCorrientes"
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
   Begin MSDataGridLib.DataGrid dgProveedores 
      Height          =   2055
      Left            =   5640
      TabIndex        =   51
      Top             =   6960
      Visible         =   0   'False
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   3625
      _Version        =   393216
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
   Begin VB.Frame Frame4 
      Caption         =   "Datos del proveedor :"
      ForeColor       =   &H00000080&
      Height          =   945
      Left            =   60
      TabIndex        =   24
      Top             =   90
      Width           =   4875
      Begin VB.CommandButton cmdNuevo 
         Height          =   315
         Index           =   0
         Left            =   4440
         Picture         =   "pcuentascorrientes.frx":6F69
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   285
      End
      Begin VB.CommandButton cmdSaldos 
         Height          =   285
         Left            =   4440
         Picture         =   "pcuentascorrientes.frx":706B
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Impre saldos de clientes"
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   285
      End
      Begin VB.TextBox txtProveedor 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1230
         TabIndex        =   0
         Top             =   270
         Width           =   3135
      End
      Begin VB.TextBox txtCuit 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1230
         TabIndex        =   1
         Top             =   600
         Width           =   3165
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "> C.U.I.T :"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "> Proveedor :"
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   25
         Top             =   330
         Width           =   1035
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Saldo Real :"
      ForeColor       =   &H00000000&
      Height          =   945
      Left            =   7600
      TabIndex        =   44
      Top             =   90
      Width           =   2415
      Begin VB.Label Saldo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataSource      =   "bcliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   435
         Left            =   180
         TabIndex        =   46
         Top             =   390
         Width           =   2115
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Saldo en Mora :"
      ForeColor       =   &H00000000&
      Height          =   945
      Left            =   5040
      TabIndex        =   43
      Top             =   90
      Width           =   2415
      Begin VB.Label rsaldo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataSource      =   "bcliente"
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
         Height          =   435
         Left            =   210
         TabIndex        =   45
         Top             =   390
         Width           =   1995
      End
   End
   Begin VB.Frame Frame5 
      Height          =   855
      Left            =   60
      TabIndex        =   27
      Top             =   975
      Width           =   9975
      Begin VB.CommandButton cmdVer 
         BackColor       =   &H007EE9FC&
         Height          =   285
         Index           =   0
         Left            =   9030
         Picture         =   "pcuentascorrientes.frx":716D
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Ver crédito"
         Top             =   210
         Width           =   255
      End
      Begin VB.CommandButton cmdVer 
         BackColor       =   &H007EE9FC&
         Height          =   285
         Index           =   1
         Left            =   9030
         Picture         =   "pcuentascorrientes.frx":726F
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Ver crédito"
         Top             =   510
         Width           =   255
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   285
         Left            =   1440
         TabIndex        =   30
         Top             =   180
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   503
         _Version        =   393216
         Format          =   70909953
         CurrentDate     =   38028
      End
      Begin MSComCtl2.DTPicker dhpHasta 
         Height          =   285
         Left            =   1440
         TabIndex        =   31
         Top             =   510
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   503
         _Version        =   393216
         Format          =   70909953
         CurrentDate     =   38028
      End
      Begin VB.Label saldocheque 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   7740
         TabIndex        =   39
         Top             =   150
         Width           =   1185
      End
      Begin VB.Label Label2 
         Caption         =   "> Saldo de Cheques no acreditados :"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   4470
         TabIndex        =   38
         Top             =   150
         Width           =   3585
      End
      Begin VB.Label Label7 
         Caption         =   "> Saldo de creditos otorgados :"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4470
         TabIndex        =   37
         Top             =   360
         Width           =   3585
      End
      Begin VB.Label credotorgado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   7740
         TabIndex        =   36
         Top             =   390
         Width           =   1185
      End
      Begin VB.Label vsanterior 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   7770
         TabIndex        =   35
         Top             =   570
         Width           =   1185
      End
      Begin VB.Label dsaldo 
         Appearance      =   0  'Flat
         Caption         =   "> Saldo anterior a la fecha "
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   4470
         TabIndex        =   34
         Top             =   600
         Width           =   3570
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         Caption         =   "> Fecha Hasta :"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   33
         Top             =   540
         Width           =   1545
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         Caption         =   "> Fecha Desde :"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   150
         TabIndex        =   32
         Top             =   210
         Width           =   1545
      End
   End
   Begin VB.Frame Frame3 
      Height          =   30
      Left            =   570
      TabIndex        =   23
      Top             =   1920
      Width           =   9465
   End
   Begin VB.CommandButton cmdActualizarSaldo 
      Caption         =   "Actualizar Saldo"
      Height          =   285
      Left            =   0
      TabIndex        =   11
      Top             =   7080
      Visible         =   0   'False
      Width           =   2745
   End
   Begin TabDlg.SSTab TabProveedor 
      Height          =   3915
      Left            =   60
      TabIndex        =   2
      Top             =   2250
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   6906
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Ing. Movimiento"
      TabPicture(0)   =   "pcuentascorrientes.frx":7371
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdImprimir(0)"
      Tab(0).Control(1)=   "cmdGuardar"
      Tab(0).Control(2)=   "Frame1"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Ficha Proveedor"
      TabPicture(1)   =   "pcuentascorrientes.frx":738D
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "dgMovimientos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "GBPagos"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Facturas de Proveedores"
      TabPicture(2)   =   "pcuentascorrientes.frx":73A9
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "vsaldo"
      Tab(2).Control(1)=   "asigna"
      Tab(2).Control(2)=   "dgAsignaciones"
      Tab(2).ControlCount=   3
      Begin XtremeSuiteControls.GroupBox GBPagos 
         Height          =   3465
         Left            =   60
         TabIndex        =   52
         Top             =   60
         Visible         =   0   'False
         Width           =   9795
         _Version        =   851968
         _ExtentX        =   17277
         _ExtentY        =   6112
         _StockProps     =   79
         Caption         =   "Detalle de Cancelacion de Factura"
         Appearance      =   1
         Begin Grid.KlexGrid KlexCobros 
            Height          =   2655
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   4683
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
            MouseIcon       =   "pcuentascorrientes.frx":73C5
            Rows            =   10
         End
         Begin XtremeSuiteControls.PushButton cmdSalirPagos 
            Height          =   375
            Left            =   8160
            TabIndex        =   55
            Top             =   3000
            Visible         =   0   'False
            Width           =   1470
            _Version        =   851968
            _ExtentX        =   2593
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Ocultar Detalle"
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            ImageAlignment  =   4
            TextImageRelation=   4
         End
      End
      Begin MSDataGridLib.DataGrid dgMovimientos 
         Height          =   3465
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   6112
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   16777215
         HeadLines       =   2
         RowHeight       =   19
         FormatLocked    =   -1  'True
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
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "Id"
            Caption         =   "Id"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
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
               LCID            =   1033
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
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Nombre"
            Caption         =   "Nombre"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Debito"
            Caption         =   "Debito"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Credito"
            Caption         =   "Credito"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Saldo"
            Caption         =   "Saldo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "Comentario"
            Caption         =   "Comentario"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "FechaCredito"
            Caption         =   "FechaCredito"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "Importe"
            Caption         =   "Importe"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "Acreditado"
            Caption         =   "Acreditado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "Ncredito"
            Caption         =   "Ncredito"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column12 
            DataField       =   "Remito"
            Caption         =   "Remito"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1484.787
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   1530.142
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1454.74
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   3240
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   0
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir Movimiento"
         Height          =   555
         Index           =   0
         Left            =   -72930
         Picture         =   "pcuentascorrientes.frx":73E1
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   2200
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar Movimiento"
         Height          =   555
         Left            =   -74520
         Picture         =   "pcuentascorrientes.frx":77E7
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   2200
         UseMaskColor    =   -1  'True
         Width           =   1605
      End
      Begin MSDataGridLib.DataGrid dgAsignaciones 
         Bindings        =   "pcuentascorrientes.frx":7C2D
         Height          =   2145
         Left            =   -74880
         TabIndex        =   12
         Top             =   30
         Width           =   9765
         _ExtentX        =   17224
         _ExtentY        =   3784
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   12648447
         HeadLines       =   2
         RowHeight       =   19
         FormatLocked    =   -1  'True
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
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   15
         BeginProperty Column00 
            DataField       =   "Id"
            Caption         =   "Id"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Fecha"
            Caption         =   "Fecha"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "d/M/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
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
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Nombre"
            Caption         =   "Nombre"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Debito"
            Caption         =   "Debito"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Credito"
            Caption         =   "Credito"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "PagoParcial"
            Caption         =   "PagoParcial"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "Saldo"
            Caption         =   "Saldo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "Comentario"
            Caption         =   "Comentario"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "FechaCredito"
            Caption         =   "FechaCredito"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "Importe"
            Caption         =   "Importe"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "Acreditado"
            Caption         =   "Acreditado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column12 
            DataField       =   "Ncredito"
            Caption         =   "Ncredito"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
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
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column14 
            DataField       =   "Concepto"
            Caption         =   "Concepto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1560.189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   4889.764
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
         EndProperty
      End
      Begin VB.Frame asigna 
         Height          =   690
         Left            =   -74880
         TabIndex        =   16
         Top             =   2490
         Visible         =   0   'False
         Width           =   9765
         Begin VB.CommandButton cmdAsignacion 
            Caption         =   "Deshacer Asignación"
            Height          =   525
            Index           =   1
            Left            =   1890
            Picture         =   "pcuentascorrientes.frx":7C4F
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   1815
         End
         Begin VB.Frame Frame2 
            Height          =   720
            Left            =   3900
            TabIndex        =   21
            Top             =   -30
            Width           =   15
         End
         Begin VB.CommandButton cmdAsignacion 
            Caption         =   "Asiganar Pago a Factura"
            Height          =   525
            Index           =   0
            Left            =   30
            Picture         =   "pcuentascorrientes.frx":803E
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   1875
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H007EE9FC&
            BackStyle       =   0  'Transparent
            Caption         =   "Importe para asignar a facturas:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   5025
            TabIndex        =   20
            Top             =   240
            Width           =   2820
         End
         Begin VB.Label vasigna 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   8070
            TabIndex        =   19
            Top             =   195
            Width           =   1500
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2025
         Left            =   -74880
         TabIndex        =   4
         Top             =   120
         Width           =   9705
         Begin VB.OptionButton o1 
            Caption         =   "Debitar <F11>"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3045
            TabIndex        =   50
            Top             =   720
            Width           =   1545
         End
         Begin VB.OptionButton o2 
            Caption         =   "Acreditar <F12>"
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   4785
            TabIndex        =   49
            Top             =   720
            Value           =   -1  'True
            Width           =   1785
         End
         Begin VB.CommandButton cmdCheques 
            Height          =   285
            Left            =   6720
            Picture         =   "pcuentascorrientes.frx":842A
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   720
            UseMaskColor    =   -1  'True
            Width           =   375
         End
         Begin VB.ComboBox CboConcepto 
            Height          =   315
            ItemData        =   "pcuentascorrientes.frx":852C
            Left            =   1650
            List            =   "pcuentascorrientes.frx":8539
            TabIndex        =   14
            Text            =   "Pago Factura (Automática)"
            Top             =   1110
            Width           =   4485
         End
         Begin VB.TextBox txtImporte 
            ForeColor       =   &H00000080&
            Height          =   315
            Left            =   1650
            TabIndex        =   7
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox txtComentario 
            Height          =   315
            Left            =   1650
            TabIndex        =   6
            Top             =   1530
            Width           =   7920
         End
         Begin MSComCtl2.DTPicker dtpAltaMovimiento 
            Height          =   315
            Left            =   1650
            TabIndex        =   5
            Top             =   330
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
            Format          =   70909953
            CurrentDate     =   38028
         End
         Begin VB.Label Label25 
            Appearance      =   0  'Flat
            BackColor       =   &H007EE9FC&
            BackStyle       =   0  'Transparent
            Caption         =   " << Efectuar pago con cheques "
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   7140
            TabIndex        =   22
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label Label20 
            Caption         =   "> Comentario :"
            Height          =   375
            Left            =   270
            TabIndex        =   13
            Top             =   1560
            Width           =   1755
         End
         Begin VB.Label Label5 
            Caption         =   "> Fecha :"
            Height          =   405
            Left            =   270
            TabIndex        =   10
            Top             =   390
            Width           =   1035
         End
         Begin VB.Label Label4 
            Caption         =   "> Concepto :"
            Height          =   375
            Left            =   270
            TabIndex        =   9
            Top             =   1140
            Width           =   1755
         End
         Begin VB.Label Label3 
            Caption         =   "> Importe :"
            Height          =   375
            Left            =   270
            TabIndex        =   8
            Top             =   750
            Width           =   1215
         End
      End
      Begin VB.Label vsaldo 
         Alignment       =   2  'Center
         Caption         =   "Saldo correspondiente hasta la fecha "
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -74940
         TabIndex        =   41
         Top             =   2310
         Width           =   9615
      End
   End
End
Attribute VB_Name = "frmCtaCteP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vvasigna As Double
Public vIdPCuentasCorrientes As Long
Dim rsProveedores As ADODB.Recordset
Dim vOpenGrilla As Boolean
Private Sub actualizar(vpago As Double)
On Error Resume Next

    Dim vcredito As Double
    
    With bPCuentasCorrientes
        .RecordSource = "SELECT * FROM pcuentascorrientes WHER (codigo = '" + Trim(txtProveedor.Tag) + "') ORDER BY fecha, idPCuentasCorrientes"
        .Refresh
    
        If .Recordset.EOF = True Then .Recordset.MoveFirst
    
        Do While vpago > 0
            vcredito = .Recordset("credito").Value

            Select Case vcredito
                Case Is <= vpago
                    .Recordset("PagoParcial") = bPCuentasCorrientes
                    .Recordset("credito") = 0
                    .Recordset.Update
                    .Recordset.MoveNext
                    .Refresh
                    vpago = vpago - 150

                Case Is > vpago
                    .Recordset("PagoParcial") = vpago
                    .Recordset("credito") = bPCuentasCorrientes - vpago
                    .Recordset.Update
                    .Recordset.MoveNext
                    .Refresh
                    vpago = vpago - 150
            End Select
        
        Loop
    
    End With
    

If Err Then GrabarLog "Actualizar", Err.Number & " " & Err.Description, Me.Name
End Sub

Public Sub asignapago() ' esta función asigna los apagos automaticamente

    Dim vvimporte As Double
    vvimporte = Val(txtImporte.Text)

    bPCuentasCorrientes.Recordset.MoveFirst

    Do Until bPCuentasCorrientes.Recordset.EOF Or vvimporte = 0

        If ((bPCuentasCorrientes.Recordset("credito").Value - bPCuentasCorrientes.Recordset("pagoparcial").Value) - vvimporte) > 0 Then
            ' no pagó el total
            bPCuentasCorrientes.Recordset("pagoparcial").Value = bPCuentasCorrientes.Recordset("pagoparcial").Value + vvimporte
            bPCuentasCorrientes.Recordset.Update
            vvimporte = 0
        Else
            ' le sobre plata
            vvimporte = ((bPCuentasCorrientes.Recordset("credito") - bPCuentasCorrientes.Recordset("pagoparcial")) - vvimporte) * -1
            bPCuentasCorrientes.Recordset("pagoparcial").Value = bPCuentasCorrientes.Recordset("credito").Value
            bPCuentasCorrientes.Recordset.Update
        End If
  
        bPCuentasCorrientes.Recordset.MoveNext
    Loop

End Sub
Private Function BuscarProveedor() As Boolean
On Error Resume Next

    If txtProveedor.Text = "" Then Exit Function
    
    Dim rsProveedores As New ADODB.Recordset, sqlProveedores As String
    
    sqlProveedores = "SELECT * FROM proveedores WHERE (nombre = '" & Trim(txtProveedor.Text) & "') OR (codigo = '" + Trim(txtProveedor.Text) + "')"
    
    With rsProveedores
        .CursorLocation = adUseClient
        Call .Open(sqlProveedores, ConnDDBB, adOpenStatic, adLockReadOnly)
    
        If .EOF Then
            'frmBuscarProveedor.Show
            'frmBuscarProveedor.txtProveedor = txtProveedor.Text
            'frmBuscarProveedor.TXTPROVEEDOR_KeyPress (13)
            'frmBuscarProveedor.txtProveedor.SetFocus
            'frmBuscarProveedor.o = 4
            BuscarProveedor = False
        Else
        
            'MousePointer = vbHourglass
    
            txtProveedor.Tag = EsNulo(.Fields("Codigo").Value)
            txtProveedor.Text = EsNulo(.Fields("Nombre").Value)
            txtCuit.Text = EsNulo(.Fields("cuit").Value)
            
            'txtImporte.SetFocus
            
           ' MousePointer = vbHourglass
            
            Exit Function
    
            With bPCuentasCorrientes
                .CursorLocation = adUseClient
                .CursorType = adOpenDynamic
                .LockType = adLockBatchOptimistic
                .ConnectionString = pathDBMySQL
                .RecordSource = "SELECT * FROM pcuentascorrientes WHERE (codigo = '" & Trim(txtProveedor.Tag) & "') ORDER BY fecha ASC, idPCuentasCorrientes ASC" 'and (debito + credito > 0) panic!!
                .Refresh
                
                If Not .Recordset.EOF Then
                    .Recordset.MoveFirst
                    dtpDesde.Value = strFecha(.Recordset("fecha").Value)
        
                    Call EjecutarScript("UPDATE PCuentasCorrientes SET Saldo = 0 WHERE codigo = '" & Trim(txtProveedor.Tag) & "';")
                    'SaldoAnterior (dtpDesde.Value)
                    CalcularSaldo (0)
                    CalcularDocumento 'Calculo los cheques que fueron dados al proveedor y no están acreditados
                Else
                    saldo.Caption = "0"
                End If
            End With
            
            BuscarProveedor = True
            
            MousePointer = vbDefault
    
        End If

    End With
    
    sqlProveedores = ""
    
    If rsProveedores.State = 1 Then
        rsProveedores.Close
        Set rsProveedores = Nothing
    End If
    
If Err Then GrabarLog "BuscarProveedor", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub CalcularCredito()
On Error Resume Next

    Dim saldoacreditado As Double
    bPCuentasCorrientes.RecordSource = "select * from pcuentascorrientes where fechacredito > '" & strfechaMySQL(Date) + "' and (codigo = '" + txtProveedor.Tag + "')"
    bPCuentasCorrientes.Refresh

    bPCuentasCorrientes.Recordset.MoveFirst

    Do Until bPCuentasCorrientes.Recordset.EOF

        If bPCuentasCorrientes.Recordset("acreditado").Value = "N" And dtpAltaMovimiento < bPCuentasCorrientes.Recordset("fechacredito").Value Then
            saldoacreditado = saldoacreditado + Val(Format(bPCuentasCorrientes.Recordset("importe").Value, "###########0.00"))
        End If

        bPCuentasCorrientes.Recordset.MoveNext
    Loop

    credotorgado.Caption = Str(saldoacreditado)


If Err Then GrabarLog "CalcularCredito", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub CalcularDocumento()
On Error Resume Next

    Dim vMontoTotal As Double
    
    vMontoTotal = Val(GenerarDato("SELECT Codigo, Estado, cp, Sum(Monto) as MontoTotal FROM cheques WHERE (estado = 'No Acreditado') AND (codigo = '" + txtProveedor.Tag + "') and (cp = 'p') GROUP BY Codigo, Estado, Cp", "MontoTotal"))

    saldocheque.Caption = Format(vMontoTotal, "#########0.00")

If Err Then GrabarLog "CalcularDocumento", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CalcularSaldo(vSaldoParcial As Double)
On Error Resume Next

' -------- desactivo el datagrid --------
'Set Me.dgMovimientos.DataSource = Nothing
'Me.dgAsignaciones.Refresh
'-----------------------------------------------------

    With bPCuentasCorrientes
        .Refresh
        If Not .Recordset.EOF = True Then .Recordset.MoveFirst
        '.Refresh
        
      
        If Not .Recordset.EOF = True Then .Recordset.MoveFirst

        Do Until .Recordset.EOF = True
        
            vSaldoParcial = vSaldoParcial + Val(Format(.Recordset("Credito").Value, "######0.00")) - Val(Format(.Recordset("debito").Value, "######0.00"))
            
            
            
            'Call EjecutarScript("update  pcuentascorrientes set saldo=" + Str(vSaldoParcial) + " where idPcuentascorrientes=" + Str(.Recordset("idPcuentascorrientes")))
            .Recordset("Saldo").Value = 0
            .Recordset("Saldo").Value = vSaldoParcial
            
           ' .Recordset.Update

            .Recordset.MoveNext

        Loop
        .Recordset.Update
       ' .Recordset.Fields.Refresh
        
        If Not .Recordset.EOF = True Then .Recordset.MoveLast
        
        saldo.Caption = Format(vSaldoParcial, "#######0.00")
        rsaldo.Caption = saldo.Caption
        
        
' -------- desactivo el datagrid --------
'Set Me.dgMovimientos.DataSource = Me.bPCuentasCorrientes
'Me.dgAsignaciones.Refresh
'-----------------------------------------------------
        
    End With
    
    If Err Then
        GrabarLog "CalcularSaldo", Err.Number & " " & Err.Description, Me.Name
        'MsgBox "Vuelva a filtrar por favor", vbInformation, "Filtrado no completado ..."
        Exit Sub
    End If
End Sub
Function CalcularSaldoParcial() As Double
On Error Resume Next

    Dim vSaldoTemp As Double

    vSaldoTemp = 0

    With bPCuentasCorrientes
        Do Until .Recordset.EOF
            vSaldoTemp = vSaldoTemp + Val(Format(.Recordset("credito").Value, "######0.00")) - Val(Format(.Recordset("pagoparcial").Value, "######0.00"))
            .Recordset.MoveNext
        Loop
    
    End With
    
    CalcularSaldoParcial = vSaldoTemp

If Err Then GrabarLog "CalcularSaldoParcial", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub cmdActualizarSaldo_Click()
On Error Resume Next

    CalcularSaldo (Val(vsanterior.Caption))

If Err Then GrabarLog "cmdActualizarSaldo_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdCheques_Click()
On Error Resume Next

    With frmCheques
        .dedonde = "pctacte"
        .vToBack = "pctacte"
        .TabCheques.tab = 0
    End With

If Err Then GrabarLog "cmdCheques_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdAsignacion_Click(Index As Integer)
    On Error Resume Next
    Dim vdif As Double

    If Index = 0 Then
    
        vdif = Val(Format(bPCuentasCorrientes.Recordset("credito").Value, "####0.00")) - Val(Format(bPCuentasCorrientes.Recordset("pagoparcial"), "####0.00")) - Val(Format(vvasigna, "####0.00"))

        If vdif > 0 Then
            bPCuentasCorrientes.Recordset("pagoparcial").Value = vvasigna
            bPCuentasCorrientes.Recordset.Update
            vvasigna = 0
            vasigna.Caption = "0"
        Else
    
            vvasigna = (bPCuentasCorrientes.Recordset("credito").Value - bPCuentasCorrientes.Recordset("pagoparcial").Value - vvasigna) * -1
            bPCuentasCorrientes.Recordset("pagoparcial").Value = bPCuentasCorrientes.Recordset("credito").Value
            bPCuentasCorrientes.Recordset.Update
            vasigna.Caption = Str(vvasigna)
        End If

    
    Else
        If MsgBox("¿Está seguro de deshacer la asignación a esta factura?", vbYesNo, "Deshaciendo pago a factura...") = vbYes Then
            vvasigna = vvasigna + bPCuentasCorrientes.Recordset("pagoparcial").Value
            vasigna = Str(vvasigna)
            bPCuentasCorrientes.Recordset("pagoparcial").Value = 0
            bPCuentasCorrientes.Recordset.Update
        End If

    End If

If Err Then
    MsgBox "Verifique si la factura fue seleccionada correctamente!", vbCritical, "Error..."
    GrabarLog "cmdAsignacion_Click", Err.Number & " " & Err.Description, Me.Name
End If
End Sub
Private Sub cmdSaldos_Click()
On Error Resume Next

    'frmSaldosProveedores.Show

If Err Then GrabarLog "cmdSaldos_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdSalir_Click()
On Error Resume Next

    Unload Me

If Err Then GrabarLog "cmdSalir_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdSalirPagos_Click()
    On Error Resume Next
    
    GBPagos.Visible = False
    
    If Err Then GrabarLog "cmdSalirPagos_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdFiltrar_Click()
On Error Resume Next

    vsanterior.Caption = Str(SaldoAnterior(dtpDesde))
    FiltrarMovimientos

If Err Then GrabarLog "cmdFiltrar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub GuardarDebito()
On Error Resume Next

    If Not Trim(txtImporte.Text) = "" Then
        
        With bPCuentasCorrientes
            .Recordset.AddNew
            
            .Recordset("Codigo").Value = EsNulo(txtProveedor.Tag)
            .Recordset("Nombre").Value = EsNulo(txtProveedor.Text)
            .Recordset("Fecha").Value = dtpAltaMovimiento.Value
            .Recordset("Debito").Value = Val(txtImporte.Text)
            .Recordset("Credito").Value = 0
            .Recordset("Comentario").Value = EsNulo(txtComentario.Text)
            .Recordset("idCheques").Value = EsNulo(txtComentario.Tag)
            .Recordset.Update

            CalcularSaldo (Val(vsanterior.Caption))
   
            frmCtaCteP.cmdGuardar.Enabled = True
            
            If Trim(CboConcepto) = "Pago Factura (Automática)" Then
                asignapago
            End If
   
            If Trim(CboConcepto.Text) = "Pago Factura (Manual)" Then
                PagoParcialManual
         
                .RecordSource = "SELECT * FROM pcuentascorrientes WHERE (Credito > pagoparcial) AND (debito = 0) ORDER BY fecha,idPCuentasCorrientes"
                .Refresh
                TabProveedor.tab = 2
            End If
     
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

    If Not txtImporte = "" Then
        bPCuentasCorrientes.Recordset.AddNew
        bPCuentasCorrientes.Recordset("codigo").Value = EsNulo(txtProveedor.Tag)
        bPCuentasCorrientes.Recordset("nombre").Value = EsNulo(txtProveedor.Text)
        bPCuentasCorrientes.Recordset("fecha").Value = strfechaMySQL(dtpAltaMovimiento.Value)
        bPCuentasCorrientes.Recordset("credito").Value = Val(txtImporte.Text)
        bPCuentasCorrientes.Recordset("comentario").Value = EsNulo(txtComentario.Text)
        bPCuentasCorrientes.Recordset.Update
        'bcliente.Recordset.Update
        cmdNuevo_Click (1)
        CalcularSaldo (Val(vsanterior.Caption))

        'If Trim(vconcepto) = "Pago Factura" Then
        '    pagoparcial
        'End If
        frmCtaCteP.cmdGuardar.Enabled = True
                    
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
        txtProveedor.Tag = ""
        txtProveedor.Text = ""
        txtCuit.Text = ""
        txtImporte.Text = ""
        txtComentario.Text = ""
        txtComentario.Tag = ""
        saldo.Caption = ""
        'rsaldo.Caption = ""
        txtProveedor.SetFocus
        
        With Me
            .Top = 300
            .Left = 300
            .Width = 10260
            '.Height = 1605
        End With
    
    Else
    
        txtImporte.Text = ""
        txtComentario.Text = ""
        txtComentario.Tag = ""
        txtImporte.SetFocus

    End If

If Err Then GrabarLog "cmdNuevo_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdBorrar_Click()
    On Error Resume Next
    Dim vmensaje, vborrado As String
    vmensaje = ""

    If MsgBox("Confirma la baja del movimiento de Cuenta Corriente del Proveedor ? ", vbYesNo) = vbNo Then
        Exit Sub
    End If

    Dim vArreglo As Double, vSaldoProveedor As Double

    With bPCuentasCorrientes
        If Not (.Recordset.EOF = True) And Not (.Recordset.BOF = True) Then
            vArreglo = Val(Format(.Recordset("debito").Value, "#######0.00")) - Val(Format(.Recordset("credito").Value, "#######0.00"))
            saldo.Caption = Str(Val(saldo.Caption) + vArreglo)
        
            If Not IsNull(.Recordset("idCheques").Value) = True Then
                Call BorrarBase("Cheques WHERE (idCheques = " & .Recordset("idCheques").Value & ")", pathDBMySQL)
                vmensaje = vmensaje + Chr(13) + "# Se ha borrado el movimiento de cheque asociado"
            End If
            If Not IsNull(.Recordset("ReMito").Value) = True Or Not (.Recordset("Remito").Value = 0) Then
                Call BorrarBase("PFactura WHERE (Remito = " & .Recordset("Remito").Value & ")", pathDBMySQL)
                Call BorrarBase("PFDetalle WHERE (Remito = " & .Recordset("Remito").Value & ")", pathDBMySQL)
                vmensaje = vmensaje + Chr(13) + "# Se ha borrado el documento asociado"
            End If
            
            
            vborrado = .Recordset("codigo") & "   " & Str(.Recordset("fecha")) & "  " & Str(.Recordset("nrointerno"))
            
            GrabarLog "Borrar.PCtaCte", vborrado, Me.Name
            
            Call BorrarBase("PCuentasCorrientes WHERE (IdPcuentascorrientes = " & .Recordset("IdPcuentascorrientes").Value & ")", pathDBMySQL)
            
            
            .Refresh
            Me.dgMovimientos.Refresh
            
            MsgBox "Fueron Borrado los siguientes datos: " + Chr(13) + vborrado + Chr(13) + vmensaje
            
            
            
        
        Else
            MsgBox "No tiene seleccionado ningun Movimiento...", vbExclamation, "Mensaje ...."
        End If
    
    End With
    
    

    
    vSaldoProveedor = 0
    vSaldoProveedor = Val(TraerDato("Proveedores", "Codigo = '" & Trim(txtProveedor.Tag) & "'", "Saldo")) + vArreglo
    
    'Call EjecutarScript("UPDATE Proveedores SET saldo = '" & vSaldoProveedor & "' WHERE (codigo = '" & Trim(txtProveedor.Tag) & "')")

    CalcularSaldo (0)

    'MsgBox "Esta accion NO BORRA ningun DOCUMENTO relacionado!!", vbExclamation, "Mensaje ...."

    If Err Then GrabarLog "cmdBorrar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdImprimir_Click(Index As Integer)
    On Error Resume Next
    
    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox "Prepare la impresora ...", vbInformation, "Mensaje ..."

    Select Case Index
    
        Case 0
            'Imprimir Movimiento de Pago
    
        Case 1
            'Listado de Movimientos
            With Mantenimiento.rspccc
                If Not .State = 0 Then .Close
                
                '.Source = bPCuentasCorrientes.RecordSource
                
                If Not .State = 1 Then .Open
                .Close
                .Open
 
                .Filter = "(Codigo = '" & txtProveedor.Tag & "') AND (Fecha <= '" & strfechaMySQL(dhpHasta.Value) + "' AND fecha >= '" & strfechaMySQL(dtpDesde.Value) + "')"
                .Sort = "Fecha ASC"
            End With
      
    
            With drpcuentascorrientes
                .Sections("TituloEmpresa").Controls("vcliente").Caption = txtProveedor.Tag + " " + txtProveedor
                .Sections("section3").Controls("vsaldo").Caption = saldo.Caption
                .Refresh
                .Show
            End With
        
        Case 2
            'Pagos Asignados
            With Mantenimiento.rscpagosasignados
                If Not .State = 1 Then .Open
                .Close
                .Open
 
                .Filter = "(Codigo = '" + Trim(txtProveedor.Tag) + "')"
                '.Sort = "Fecha ASC"
            End With
            
            'With drlasignaciones
            '    .Show
            'End With
    End Select


    If Err Then GrabarLog "cmdImprimir_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdGuardar_Click()
On Error Resume Next

     If o1.Value = True Then GuardarDebito
    If o2.Value = True Then GuardarCredito

If Err Then GrabarLog "cmdGuardar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdVer_Click(Index As Integer)
On Error Resume Next

    If Index = 0 Then
        With frmCheques
            .txtNombre = txtProveedor.Tag
            .txtNombre_KeyPress 13
            .TabCheques.tab = 2
        End With
    Else
        'With frmCreditos
        '    .cnombre = txtProveedor.Tag
        '    .d1.Value = False
        '    .Command7_Click
        '    .SSTab1.Tab = 1
        'End With
    End If
    
If Err Then GrabarLog "cmdVer_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub FiltrarMovimientos()
On Error Resume Next

Set Me.dgMovimientos.DataSource = Nothing
Set Me.dgAsignaciones.DataSource = Nothing
Me.dgMovimientos.Refresh
Me.dgAsignaciones.Refresh

    If Not txtProveedor.Tag = "" Then
        
        With bPCuentasCorrientes
            If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
            .RecordSource = "SELECT * FROM pcuentascorrientes WHERE (codigo ='" + txtProveedor.Tag + "') AND (fecha >= '" & strfechaMySQL(dtpDesde.Value) & "' and fecha <= '" & strfechaMySQL(dhpHasta.Value) & "') ORDER BY fecha ASC, idPCuentasCorrientes ASC"
            .Refresh
        
            CalcularSaldo (Val(vsanterior.Caption))
        
            vsaldo.Caption = "Saldo correspondiente hasta la fecha : " & Str(dhpHasta.Value) & " ....: " & Format(CalcularSaldoParcial, "###########0.00")
        
        End With
    
 Set Me.dgMovimientos.DataSource = Me.bPCuentasCorrientes
 Set Me.dgAsignaciones.DataSource = Me.bPCuentasCorrientes
 
Me.dgMovimientos.Refresh
Me.dgAsignaciones.Refresh
     
    End If

If Err Then GrabarLog "FiltrarMovimientos", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdVerDetalle_Click(Index As Integer)
On Error Resume Next


    If bPCuentasCorrientes.Recordset.EOF = True Or bPCuentasCorrientes.Recordset.BOF = True Then
        MsgBox "Debe Seleccionar un Movimiento para ver el Detalle ", vbExclamation, "Mensaje ..."
        Exit Sub
    End If
    
    If Not IsNull(bPCuentasCorrientes.Recordset("idCheques").Value) = True And Not Val(Format(bPCuentasCorrientes.Recordset("idCheques").Value, "#####0.00")) = 0 Then
        
        With frmCheques
            .opModo(4).Value = True
            .cvnombre.Text = bPCuentasCorrientes.Recordset("codigo").Value
            .cvnombre_KeyPress 13
            .cvncheque.Text = bPCuentasCorrientes.Recordset("ncheque").Value
            .cmdBuscar_Click
            '.Command7_Click
        End With
    
    Else
        If bPCuentasCorrientes.Recordset("remito").Value > 0 Then
            With bPCuentasCorrientes
                If Not .Recordset.EOF = True Then
                    CargarDetallePagos (.Recordset("Remito").Value)
                Else
        
                End If
            End With
            GBPagos.Visible = True
            'frmBuscarCompra.vViene = "pctacte"
            'frmBuscarCompra.vRemito = (bPCuentasCorrientes.Recordset("remito").Value)
            'frmBuscarCompra.Show
        
        Else
            With bPCuentasCorrientes
                If Not .Recordset.EOF = True Then
                    CargarDetallePagos (.Recordset("Remito").Value)
                Else
        
                End If
            End With
            GBPagos.Visible = True
        End If
    End If


If Err Then GrabarLog "cmdVerDetalle_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub dgMovimientos_DblClick()
On Error Resume Next
    
    With Me.bPCuentasCorrientes
        If Not .Recordset.EOF = True Then
            If IsNull(.Recordset("Remito").Value) = True Or Val(.Recordset("Remito").Value) = 0 Then
            
            Else
                CargarDetallePagos (Val(EsNulo(.Recordset("Remito").Value)))
            End If
        
        End If
    End With

If Err Then GrabarLog "dgMovimientos_DblClick", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub CargarDetallePagos(vnroremito As Long)
On Error Resume Next

    Dim rsCobros As New ADODB.Recordset, sqlCobros As String
    
    sqlCobros = "SELECT * FROM Pagos WHERE (Remito = " & vnroremito & ")"
    
    With rsCobros
        .CursorLocation = adUseClient
        
        Call .Open(sqlCobros, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then
            .MoveFirst
            FormatoGrillaPagos (.RecordCount)
        Else
            FormatoGrillaPagos (1)
        End If
        
        cmdSalirPagos.Visible = True
        
        
        Do Until .EOF = True
            DoEvents
            KlexCobros.TextMatrix(.AbsolutePosition, 1) = EsNulo(.Fields("idPagos").Value)
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
    
If Err Then GrabarLog "CargarDetallePagos", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub FormatoGrillaPagos(vCantidadRenglones As Integer)
On Error Resume Next

    Dim i As Integer

    GBPagos.Visible = True
    'GBPagos.Top = 500
    'GBPagos.Left = 60
    
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
        
        .TextMatrix(0, 1) = "idPagos"
        .ColWidth(1) = 1100
               
        .TextMatrix(0, 2) = "Fecha"
        .ColWidth(2) = 2340
        
        .TextMatrix(0, 3) = "Remito"
        .ColWidth(3) = 1500
        
        .TextMatrix(0, 4) = "idMedioPago"
        .ColWidth(4) = 1000
        
        .TextMatrix(0, 5) = "Importe"
        .ColWidth(5) = 1000
        .ColDisplayFormat(5) = "#0.00"
        
        .TextMatrix(0, 6) = "Tipo Movimiento"
        .ColWidth(6) = 1000
        
        .TextMatrix(0, 7) = "Nro Interno"
        .ColWidth(7) = 1000

    End With
    
If Err Then GrabarLog "FormatoGrillaPagos", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub dgProveedores_DblClick()
On Error Resume Next

    txtProveedor.Text = rsProveedores.Fields("Codigo").Value
    Call txtProveedor_KeyPress(13)
    dgProveedores.Visible = Not True
        
If Err Then GrabarLog "dgProveedores_DblClick", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub dgProveedores_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        dgProveedores_DblClick
    End If

If Err Then GrabarLog "dgProveedores_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, _
                       Shift As Integer)
    
    If KeyCode = vbKeyF1 Then cmdNuevo_Click (0)
    If KeyCode = vbKeyF11 Then o1.Value = True
    If KeyCode = vbKeyF12 Then o2.Value = True

    If KeyCode = 27 Then GBPagos.Visible = False


End Sub
Private Sub Form_Load()
    On Error Resume Next

    With Me
        .Top = 300
        .Left = 900
        .Width = 10170
        .Height = 7170
        .KeyPreview = True
    End With
    
    dtpAltaMovimiento.Value = Date
    dtpDesde.Value = Date
    dhpHasta.Value = Date
    saldo.Caption = ""

    If vIdPCuentasCorrientes > 0 Then
        TabProveedor.tab = 1
        
        With bPCuentasCorrientes
            If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
            .RecordSource = "SELECT * FROM pcuentascorrientes WHERE (idPCuentasCorrientes = " & Val(vIdPCuentasCorrientes) & ")"
            
            'Me.height = 6285
        End With
    End If

    TabProveedor.tab = 1

    If Err Then GrabarLog "Form_load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub o1_Click()
On Error Resume Next

    txtComentario.SetFocus
    
    If Err Then GrabarLog "o1_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub o2_Click()
On Error Resume Next

    txtComentario.SetFocus

If Err Then GrabarLog "o2_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub PagoParcial()
On Error Resume Next

    Dim vpago As Double
    
    vpago = Val(txtImporte.Text)
    
    actualizar (vpago)
    
    TabProveedor.tab = 2

If Err Then GrabarLog "PagoParcial", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub PagoParcialManual()
On Error Resume Next

    vvasigna = Val(txtImporte.Text)
    vasigna.Caption = Str(vvasigna)
    asigna.Visible = True

If Err Then GrabarLog "PagoParcialManual", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub PasarCredito()
On Error Resume Next

    Dim rsCCP As New ADODB.Recordset, sqlCCP As String

    sqlCCP = "SELECT * FROM pcuentascorrientes WHERE (fechacredito <= '" & strfechaMySQL(Date) + "') AND (Acreditado = 'N')"
    
    With rsCCP
        Call .Open(sqlCCP, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not .EOF = True Then .MoveFirst

        Do Until .EOF = True
            .Fields("Acreditado").Value = "S"
            .Fields("Debito").Value = .Fields("importe").Value
            .Fields("Credito").Value = 0
    
            .MoveNext
        Loop
    
    End With
    
    sqlCCP = ""
    
    If rsCCP.State = 1 Then
        rsCCP.Close
        Set rsCCP = Nothing
    End If

If Err Then GrabarLog "PasarCredito", Err.Number & " " & Err.Description, Me.Name
End Sub
Function SaldoAnterior(vfdesde As Date) As Double
    On Error Resume Next
    
    Dim vsaldoanterior As Double
    vsaldoanterior = 0

    With bPCuentasCorrientes
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM pcuentascorrientes WHERE (codigo = '" & txtProveedor.Tag & "') AND (Fecha < '" & strfechaMySQL(dtpDesde.Value) & "') ORDER BY fecha ASC,idPCuentasCorrientes ASC"
        .Refresh

        vsaldoanterior = 0
    
        If Not .Recordset.EOF = True Then .Recordset.MoveFirst
        
        Do Until .Recordset.EOF = True
            vsaldoanterior = vsaldoanterior + Val(Format(.Recordset("credito").Value, "#####0.00")) - Val(Format(.Recordset("debito").Value, "#####0.00"))
            .Recordset.MoveNext
        Loop
    
    End With
    
    rsaldo.Caption = vsaldoanterior
    
    SaldoAnterior = vsaldoanterior

    If Err Then GrabarLog "SaldoAnterior", Err.Number & " " & Err.Description, Me.Name
End Function

Private Sub PusFiltrarDatos_Click()
On Error Resume Next

    vsanterior.Caption = Str(SaldoAnterior(dtpDesde))
    FiltrarMovimientos

If Err Then GrabarLog "cmdFiltrar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub saldo_Change()
    saldo.Caption = Format(Val(saldo.Caption), "######0.00")
    rsaldo.Caption = Format(Val(saldo.Caption) - Val(saldocheque.Caption), "######0.00")
End Sub

Private Sub saldocheque_Change()
    rsaldo.Caption = Format(Val(saldo.Caption) - Val(saldocheque.Caption), "######0.00")
End Sub
Private Sub TabProveedor_Click(PreviousTab As Integer)
On Error Resume Next
    
    If TabProveedor.tab = 0 Then txtImporte.SetFocus

    If TabProveedor.tab = 1 Then FiltrarMovimientos ' filtra por fdesde, fhasta

    If TabProveedor.tab = 2 Then
        If CboConcepto.Text = "Pago Factura (Manual)" Then
            bPCuentasCorrientes.RecordSource = "SELECT * FROM pcuentascorrientes WHERE (Codigo = '" & txtProveedor.Tag & "') and (Credito > pagoparcial) and (debito = 0) order by fecha ASC,idPCuentasCorrientes ASC"
        Else
            bPCuentasCorrientes.RecordSource = "SELECT * FROM pcuentascorrientes WHERE (codigo = '" & txtProveedor.Tag & "') and (fecha >= '" & strfechaMySQL(dtpDesde.Value) + "' AND fecha <= '" & strfechaMySQL(dhpHasta.Value) + "') ORDER BY fecha ASC,idPCuentasCorrientes ASC"
        End If

        bPCuentasCorrientes.Refresh
        
        vsaldo.Caption = "Saldo correspondiente hasta la fecha : " & (dhpHasta.Value) & " ....: " & Format(CalcularSaldoParcial, "###########0.00")
    End If

    'asignapago
    
If Err Then GrabarLog "TabProveedor_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub vasigna_Change()
    vasigna = Format(vasigna.Caption, "######0.000")
End Sub
Public Sub txtProveedor_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        If BuscarProveedor = True Then
            dgProveedores.Visible = Not True
        End If
    End If

If Err Then GrabarLog "txtProveedor_Keypress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtProveedor_Change()
On Error Resume Next

    Call MostrarCoincidencias(txtProveedor.Text)
    vOpenGrilla = True
    
If Err Then GrabarLog "txtProveedor_Change", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub MostrarCoincidencias(vBusqueda As String)
On Error Resume Next

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
            Call FormatoGrilla
        Else
            Set dgProveedores.DataSource = Nothing
        End If
    
    End With

    sqlProveedores = ""

If Err Then GrabarLog "MostrarCoincidencias", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub FormatoGrilla()
On Error Resume Next
    
    Dim i As Integer

    With dgProveedores
    
        .ZOrder (0)
        .Top = 700
        .Left = 1350

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
    
If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtProveedor_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

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

    
    If KeyCode = 13 Then
        dgProveedores_DblClick
    End If
    
If Err Then GrabarLog "txtProveedor_KeyUp", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtProveedor_LostFocus()
    'BuscarProveedor
End Sub
Private Sub txtComentario_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdGuardar.SetFocus
    End If

End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        txtComentario.SetFocus
    End If

End Sub
Private Sub vsanterior_Change()
    dsaldo.Caption = "Saldo anterior a la fecha " + Str(dtpDesde.Value) + " :"
End Sub
