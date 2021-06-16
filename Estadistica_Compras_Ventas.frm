VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{63BEADB1-20E1-478A-9B40-DDDAFBF3624F}#1.0#0"; "bsGradientLabel.ocx"
Begin VB.Form frmEstadisticasGral 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Estadistica de Compras/Ventas/Caja Diaria"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   12630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   12630
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdGraficar 
      Caption         =   "Confeccionar Gráfica"
      Height          =   375
      Left            =   1800
      TabIndex        =   48
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir Reporte."
      Height          =   375
      Left            =   120
      TabIndex        =   49
      Top             =   4080
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc bventas_clientes 
      Height          =   330
      Left            =   3000
      Top             =   7800
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
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
      Caption         =   "bventas_clientes"
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
   Begin MSAdodcLib.Adodc bventas_articulos 
      Height          =   330
      Left            =   3000
      Top             =   7080
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
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
      ConnectStringType=   3
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
      Caption         =   "bventas_articulos"
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
   Begin MSAdodcLib.Adodc bcompras_proveedores 
      Height          =   330
      Left            =   0
      Top             =   7800
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
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
      Caption         =   "bcompras_proveedores"
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
   Begin MSAdodcLib.Adodc bcompras_rubros 
      Height          =   330
      Left            =   0
      Top             =   7440
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
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
      Caption         =   "bcompras_rubros"
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
   Begin VB.Frame fraMaximizar 
      Height          =   475
      Left            =   8160
      TabIndex        =   41
      Top             =   6240
      Width           =   3100
      Begin VB.ComboBox cboMaximizar 
         Height          =   315
         ItemData        =   "Estadistica_Compras_Ventas.frx":0000
         Left            =   1590
         List            =   "Estadistica_Compras_Ventas.frx":0028
         TabIndex        =   43
         Text            =   "Elegir Gráfica"
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdMaximizar 
         Caption         =   "Maximizar"
         Height          =   315
         Left            =   50
         TabIndex        =   42
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame fraFiltro 
      Height          =   645
      Left            =   30
      TabIndex        =   7
      Top             =   0
      Width           =   11565
      Begin VB.CommandButton cmdEjecutar 
         Caption         =   "Ejecutar !"
         Height          =   345
         Left            =   8910
         TabIndex        =   12
         Top             =   180
         Width           =   2445
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   315
         Left            =   3930
         TabIndex        =   10
         Top             =   210
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   71761921
         CurrentDate     =   38406
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   315
         Left            =   7290
         TabIndex        =   11
         Top             =   210
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   71761921
         CurrentDate     =   38406
      End
      Begin VB.Label lblFiltro 
         Caption         =   "Confeccionar estadísica desde la fecha :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2910
      End
      Begin VB.Label lblFiltro 
         Caption         =   "hasta la fecha :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   5730
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
   Begin TabDlg.SSTab TabEstadisticas 
      Height          =   6090
      Left            =   0
      TabIndex        =   0
      Top             =   705
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   10742
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "Ventas p/ Artículos"
      TabPicture(0)   =   "Estadistica_Compras_Ventas.frx":00C1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCArticulo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "bsTitulo(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "dgEstadistica(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtArticulo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "gsEstadistica(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Ventas p/ Rubros"
      TabPicture(1)   =   "Estadistica_Compras_Ventas.frx":00DD
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "gsEstadistica(1)"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(2)=   "Frame5"
      Tab(1).Control(3)=   "Frame14"
      Tab(1).Control(4)=   "dgEstadistica(1)"
      Tab(1).Control(5)=   "bsTitulo(1)"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Ventas p/ Clientes"
      TabPicture(2)   =   "Estadistica_Compras_Ventas.frx":00F9
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "gsEstadistica(2)"
      Tab(2).Control(1)=   "Frame7"
      Tab(2).Control(2)=   "FraTotalM"
      Tab(2).Control(3)=   "dgEstadistica(2)"
      Tab(2).Control(4)=   "bsTitulo(2)"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Compras p/ Artículos"
      TabPicture(3)   =   "Estadistica_Compras_Ventas.frx":0115
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "gsEstadistica(3)"
      Tab(3).Control(1)=   "vcarticulo"
      Tab(3).Control(2)=   "Frame8"
      Tab(3).Control(3)=   "Frame6"
      Tab(3).Control(4)=   "dgEstadistica(3)"
      Tab(3).Control(5)=   "bsTitulo(3)"
      Tab(3).Control(6)=   "lblPArticulo"
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "Compras p/ Rubros"
      TabPicture(4)   =   "Estadistica_Compras_Ventas.frx":0131
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "gsEstadistica(4)"
      Tab(4).Control(1)=   "dgEstadistica(4)"
      Tab(4).Control(2)=   "Frame9"
      Tab(4).Control(3)=   "Frame10"
      Tab(4).Control(4)=   "Frame15"
      Tab(4).Control(5)=   "bsTitulo(4)"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "Compras p/ Proveedor"
      TabPicture(5)   =   "Estadistica_Compras_Ventas.frx":014D
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "gsEstadistica(5)"
      Tab(5).Control(1)=   "dgEstadistica(5)"
      Tab(5).Control(2)=   "Frame12"
      Tab(5).Control(3)=   "FraTotalMonto"
      Tab(5).Control(4)=   "bsTitulo(5)"
      Tab(5).ControlCount=   5
      Begin MSChart20Lib.MSChart gsEstadistica 
         Height          =   2205
         Index           =   5
         Left            =   -74880
         OleObjectBlob   =   "Estadistica_Compras_Ventas.frx":0169
         TabIndex        =   55
         Top             =   3800
         Width           =   11415
      End
      Begin MSChart20Lib.MSChart gsEstadistica 
         Height          =   2205
         Index           =   4
         Left            =   -74880
         OleObjectBlob   =   "Estadistica_Compras_Ventas.frx":24BF
         TabIndex        =   54
         Top             =   3800
         Width           =   11415
      End
      Begin MSChart20Lib.MSChart gsEstadistica 
         Height          =   2205
         Index           =   3
         Left            =   -74880
         OleObjectBlob   =   "Estadistica_Compras_Ventas.frx":4815
         TabIndex        =   53
         Top             =   3800
         Width           =   11415
      End
      Begin MSChart20Lib.MSChart gsEstadistica 
         Height          =   2205
         Index           =   2
         Left            =   -74880
         OleObjectBlob   =   "Estadistica_Compras_Ventas.frx":6B6B
         TabIndex        =   52
         Top             =   3800
         Width           =   11415
      End
      Begin MSChart20Lib.MSChart gsEstadistica 
         Height          =   2205
         Index           =   1
         Left            =   -74880
         OleObjectBlob   =   "Estadistica_Compras_Ventas.frx":8EC1
         TabIndex        =   51
         Top             =   3800
         Width           =   11415
      End
      Begin MSChart20Lib.MSChart gsEstadistica 
         Height          =   2205
         Index           =   0
         Left            =   120
         OleObjectBlob   =   "Estadistica_Compras_Ventas.frx":B217
         TabIndex        =   50
         Top             =   3800
         Width           =   11415
      End
      Begin VB.Frame Frame4 
         Caption         =   "Total Cantidad :"
         Height          =   525
         Left            =   -67740
         TabIndex        =   17
         Top             =   3210
         Width           =   1395
         Begin VB.Label vr_cantidad 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   120
            TabIndex        =   18
            Top             =   180
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Total Cdo :"
         Height          =   525
         Left            =   -66360
         TabIndex        =   19
         Top             =   3210
         Width           =   1395
         Begin VB.Label vr_total 
            Alignment       =   2  'Center
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   120
            TabIndex        =   20
            Top             =   180
            Width           =   1155
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Total Cantidad :"
         Height          =   525
         Left            =   8640
         TabIndex        =   13
         Top             =   3210
         Width           =   1395
         Begin VB.Label va_cantidad 
            Alignment       =   2  'Center
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   210
            TabIndex        =   15
            Top             =   180
            Width           =   945
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Total Monto :"
         Height          =   525
         Left            =   -66360
         TabIndex        =   21
         Top             =   3210
         Width           =   1395
         Begin VB.Label vc_total 
            Alignment       =   2  'Center
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   120
            TabIndex        =   22
            Top             =   180
            Width           =   1095
         End
      End
      Begin VB.Frame FraTotalM 
         Caption         =   "Total M. Ctacte :"
         Height          =   525
         Left            =   -64980
         TabIndex        =   46
         Top             =   3210
         Width           =   1395
         Begin VB.Label vc_total_ctacte 
            Alignment       =   2  'Center
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   210
            TabIndex        =   47
            Top             =   180
            Width           =   975
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Total M. CtaCte :"
         Height          =   525
         Left            =   -64980
         TabIndex        =   37
         Top             =   3210
         Width           =   1395
         Begin VB.Label vr_total_ctacte 
            Alignment       =   2  'Center
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   120
            TabIndex        =   38
            Top             =   180
            Width           =   1155
         End
      End
      Begin VB.TextBox vcarticulo 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -71250
         TabIndex        =   35
         ToolTipText     =   "<Enter> Para filtrar"
         Top             =   3480
         Width           =   4755
      End
      Begin VB.TextBox txtArticulo 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3630
         TabIndex        =   33
         ToolTipText     =   "<Enter> Para filtrar"
         Top             =   3480
         Width           =   4755
      End
      Begin VB.Frame Frame3 
         Caption         =   "Total Monto :"
         Height          =   525
         Left            =   10020
         TabIndex        =   14
         Top             =   3210
         Width           =   1395
         Begin VB.Label va_total 
            Alignment       =   2  'Center
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   180
            TabIndex        =   16
            Top             =   180
            Width           =   1065
         End
      End
      Begin MSDataGridLib.DataGrid dgEstadistica 
         Bindings        =   "Estadistica_Compras_Ventas.frx":D56D
         Height          =   2385
         Index           =   0
         Left            =   90
         TabIndex        =   1
         Top             =   750
         Width           =   11450
         _ExtentX        =   20188
         _ExtentY        =   4207
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
      Begin MSDataGridLib.DataGrid dgEstadistica 
         Bindings        =   "Estadistica_Compras_Ventas.frx":D58D
         Height          =   2385
         Index           =   4
         Left            =   -74910
         TabIndex        =   5
         Top             =   750
         Width           =   11450
         _ExtentX        =   20188
         _ExtentY        =   4207
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
      Begin MSDataGridLib.DataGrid dgEstadistica 
         Bindings        =   "Estadistica_Compras_Ventas.frx":D5AB
         Height          =   2385
         Index           =   5
         Left            =   -74910
         TabIndex        =   6
         Top             =   750
         Width           =   11450
         _ExtentX        =   20188
         _ExtentY        =   4207
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
      Begin VB.Frame Frame8 
         Caption         =   "Total Monto :"
         Height          =   525
         Left            =   -64980
         TabIndex        =   25
         Top             =   3210
         Width           =   1395
         Begin VB.Label ca_total 
            Alignment       =   2  'Center
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   240
            TabIndex        =   26
            Top             =   180
            Width           =   915
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Total Cantidad :"
         Height          =   525
         Left            =   -66360
         TabIndex        =   23
         Top             =   3210
         Width           =   1395
         Begin VB.Label ca_cantidad 
            Alignment       =   2  'Center
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   360
            TabIndex        =   24
            Top             =   180
            Width           =   615
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Total Cdo :"
         Height          =   525
         Left            =   -66360
         TabIndex        =   31
         Top             =   3210
         Width           =   1395
         Begin VB.Label cp_total 
            Alignment       =   2  'Center
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   150
            TabIndex        =   32
            Top             =   240
            Width           =   1125
         End
      End
      Begin VB.Frame FraTotalMonto 
         Caption         =   "Total M. Ctacte :"
         Height          =   525
         Left            =   -64980
         TabIndex        =   44
         Top             =   3210
         Width           =   1395
         Begin VB.Label cp_total_ctacte 
            Alignment       =   2  'Center
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   120
            TabIndex        =   45
            Top             =   180
            Width           =   1155
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Total Cantidad :"
         Height          =   525
         Left            =   -67740
         TabIndex        =   27
         Top             =   3210
         Width           =   1395
         Begin VB.Label cr_cantidad 
            Alignment       =   2  'Center
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   270
            TabIndex        =   28
            Top             =   210
            Width           =   945
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Total Cdo :"
         Height          =   525
         Left            =   -66360
         TabIndex        =   29
         Top             =   3210
         Width           =   1395
         Begin VB.Label cr_total 
            Alignment       =   2  'Center
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   150
            TabIndex        =   30
            Top             =   180
            Width           =   1155
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Total M. CtaCte:"
         Height          =   525
         Left            =   -64980
         TabIndex        =   39
         Top             =   3210
         Width           =   1395
         Begin VB.Label cr_total_ctacte 
            Alignment       =   2  'Center
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   120
            TabIndex        =   40
            Top             =   180
            Width           =   1155
         End
      End
      Begin MSDataGridLib.DataGrid dgEstadistica 
         Bindings        =   "Estadistica_Compras_Ventas.frx":D5CE
         Height          =   2385
         Index           =   1
         Left            =   -74910
         TabIndex        =   2
         Top             =   750
         Width           =   11450
         _ExtentX        =   20188
         _ExtentY        =   4207
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
      Begin MSDataGridLib.DataGrid dgEstadistica 
         Bindings        =   "Estadistica_Compras_Ventas.frx":D5EB
         Height          =   2385
         Index           =   2
         Left            =   -74910
         TabIndex        =   3
         Top             =   750
         Width           =   11450
         _ExtentX        =   20188
         _ExtentY        =   4207
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
      Begin MSDataGridLib.DataGrid dgEstadistica 
         Bindings        =   "Estadistica_Compras_Ventas.frx":D60A
         Height          =   2385
         Index           =   3
         Left            =   -74910
         TabIndex        =   4
         Top             =   750
         Width           =   11450
         _ExtentX        =   20188
         _ExtentY        =   4207
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
      Begin Project1.bsGradientLabel bsTitulo 
         Height          =   345
         Index           =   0
         Left            =   90
         Top             =   390
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   609
         Caption         =   "Estadistica de Ventas por Articulos"
         BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionColour   =   33023
         Colour2         =   16777215
         CaptionAlignment=   1
         BorderStyle     =   2
      End
      Begin Project1.bsGradientLabel bsTitulo 
         Height          =   345
         Index           =   1
         Left            =   -74910
         Top             =   390
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   609
         Caption         =   "Estadistica de Ventas Por Rubros"
         BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionColour   =   33023
         Colour2         =   16777215
         CaptionAlignment=   1
         BorderStyle     =   2
      End
      Begin Project1.bsGradientLabel bsTitulo 
         Height          =   345
         Index           =   2
         Left            =   -74910
         Top             =   390
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   609
         Caption         =   "Estadistica General de Ventas por Clientes"
         BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionColour   =   33023
         Colour2         =   16777215
         CaptionAlignment=   1
         BorderStyle     =   2
      End
      Begin Project1.bsGradientLabel bsTitulo 
         Height          =   345
         Index           =   3
         Left            =   -74910
         Top             =   390
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   609
         Caption         =   "Estadistica de Compras por Articulos"
         BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionColour   =   33023
         Colour2         =   16777215
         CaptionAlignment=   1
         BorderStyle     =   2
      End
      Begin Project1.bsGradientLabel bsTitulo 
         Height          =   345
         Index           =   4
         Left            =   -74910
         Top             =   390
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   609
         Caption         =   "Estadistica de Compras Por Rubros"
         BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionColour   =   33023
         Colour2         =   16777215
         CaptionAlignment=   1
         BorderStyle     =   2
      End
      Begin Project1.bsGradientLabel bsTitulo 
         Height          =   345
         Index           =   5
         Left            =   -74910
         Top             =   390
         Width           =   11445
         _ExtentX        =   20188
         _ExtentY        =   609
         Caption         =   "Estadistica General de Compras por Proveedor"
         BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionColour   =   33023
         Colour2         =   16777215
         CaptionAlignment=   1
         BorderStyle     =   2
      End
      Begin VB.Label lblPArticulo 
         Alignment       =   2  'Center
         Caption         =   "Ing. Artículos que desea  filtrar :"
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
         Left            =   -71190
         TabIndex        =   36
         Top             =   3210
         Width           =   4635
      End
      Begin VB.Label lblCArticulo 
         Alignment       =   2  'Center
         Caption         =   "Ing. Artículos que desea  filtrar :"
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
         Left            =   3690
         TabIndex        =   34
         Top             =   3210
         Width           =   4635
      End
   End
   Begin MSAdodcLib.Adodc bventas_rubros 
      Height          =   330
      Left            =   3000
      Top             =   7440
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
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
      Caption         =   "bventas_rubros"
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
   Begin MSAdodcLib.Adodc bcompras_articulos 
      Height          =   330
      Left            =   0
      Top             =   7080
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
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
      Caption         =   "bcompras_articulos"
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
Attribute VB_Name = "frmEstadisticasGral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function CambiarGrafico(ByRef vGrafico As MSChart20Lib.MSChart)
    On Error Resume Next

    Select Case cboMaximizar.Text

        Case "Area (2D)"
            vGrafico.chartType = VtChChartType2dArea

        Case "Barra (2D)"
            vGrafico.chartType = VtChChartType2dBar

        Case "Combinacion (2D)"
            vGrafico.chartType = VtChChartType2dCombination

        Case "Linea (2D)"
            vGrafico.chartType = VtChChartType2dLine

        Case "Torta (2D)"
            vGrafico.chartType = VtChChartType2dPie

        Case "Paso (2D)"
            vGrafico.chartType = VtChChartType2dStep

        Case "XY (2D)"
            vGrafico.chartType = VtChChartType2dXY

        Case "Area (3D)"
            vGrafico.chartType = VtChChartType3dArea

        Case "Barra (3D)"
            vGrafico.chartType = VtChChartType3dBar

        Case "Combinacion (3D)"
            vGrafico.chartType = VtChChartType3dCombination

        Case "Linea (3D)"
            vGrafico.chartType = VtChChartType3dLine

        Case "Paso (3D)"
            vGrafico.chartType = VtChChartType3dStep
    End Select

    If Err Then GrabarLog "CambiarGrafico", Err.Description & " " & Err.Description, Me.Name
End Function
Private Sub GraficaCArticulos()
    On Error Resume Next
    Dim i As Integer
    Dim tcantidad, ttotal As Double

    If bventas_articulos.Recordset.RecordCount = 0 Then
        va_cantidad.Caption = ""
        va_total.Caption = ""
        Exit Sub
    End If
    
    bventas_articulos.Recordset.MoveLast
    gsEstadistica(TabEstadisticas.Tab).RowCount = bventas_articulos.Recordset.RecordCount
    bventas_articulos.Recordset.MoveFirst

    Do Until bventas_articulos.Recordset.EOF = True
    
        i = i + 1
        gsEstadistica(TabEstadisticas.Tab).Row = i

        gsEstadistica(TabEstadisticas.Tab).Column = 1
        gsEstadistica(TabEstadisticas.Tab).Data = Val(Format(bventas_articulos.Recordset("Cantidad"), "########0"))
        gsEstadistica(TabEstadisticas.Tab).RowLabel = Str(bventas_articulos.Recordset("UFecha"))
    
        tcantidad = tcantidad + Val(Format(bventas_articulos.Recordset("Cantidad"), "########0.00"))
        ttotal = ttotal + Val(Format(bventas_articulos.Recordset("TotalCdo").Value, "########0.00")) + Val(Format(bventas_articulos.Recordset("TotalCtaCte").Value, "########0.00"))
    
        gsEstadistica(TabEstadisticas.Tab).Column = 2
        gsEstadistica(TabEstadisticas.Tab).Data = Val(Format(bventas_articulos.Recordset("TotalCdo"), "########0.00"))
        gsEstadistica(TabEstadisticas.Tab).RowLabel = Str(bventas_articulos.Recordset("UFecha"))
    
        bventas_articulos.Recordset.MoveNext
  
    Loop

    va_cantidad.Caption = Format(tcantidad, "########0")
    va_total.Caption = Format(ttotal, "########0.00")

    If Err Then GrabarLog "GraficaCArticulos", Err.Description & " " & Err.Description, Me.Name
End Sub
Private Sub GraficaCRubros()
    On Error Resume Next

    Dim i As Integer
    Dim tcantidad, ttotal, ttotal_ctacte As Double

    If bventas_rubros.Recordset.RecordCount = 0 Then
        vr_cantidad.Caption = ""
        vr_total.Caption = ""
        vr_total_ctacte.Caption = ""
        bventas_rubros.Refresh
        Exit Sub
    End If
    
    bventas_rubros.Recordset.MoveLast
    gsEstadistica(TabEstadisticas.Tab).RowCount = bventas_rubros.Recordset.RecordCount
    bventas_rubros.Recordset.MoveFirst

    Do Until bventas_rubros.Recordset.EOF = True
    
        i = i + 1
        gsEstadistica(TabEstadisticas.Tab).Row = i

        gsEstadistica(TabEstadisticas.Tab).Column = 1
        gsEstadistica(TabEstadisticas.Tab).Data = Val(Format(bventas_rubros.Recordset("Cantidad"), "########0"))
        gsEstadistica(TabEstadisticas.Tab).RowLabel = Str(bventas_rubros.Recordset("UFecha"))
    
        tcantidad = tcantidad + Val(Format(bventas_rubros.Recordset("Cantidad"), "########0"))
        ttotal = ttotal + Val(Format(bventas_rubros.Recordset("TotalCdo"), "########0.00"))
    
        gsEstadistica(TabEstadisticas.Tab).Column = 2
        gsEstadistica(TabEstadisticas.Tab).Data = Val(Format(bventas_rubros.Recordset("TotalCdo"), "########0.00"))
        gsEstadistica(TabEstadisticas.Tab).RowLabel = Str(bventas_rubros.Recordset("UFecha"))
    
        gsEstadistica(TabEstadisticas.Tab).Column = 3
        gsEstadistica(TabEstadisticas.Tab).Data = Val(Format(bventas_rubros.Recordset("TotalCtaCte"), "########0.00"))
        gsEstadistica(TabEstadisticas.Tab).RowLabel = Str(bventas_rubros.Recordset("UFecha"))
        
        ttotal_ctacte = ttotal_ctacte + Val(Format(bventas_rubros.Recordset("TotalCtaCte"), "#######0.00"))
        
        bventas_rubros.Recordset.MoveNext
  
    Loop

    vr_cantidad.Caption = Format(tcantidad, "#####0")
    vr_total.Caption = Format(ttotal, "#####0.00")
    vr_total_ctacte.Caption = Format(ttotal_ctacte, "#####0.00")

    If Err Then GrabarLog "GraficaCRubros", Err.Description & " " & Err.Description, Me.Name
End Sub
Private Sub GraficaCVentas()
    On Error Resume Next
    Dim i As Integer
    Dim tcantidad, ttotal, ttotal_ctacte As Double

    If bventas_clientes.Recordset.RecordCount = 0 Then
        vc_total_ctacte.Caption = ""
        vc_total.Caption = ""
        Exit Sub
    End If
    
    bventas_clientes.Recordset.MoveLast
    gsEstadistica(TabEstadisticas.Tab).RowCount = bventas_clientes.Recordset.RecordCount
    bventas_clientes.Recordset.MoveFirst

    Do Until bventas_clientes.Recordset.EOF = True
    
        i = i + 1
        gsEstadistica(TabEstadisticas.Tab).Row = i

        ttotal = ttotal + Val(Format(bventas_clientes.Recordset("TotalCdo"), "########0.00"))
        ttotal_ctacte = ttotal_ctacte + Val(Format(bventas_clientes.Recordset("TotalCtaCte"), "########0.00"))

        gsEstadistica(TabEstadisticas.Tab).Column = 2
        gsEstadistica(TabEstadisticas.Tab).Data = Val(Format(bventas_clientes.Recordset("TotalCdo"), "########0.00"))
        gsEstadistica(TabEstadisticas.Tab).RowLabel = Str(bventas_clientes.Recordset("UFecha"))
    
        bventas_clientes.Recordset.MoveNext
  
    Loop

    vc_total.Caption = Format(ttotal, "#####0.00")
    vc_total_ctacte.Caption = Format(ttotal_ctacte, "######0.00")

    If Err Then GrabarLog "GraficaCVentas", Err.Description & " " & Err.Description, Me.Name
End Sub
Private Sub GraficaPArticulos()
    On Error Resume Next
    Dim i As Integer
    Dim tcantidad, ttotal As Double

    If bcompras_articulos.Recordset.RecordCount = 0 Then
        Exit Sub
    End If
    
    bcompras_articulos.Recordset.MoveLast
    gsEstadistica(TabEstadisticas.Tab).RowCount = bcompras_articulos.Recordset.RecordCount
    bcompras_articulos.Recordset.MoveFirst

    Do Until bcompras_articulos.Recordset.EOF
    
        i = i + 1
        gsEstadistica(TabEstadisticas.Tab).Row = i

        gsEstadistica(TabEstadisticas.Tab).Column = 1
        gsEstadistica(TabEstadisticas.Tab).Data = Val(Format(bcompras_articulos.Recordset("Cantidad"), "########0"))
        gsEstadistica(TabEstadisticas.Tab).RowLabel = Format(bcompras_articulos.Recordset("UFecha"), "dd/mm/yy")
        tcantidad = tcantidad + Val(Format(bcompras_articulos.Recordset("Cantidad"), "########0"))
    
        gsEstadistica(TabEstadisticas.Tab).Column = 2
        gsEstadistica(TabEstadisticas.Tab).Data = Val(Format(bcompras_articulos.Recordset("TotalCdo"), "########0.00"))
        gsEstadistica(TabEstadisticas.Tab).RowLabel = Format(bcompras_articulos.Recordset("UFecha"), "dd/mm/yy")
        ttotal = ttotal + Val(Format(bcompras_articulos.Recordset("TotalCdo"), "########0.00"))
    
        bcompras_articulos.Recordset.MoveNext
  
    Loop

    ca_cantidad.Caption = Format(tcantidad, "########0")
    ca_total.Caption = Format(ttotal, "########0.00")

    If Err Then GrabarLog "Command14_Click", Err.Description & " " & Err.Description, Me.Name
End Sub
Private Sub GraficaPCompras()
    On Error Resume Next
    Dim i As Integer
    Dim tcantidad, ttotal, ttotal_ctacte As Double

    If bcompras_proveedores.Recordset.RecordCount = 0 Then
        cp_total.Caption = ""
        cp_total_ctacte.Caption = ""
        Exit Sub
    End If
    
    bcompras_proveedores.Recordset.MoveLast
    gsEstadistica(TabEstadisticas.Tab).RowCount = bcompras_proveedores.Recordset.RecordCount
    bcompras_proveedores.Recordset.MoveFirst

    Do Until bcompras_proveedores.Recordset.EOF
    
        i = i + 1
        gsEstadistica(TabEstadisticas.Tab).Row = i
    
        gsEstadistica(TabEstadisticas.Tab).Column = 2
        
        gsEstadistica(TabEstadisticas.Tab).Data = Val(Format(bcompras_proveedores.Recordset("TotalCdo"), "########0.00"))
        gsEstadistica(TabEstadisticas.Tab).RowLabel = Str(bcompras_proveedores.Recordset("UFecha"))
        ttotal = ttotal + Val(Format(bcompras_proveedores.Recordset("TotalCdo"), "########0.00"))
        
        ttotal_ctacte = ttotal_ctacte + Val(Format(bcompras_proveedores.Recordset("TotalCtaCte"), "########0.00"))
        bcompras_proveedores.Recordset.MoveNext
  
    Loop

    cp_total.Caption = Format(ttotal, "########0.00")
    cp_total_ctacte = Format(ttotal_ctacte, "########0.00")

    If Err Then GrabarLog "Command16_Click", Err.Description & " " & Err.Description, Me.Name
End Sub
Private Sub GraficaPRubros()
    On Error Resume Next
    Dim i As Integer
    Dim tcantidad, ttotal, ttotal_ctacte As Double
    
    If bcompras_rubros.Recordset.RecordCount = 0 Then
        cr_cantidad.Caption = ""
        cr_total.Caption = ""
        cr_total_ctacte.Caption = ""
        bcompras_rubros.Refresh
        Exit Sub
    End If
    
    bcompras_rubros.Recordset.MoveLast
    gsEstadistica(TabEstadisticas.Tab).RowCount = bcompras_rubros.Recordset.RecordCount
    bcompras_rubros.Recordset.MoveFirst

    Do Until bcompras_rubros.Recordset.EOF
    
        i = i + 1
        gsEstadistica(TabEstadisticas.Tab).Row = i

        gsEstadistica(TabEstadisticas.Tab).Column = 1
        gsEstadistica(TabEstadisticas.Tab).Data = Val(Format(bcompras_rubros.Recordset("Cantidad"), "########0"))
        gsEstadistica(TabEstadisticas.Tab).RowLabel = Str(bcompras_rubros.Recordset("UFecha"))
        tcantidad = tcantidad + Val(Format(bcompras_rubros.Recordset("Cantidad"), "########0"))
    
        gsEstadistica(TabEstadisticas.Tab).Column = 2
        gsEstadistica(TabEstadisticas.Tab).Data = Val(Format(bcompras_rubros.Recordset("TotalCdo"), "########0"))
        gsEstadistica(TabEstadisticas.Tab).RowLabel = Str(bcompras_rubros.Recordset("ÚltimoDeFecha"))
    
        ttotal = ttotal + Val(Format(bcompras_rubros.Recordset("TotalCdo"), "########0.00"))
        ttotal_ctacte = ttotal_ctacte + Val(Format(bcompras_rubros.Recordset("TotalCtaCte"), "########0.00"))
    
        bcompras_rubros.Recordset.MoveNext
  
    Loop

    cr_cantidad.Caption = Format(tcantidad, "########0")
    cr_total.Caption = Format(ttotal, "########0.00")
    cr_total_ctacte.Caption = Format(ttotal_ctacte, "########0.00")

    If Err Then GrabarLog "Command15_Click", Err.Description & " " & Err.Description, Me.Name
End Sub
Private Sub cmdEjecutar_Click()
    On Error Resume Next

    Dim a, b, c, d As String
    Dim e As String
    Dim f As String
    
    MousePointer = vbHourglass
    
    a = ""
    b = ""
    c = ""
    d = ""
    'e = ""
    'f = ""
    
    a = "(fdetalle.Fecha >= '" & strfechaMySQL(dtpDesde) & "' And fdetalle.Fecha <= '" & strfechaMySQL(dtpHasta) & "')" ' Filtro de fecha Cliente
    b = "(pfdetalle.Fecha >= '" & strfechaMySQL(dtpDesde) & "' And pfdetalle.Fecha <= '" & strfechaMySQL(dtpHasta) & "')" ' Filtro de fecha Cliente
    c = "(Factura.Fecha >= '" & strfechaMySQL(dtpDesde) & "' And Factura.Fecha <= '" & strfechaMySQL(dtpHasta) & "')" ' Filtro de fecha Cliente
    d = "(PFactura.Fecha >= '" & strfechaMySQL(dtpDesde) & "' And PFactura.Fecha <= '" & strfechaMySQL(dtpHasta) & "')" ' Filtro de fecha Cliente

    With bventas_articulos
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT Max(fdetalle.fecha) as UFecha ,Articulos.Codigo, Articulos.Descrip, Sum(fdetalle.Cantidad) AS Cantidad, Sum(fdetalle.Total_cdo) AS TotalCdo, Sum(fdetalle.total_ctacte) as TotalCtaCte FROM Articulos INNER JOIN fdetalle ON Articulos.Codigo = fdetalle.Codigo WHERE ((" + a + ")) GROUP BY Articulos.Codigo, Articulos.Descrip"
        .Refresh
    End With
    
    With bventas_rubros
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT Max(factura.fecha) as UFecha, Rubros.idRubros, Rubros.Rubro, Sum(fdetalle.Cantidad) AS Cantidad, Sum(fdetalle.Total_cdo) AS TotalCdo, Sum(fdetalle.total_ctacte) as TotalCtaCte, Max(Factura.Cventa) AS UCventa FROM Factura INNER JOIN (Rubros INNER JOIN (Articulos INNER JOIN fdetalle ON Articulos.Codigo = fdetalle.Codigo) ON Rubros.idRubros = Articulos.idRubros) ON Factura.Remito = fdetalle.Remito WHERE ((" + a + ")) GROUP BY Rubros.idRubros, Rubros.Rubro"
        .Refresh
    End With
    
    With bventas_clientes
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT Max(factura.fecha) as UFecha,Clientes.Codigo, Clientes.Nombre, Sum(Factura.Total_cdo) AS TotalCdo, Sum(Factura.Total_CtaCte) AS TotalCtaCte, Max(Factura.CVenta) as UCVenta FROM Factura INNER JOIN Clientes ON Factura.Codigo = Clientes.Codigo WHERE ((" + c + ")) GROUP BY Clientes.Codigo, Clientes.Nombre"
        .Refresh
    End With
        
    With bcompras_articulos
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT Max(pfdetalle.fecha) as UFecha, Articulos.Codigo, Articulos.Descrip, Sum(PFDetalle.Cantidad) AS Cantidad, Sum(PFDetalle.Total_cdo) AS TotalCdo, Sum(PFDetalle.Total_CtaCte) as TotalCtaCte FROM Articulos INNER JOIN PFDetalle ON Articulos.Codigo = PFDetalle.Codigo WHERE ((" + b + ")) GROUP BY Articulos.Codigo, Articulos.Descrip"
        .Refresh
    End With
    
    With bcompras_rubros
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT Max(pfactura.fecha) as UFecha, Rubros.idRubros, Rubros.Rubro, Sum(PFDetalle.Cantidad) AS Cantidad, Sum(PFDetalle.Total_cdo) AS TotalCdo, Sum(PFDetalle.total_ctacte) AS TotalCtaCte, Max(PFactura.CVenta) as UCVenta FROM (PFactura INNER JOIN (PFDetalle INNER JOIN Articulos ON PFDetalle.Codigo = Articulos.Codigo) ON PFactura.Remito = PFDetalle.Remito) INNER JOIN Rubros ON Articulos.idRubros = Rubros.idRubros WHERE ((" + b + ")) GROUP BY Rubros.idRubros, Rubros.Rubro"
        .Refresh
    End With
    
    With bcompras_proveedores
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT Max(Pfactura.fecha) as UFecha, Proveedores.Codigo, Proveedores.Nombre, Sum(PFactura.Total_Cdo) AS TotalCdo, Sum(PFactura.total_ctacte) AS TotalCtaCte, MAX(PFactura.CVenta) as UCVenta FROM PFactura INNER JOIN Proveedores ON PFactura.Codigo = Proveedores.Codigo WHERE ((" + d + ")) GROUP BY Proveedores.Codigo, Proveedores.Nombre;"
        .Refresh
    End With
        
    FormatearGrilla
    
    MousePointer = vbDefault

    If Err Then GrabarLog "cmdEjecutar", Err.Description & " " & Err.Description, Me.Name
End Sub
Private Sub cmdGraficar_Click()
On Error Resume Next

    Select Case TabEstadisticas.Tab
    
        Case 0
            GraficaCArticulos
        Case 1
            GraficaCRubros
        Case 2
            GraficaCVentas
        Case 3
            GraficaPArticulos
        Case 4
            GraficaPRubros
        Case 5
            GraficaPCompras
        Case 6
        Case 7
        Case 8
    
    End Select

If Err Then GrabarLog "cmdGraficar_Click", Err.Description & " " & Err.Description, Me.Name
End Sub
Private Sub cmdImprimir_Click()
On Error Resume Next

    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox "Prepare la Impresora...", vbInformation, "Mensaje ..."
    
    Select Case TabEstadisticas.Tab
    
        Case 0
            With Mantenimiento.rsVentas_Articulos
                If .State = 1 Then .Close
                
                .Source = bventas_articulos.RecordSource
                
                If .State = 0 Then .Open
                .Close
                .Open
            
            End With
    
            With drventas_articulos
                .Show
            End With
        
        Case 1
            With Mantenimiento.rsVentas_Rubros
                If .State = 1 Then .Close
                
                .Source = bventas_rubros.RecordSource
                
                If .State = 0 Then .Open
                .Close
                .Open
            
            End With
    
            With drventas_rubros
                .Show
            End With
        
        Case 2
            With Mantenimiento.rsVentas_Clientes
                If .State = 1 Then .Close
                
                .Source = bventas_clientes.RecordSource
                
                If .State = 0 Then .Open
                .Close
                .Open
            
            End With
    
            With drventas_clientes
                .Show
            End With
        
        Case 3
            With Mantenimiento.rsCompras_Articulos
                If .State = 1 Then .Close
                
                .Source = bcompras_articulos.RecordSource
                
                If .State = 0 Then .Open
                .Close
                .Open
            
            End With
    
            With drcompras_articulos
                .Show
            End With
        
        Case 4
            With Mantenimiento.rsCompras_Rubros
                If .State = 1 Then .Close
                
                .Source = bcompras_rubros.RecordSource
                
                If .State = 0 Then .Open
                .Close
                .Open
            End With
    
            With drcompras_rubros
                .Show
            End With
        
        Case 5
            With Mantenimiento.rsCompras_Proveedores
                If .State = 1 Then .Close
                
                .Source = bcompras_proveedores.RecordSource
                
                If .State = 0 Then .Open
                .Close
                .Open
            
            End With
    
            With drcompras_proveedores
                .Show
            End With
        
        Case 6
        Case 7
        Case 8
    
    End Select

If Err Then GrabarLog "cmdImprimir_Click", Err.Description & " " & Err.Description, Me.Name
End Sub
Private Sub cmdMaximizar_Click()
    On Error Resume Next

    With gsEstadistica(TabEstadisticas.Tab)

        If .Width < TabEstadisticas.Width Then
            
            .Height = TabEstadisticas.Height
            .Width = TabEstadisticas.Width
            .Top = 0
            .Left = 0
            cmdMaximizar.Caption = "Minimizar"
            cmdImprimir.Visible = Not True
            cmdGraficar.Visible = Not True
        
        Else
            
            .Height = 2025
            .Width = 11445
            .Top = 3800
            .Left = 120
            cmdMaximizar.Caption = "Maximizar"
            cmdImprimir.Visible = True
            cmdGraficar.Visible = True
        
        End If

    End With

    If Err Then GrabarLog "cmdMaximizar_Click", Err.Description & " " & Err.Description, Me.Name
End Sub

Private Sub Form_Load()
On Error Resume Next

    dtpDesde.Value = #1/1/2009#
    dtpHasta.Value = Date - 1
    
    With Me
        .Top = 0
        .Left = 0
        .Height = 7185
        .Width = 11750
        .KeyPreview = True
        
    End With
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cboMaximizar_Click()
On Error Resume Next

    Call CambiarGrafico(gsEstadistica(TabEstadisticas.Tab))

If Err Then GrabarLog "graficas_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub txtarticulo_keypress(KeyAscii As Integer)
    On Error Resume Next

    If KeyAscii = 13 Then
        bventas_articulos.RecordSource = "SELECT * FROM ventas_articulos WHERE últimodefecha >= '" & strfechaMySQL(dtpDesde) & "' and últimodefecha <= '" & strfechaMySQL(dtpHasta) & "' and descrip like '%" + txtarticulo + "%'"
        bventas_articulos.Refresh
        
        txtarticulo = bventas_articulos.Recordset("Descrip").Value
    End If

    If Err Then GrabarLog "varticulo_KeyPress", Err.Description & " " & Err.Description, Me.Name
End Sub
Private Sub vcarticulo_KeyPress(KeyAscii As Integer)
    On Error Resume Next

    If KeyAscii = 13 Then
        bcompras_articulos.RecordSource = "select * from Compras_Articulos where últimodefecha >= '" & strfechaMySQL(dtpDesde) & "' and últimodefecha <= '" & strfechaMySQL(dtpHasta) & "' and descrip like '%" + vcarticulo + "%'"
        bcompras_articulos.Refresh
        vcarticulo.Text = bcompras_articulos.Recordset("descrip").Value
    End If

    If Err Then GrabarLog "vcarticulo_KeyPress", Err.Description & " " & Err.Description, Me.Name
End Sub
Private Sub FormatearGrilla()
On Error Resume Next

    Dim i As Integer
    
    For i = 0 To 5
        Select Case i
    
            Case 0
                With dgEstadistica(i)
                    .HeadLines = 2
                
                    .Columns(0).Caption = "Ult. Fecha"
                    .Columns(0).Width = 1250
                    
                    .Columns(1).Caption = "Cod. Articulo"
                    .Columns(1).Width = 1000
                    .Columns(1).Alignment = dbgCenter
                    
                    .Columns(2).Caption = "Descripcion"
                    .Columns(2).Width = 5000
                    
                    .Columns(3).Caption = "Cantidad"
                    .Columns(3).Alignment = dbgRight
                    .Columns(3).DataFormat.Format = "#####0.00"
                    .Columns(3).Width = 1000
                    
                    .Columns(4).Caption = "T. Contado"
                    .Columns(4).Alignment = dbgRight
                    .Columns(4).DataFormat.Format = "$ #####0.00"
                    .Columns(4).Width = 1000
                    
                    .Columns(5).Caption = "T. Cta. Cte"
                    .Columns(5).Alignment = dbgRight
                    .Columns(5).DataFormat.Format = "$ #####0.00"
                    .Columns(5).Width = 1000
                End With
        
            Case 1
                With dgEstadistica(i)
                    .HeadLines = 2
                    
                    .Columns(0).Caption = "Ult. Fecha"
                    .Columns(0).Width = 1250
                    
                    .Columns(1).Caption = "Cod. Articulo"
                    .Columns(1).Width = 1000
                    
                    .Columns(2).Caption = "Descripcion"
                    .Columns(2).Width = 3750
                    
                    .Columns(3).Caption = "Cantidad"
                    .Columns(3).Alignment = dbgRight
                    .Columns(3).DataFormat.Format = "#####0.00"
                    .Columns(3).Width = 1000
                    
                    .Columns(4).Caption = "T. Contado"
                    .Columns(4).Alignment = dbgRight
                    .Columns(4).DataFormat.Format = "$ #####0.00"
                    .Columns(4).Width = 1000
                    
                    .Columns(5).Caption = "T. Cta. Cte"
                    .Columns(5).Alignment = dbgRight
                    .Columns(5).DataFormat.Format = "$ #####0.00"
                    .Columns(5).Width = 1000
                    
                    .Columns(6).Caption = "Ult. C. Venta"
                    .Columns(6).Alignment = dbgRight
                End With
            
            Case 2
                With dgEstadistica(i)
                    .HeadLines = 2
                    
                    .Columns(0).Caption = "Ult. Fecha"
                    .Columns(0).Width = 1250
                    
                    .Columns(1).Caption = "Cod. Articulo"
                    .Columns(1).Width = 1000
                    
                    .Columns(2).Caption = "Cliente"
                    .Columns(2).Width = 4500
                    
                    .Columns(3).Caption = "T. Contado"
                    .Columns(3).Alignment = dbgRight
                    .Columns(3).DataFormat.Format = "$ #####0.00"
                    .Columns(3).Width = 1250
                    
                    .Columns(4).Caption = "T. Cta. Cte"
                    .Columns(4).Alignment = dbgRight
                    .Columns(4).DataFormat.Format = "$ #####0.00"
                    .Columns(4).Width = 1250
                
                    .Columns(5).Caption = "Ult. C. Venta"
                    .Columns(5).Alignment = dbgRight
                    .Columns(5).Width = 1500
                End With
        
            Case 3
                With dgEstadistica(i)
                    .HeadLines = 2
                
                    .Columns(0).Caption = "Ult. Fecha"
                    .Columns(0).Width = 1250
                    
                    .Columns(1).Caption = "Cod. Articulo"
                    .Columns(1).Width = 1000
                    .Columns(1).Alignment = dbgCenter
                    
                    .Columns(2).Caption = "Descripcion"
                    .Columns(2).Width = 5000
                    
                    .Columns(3).Caption = "Cantidad"
                    .Columns(3).Alignment = dbgRight
                    .Columns(3).DataFormat.Format = "#####0.00"
                    .Columns(3).Width = 1000
                    
                    .Columns(4).Caption = "T. Contado"
                    .Columns(4).Alignment = dbgRight
                    .Columns(4).DataFormat.Format = "$ #####0.00"
                    .Columns(4).Width = 1000
                    
                    .Columns(5).Caption = "T. Cta. Cte"
                    .Columns(5).Alignment = dbgRight
                    .Columns(5).DataFormat.Format = "$ #####0.00"
                    .Columns(5).Width = 1000
                End With
        
            Case 4
                With dgEstadistica(i)
                    .HeadLines = 2
                    
                    .Columns(0).Caption = "Ult. Fecha"
                    .Columns(0).Width = 1250
                    
                    .Columns(1).Caption = "Cod. Articulo"
                    .Columns(1).Width = 1000
                    
                    .Columns(2).Caption = "Descripcion"
                    .Columns(2).Width = 3750
                    
                    .Columns(3).Caption = "Cantidad"
                    .Columns(3).Alignment = dbgRight
                    .Columns(3).DataFormat.Format = "#####0.00"
                    .Columns(3).Width = 1000
                    
                    .Columns(4).Caption = "T. Contado"
                    .Columns(4).Alignment = dbgRight
                    .Columns(4).DataFormat.Format = "$ #####0.00"
                    .Columns(4).Width = 1000
                    
                    .Columns(5).Caption = "T. Cta. Cte"
                    .Columns(5).Alignment = dbgRight
                    .Columns(5).DataFormat.Format = "$ #####0.00"
                    .Columns(5).Width = 1000
                    
                    .Columns(6).Caption = "Ult. C. Venta"
                    .Columns(6).Alignment = dbgRight
                End With
            
            Case 5
                With dgEstadistica(i)
                    .HeadLines = 2
                    
                    .Columns(0).Caption = "Ult. Fecha"
                    .Columns(0).Width = 1250
                    
                    .Columns(1).Caption = "Cod. Articulo"
                    .Columns(1).Width = 1000
                    
                    .Columns(2).Caption = "Cliente"
                    .Columns(2).Width = 4500
                    
                    .Columns(3).Caption = "T. Contado"
                    .Columns(3).Alignment = dbgRight
                    .Columns(3).DataFormat.Format = "$ #####0.00"
                    .Columns(3).Width = 1250
                    
                    .Columns(4).Caption = "T. Cta. Cte"
                    .Columns(4).Alignment = dbgRight
                    .Columns(4).DataFormat.Format = "$ #####0.00"
                    .Columns(4).Width = 1250
                
                    .Columns(5).Caption = "Ult. C. Venta"
                    .Columns(5).Alignment = dbgRight
                    .Columns(5).Width = 1500
                End With
        End Select

    Next
    
    dgEstadistica(0).BackColor = &H80000018
    dgEstadistica(1).BackColor = &H80000018
    dgEstadistica(2).BackColor = &H80000018
    dgEstadistica(3).BackColor = &H80000018
    dgEstadistica(4).BackColor = &H80000018
    dgEstadistica(5).BackColor = &H80000018

If Err Then GrabarLog "FormatearGrilla", Err.Description & " " & Err.Description, Me.Name
End Sub
