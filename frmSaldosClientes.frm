VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmSaldosClientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Saldos_Clientes"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12375
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   12375
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc bquebranto 
      Height          =   330
      Left            =   8400
      Top             =   6000
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
      Caption         =   "bquebranto"
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
   Begin MSAdodcLib.Adodc blistas 
      Height          =   330
      Left            =   8400
      Top             =   5640
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
      Caption         =   "blistas"
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
      Left            =   8400
      Top             =   5280
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
   Begin VB.Frame framImprimir 
      Height          =   615
      Left            =   2040
      TabIndex        =   33
      Top             =   3840
      Width           =   3135
      Begin VB.TextBox vclidesde 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Top             =   210
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdverarticulos 
      Caption         =   "Ver Articulos"
      Height          =   385
      Left            =   9825
      TabIndex        =   30
      Top             =   3480
      Width           =   1000
   End
   Begin VB.CommandButton cmdborrarc 
      Caption         =   "Borrar"
      Height          =   385
      Left            =   8800
      TabIndex        =   32
      Top             =   3480
      Width           =   1000
   End
   Begin VB.CommandButton cmdfiltrar 
      Caption         =   "Filtrar!"
      Height          =   385
      Left            =   7800
      TabIndex        =   29
      Top             =   3480
      Width           =   1000
   End
   Begin MSComCtl2.DTPicker fquebrantos 
      Height          =   315
      Left            =   6240
      TabIndex        =   27
      Top             =   3480
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   61210625
      CurrentDate     =   39007
   End
   Begin MSDataGridLib.DataGrid Dgquebrantos 
      Bindings        =   "frmSaldosClientes.frx":0000
      Height          =   3375
      Left            =   5400
      TabIndex        =   26
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowDelete     =   -1  'True
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
         DataField       =   "codigo"
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
      BeginProperty Column01 
         DataField       =   "nombre"
         Caption         =   "Cliente"
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
         DataField       =   "U_venta"
         Caption         =   "U_venta"
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
         DataField       =   "U_pago"
         Caption         =   "U_pago"
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
         DataField       =   "Saldo"
         Caption         =   "Saldo"
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
         DataField       =   "Condicion"
         Caption         =   "Condicion"
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
         DataField       =   "Empleado"
         Caption         =   "Empleado"
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
         DataField       =   "Cod_empleado"
         Caption         =   "Cod_empleado"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2640.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   629.858
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cbolista 
      Height          =   315
      ItemData        =   "frmSaldosClientes.frx":001E
      Left            =   1380
      List            =   "frmSaldosClientes.frx":003D
      TabIndex        =   24
      Top             =   980
      Width           =   3615
   End
   Begin MSComctlLib.ProgressBar barra 
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   3480
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.ComboBox vlocalidad 
      Height          =   315
      Left            =   1380
      TabIndex        =   22
      Top             =   630
      Width           =   3615
   End
   Begin VB.TextBox vcdesde 
      Height          =   315
      Left            =   1380
      TabIndex        =   18
      Top             =   30
      Width           =   3585
   End
   Begin VB.TextBox vchasta 
      Height          =   315
      Left            =   1380
      TabIndex        =   17
      Top             =   330
      Width           =   3585
   End
   Begin MSAdodcLib.Adodc bLibreta 
      Height          =   330
      Left            =   8370
      Top             =   4560
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
      Caption         =   "bLibreta"
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
   Begin MSAdodcLib.Adodc breparto_agrupado 
      Height          =   330
      Left            =   5400
      Top             =   5640
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
      Caption         =   "breparto_agrupado"
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
   Begin MSAdodcLib.Adodc bReparto_Repartidor 
      Height          =   330
      Left            =   5400
      Top             =   5280
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
      Caption         =   "bReparto_Repartidor"
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
   Begin VB.Frame Frame3 
      Height          =   675
      Left            =   90
      TabIndex        =   15
      Top             =   3810
      Width           =   1910
      Begin VB.CommandButton cmdchangeprinter 
         Caption         =   "Elegir Impresora"
         Height          =   495
         Left            =   960
         TabIndex        =   31
         ToolTipText     =   "Generar reporte para imprimir"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Imprimir"
         Height          =   495
         Left            =   0
         Picture         =   "frmSaldosClientes.frx":0065
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Generar reporte para imprimir"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   915
      End
   End
   Begin VB.ComboBox vreparto 
      BackColor       =   &H80000016&
      Height          =   315
      Left            =   1380
      TabIndex        =   14
      Top             =   1320
      Width           =   3615
   End
   Begin VB.ComboBox tlistado 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmSaldosClientes.frx":0167
      Left            =   2550
      List            =   "frmSaldosClientes.frx":0174
      TabIndex        =   10
      Text            =   "Saldos"
      Top             =   2160
      Width           =   2595
   End
   Begin VB.Frame fecha 
      ForeColor       =   &H00000080&
      Height          =   825
      Left            =   2520
      TabIndex        =   5
      Top             =   2490
      Width           =   2595
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   285
         Left            =   750
         TabIndex        =   8
         Top             =   180
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   61210625
         CurrentDate     =   38238
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   285
         Left            =   750
         TabIndex        =   9
         Top             =   450
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   61210625
         CurrentDate     =   38238
      End
      Begin VB.Label Label5 
         Caption         =   "Desde :"
         Height          =   195
         Left            =   90
         TabIndex        =   7
         Top             =   210
         Width           =   705
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta :"
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   480
         Width           =   585
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   90
      TabIndex        =   0
      Top             =   2220
      Width           =   2355
      Begin VB.OptionButton Option1 
         Caption         =   "Deudores"
         Height          =   225
         Left            =   60
         TabIndex        =   4
         Top             =   150
         Value           =   -1  'True
         Width           =   2025
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Saldados"
         Height          =   225
         Left            =   60
         TabIndex        =   3
         Top             =   360
         Width           =   2025
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Saldos a favor de Clientes"
         Height          =   225
         Left            =   60
         TabIndex        =   2
         Top             =   570
         Width           =   2205
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Todos los Clientes"
         Height          =   225
         Left            =   60
         TabIndex        =   1
         Top             =   780
         Width           =   2055
      End
   End
   Begin MSAdodcLib.Adodc bcliente 
      Height          =   330
      Left            =   5400
      Top             =   4560
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
   Begin MSAdodcLib.Adodc bccliente 
      Height          =   330
      Left            =   5400
      Top             =   4920
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
      Caption         =   "bccliente"
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
   Begin MSAdodcLib.Adodc bfactura 
      Height          =   330
      Left            =   8400
      Top             =   4920
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
   Begin MSAdodcLib.Adodc btemp_quebranto 
      Height          =   330
      Left            =   5400
      Top             =   6360
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
      Caption         =   "btemp_quebranto"
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
   Begin MSAdodcLib.Adodc barticulos 
      Height          =   330
      Left            =   8400
      Top             =   6360
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
      Caption         =   "barticulos"
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
   Begin MSAdodcLib.Adodc bpagos_libretas 
      Height          =   330
      Left            =   5400
      Top             =   6720
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
      Caption         =   "bpagos_libretas"
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
   Begin VB.ListBox display 
      Height          =   1230
      Left            =   -2130
      TabIndex        =   35
      Top             =   5400
      Width           =   7365
   End
   Begin MSAdodcLib.Adodc bsaldos_clientes 
      Height          =   330
      Left            =   5400
      Top             =   6000
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
      Caption         =   "bsaldos_clientes"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8400
      Top             =   6720
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
      RecordSource    =   "saldos_clientes"
      Caption         =   "bsaldos_clientes"
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
   Begin VB.Label Label10 
      Caption         =   "Artículos que estaban pagos y quedan con sobrantes:"
      Height          =   375
      Left            =   150
      TabIndex        =   36
      Top             =   4680
      Width           =   5085
   End
   Begin VB.Label lbltultimopago 
      AutoSize        =   -1  'True
      Caption         =   "Ú. Pago:"
      Height          =   195
      Left            =   5520
      TabIndex        =   28
      Top             =   3525
      Width           =   630
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "> Lista :"
      Height          =   195
      Left            =   600
      TabIndex        =   25
      Top             =   1020
      Width           =   555
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "> Localidad :"
      Height          =   195
      Left            =   300
      TabIndex        =   21
      Top             =   690
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "> Cliente desde :"
      Height          =   195
      Left            =   60
      TabIndex        =   20
      Top             =   90
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "> Cliente hasta:"
      Height          =   195
      Left            =   60
      TabIndex        =   19
      Top             =   390
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "> Reparto:"
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
      Left            =   240
      TabIndex        =   13
      Top             =   1365
      Width           =   915
   End
   Begin VB.Label Label7 
      Caption         =   "Seleccione tipo de saldo:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   90
      TabIndex        =   12
      Top             =   1950
      Width           =   2355
   End
   Begin VB.Label Label6 
      Caption         =   "Tipo de Listado:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2580
      TabIndex        =   11
      Top             =   1950
      Width           =   2565
   End
End
Attribute VB_Name = "frmSaldosClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vcodigodesde, vcodigohasta, vcod_reparto, vSQL As String
Dim vcantidad, vDetalle, vtotal, vpago, vsaldo As String
Dim vtresta, vdiferencia, vblibreta_resta As Double
Dim lado As Integer
Dim vnlista As Integer
Dim vcodigodesdei As String 'Me quedo con el codigo del cliente para hacer la impresion desde ese punto
Dim vbdesde As Boolean
Function Busca_precio(ByRef vnlistas As Integer, vcodigoart As String) As Double
    With barticulos
        .Refresh
        .Recordset.Find ("Codigo = '" + vcodigoart + "'")
        If Not .Recordset.EOF = True Then
            Busca_precio = Val(Format(.Recordset("Pventa" & Val(vnlista)), "######0.00"))
        Else
            Busca_precio = 0
        End If
    End With
End Function

Private Sub buscacli(vcliente As String, _
                     dh As String)
    On Error Resume Next
    bcliente.RecordSource = "select * from clientes where (nombre = '" + vcliente + "') or (codigo = '" + vcliente + "')"
    bcliente.Refresh

    If bcliente.Recordset.EOF Then
    
        frmBuscarCliente.Show

        If dh = "d" Then
            frmBuscarCliente.o = 8
        Else
            frmBuscarCliente.o = 9
        End If
    
        frmBuscarCliente.Show
        frmBuscarCliente.txtClientes.Text = vcliente
        'frmBuscarCliente.varticulo_KeyPress (13)
        frmBuscarCliente.Show
        frmBuscarCliente.txtClientes.SetFocus
    
    Else
        Dim j As Integer
    
        If dh = "d" Then
            vcdesde = bcliente.Recordset("nombre").Value
            vcodigodesde = bcliente.Recordset(0).Value
            vchasta.SetFocus
        Else
            
            If dh = "j" Then
                
                vclidesde = bcliente.Recordset("nombre")
                vcodigodesdei = bcliente.Recordset(0)
                Command7.SetFocus
            Else
                vchasta = bcliente.Recordset("nombre")
                vcodigohasta = bcliente.Recordset(0)
                vlocalidad.SetFocus
            End If
        End If
        
    
    End If

    If Err Then GrabarLog "Buscacli", Err.Number & " " & Err.Description, Me.Name
End Sub

Function calsaldo(vCodigo As String)
    On Error Resume Next
    Dim vtotal As Double

    bccliente.RecordSource = "select * from cuentascorrientes where codigo = '" + vCodigo + "'"
    bccliente.Refresh

    vtotal = 0

    Do Until bccliente.Recordset.EOF
        vtotal = vtotal + bccliente.Recordset("debito") - bccliente.Recordset("credito")
        bccliente.Recordset.MoveNext
    Loop

    calsaldo = vtotal

    If Err Then GrabarLog "calsaldo", Err.Number & " " & Err.Description, Me.Name
End Function
Function calsaldoanterior(vcliente As String) As Double
    Dim vtotal As Double
    Dim vfhasta As Date
    On Error Resume Next
    vfhasta = fdesde.Value

    With bsaldos_clientes
        .RecordSource = "SELECT Max(cuentascorrientes.Fecha) AS ÚltimoDeFecha, cuentascorrientes.Codigo, Max(Clientes.Nombre) AS Nombre, Sum(cuentascorrientes.Debito) AS SumaDeDebito, Sum(cuentascorrientes.Credito) AS SumaDeCredito, Sum([debito]-[cuentascorrientes.credito]) AS Expr1, Clientes.Direccion, clientes.idtipoiva, Clientes.Cuit, Clientes.codigonum, Clientes.idReparto FROM cuentascorrientes INNER JOIN Clientes ON cuentascorrientes.Codigo = Clientes.Codigo Where (((CuentasCorrientes.Noimputar) = false ) And ((CuentasCorrientes.fechaInput) < '" & strfechaMySQL(vfhasta) + "')) GROUP BY cuentascorrientes.Codigo, Clientes.Direccion, clientes.idtipoiva, Clientes.Cuit, Clientes.codigonum, Clientes.idReparto HAVING (((cuentascorrientes.Codigo)='" + vcliente + "'))"
        .Refresh

        If .Recordset.RecordCount = 0 Then
            calsaldoanterior = 0
        Else
            calsaldoanterior = Val(Format(.Recordset("expr1"), "#######0.00"))
        End If
    End With
    If Err Then GrabarLog "calsaldoanterior", Err.Number & " " & Err.Description, Me.Name
End Function

Private Sub cboLista_Click()
    On Error Resume Next

    Select Case Trim(cbolista.Text)

        Case 3
            vnlista = 2

        Case 4
            vnlista = 3

        Case 5
            vnlista = 4

        Case 7
            vnlista = 5

        Case 8
            vnlista = 6

        Case 9
            vnlista = 1
            
        Case 6
            vnlista = 7
    End Select

    If Err Then GrabarLog "cbolista_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cboLista_GotFocus()
    On Error Resume Next

    With blistas
        .Refresh

        If Not .Recordset.RecordCount = 0 Then .Recordset.MoveFirst
        cbolista.Clear

        Do Until .Recordset.EOF = True
            cbolista.AddItem .Recordset("Lista")
            .Recordset.MoveNext
        Loop

    End With

    If Err Then GrabarLog "cbolista_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdborrarc_Click()
    On Error Resume Next
        With btemp_quebranto
            If Not .Recordset.RecordCount = 0 Then
                .Recordset.Delete
                .Recordset.Update
            End If
        End With
    If Err Then GrabarLog "cmdBorrarc_click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdchangeprinter_Click()
    frmChangePrinter.Show
End Sub
Private Sub cmdFiltrar_Click()
Dim i As Integer
On Error Resume Next
    If Trim(vreparto.Text) = "" Then
        MsgBox "Por favor ingrese un reparto!!", vbExclamation, "Mensaje ..."
        Exit Sub
    End If
    BorrarBase "Temp_quebranto", pathDBMySQL
    btemp_quebranto.Refresh
    With bquebranto
        .RecordSource = "Select * from quebranto where (u_pago <= '" & strfechaMySQL(fquebrantos.Value) + "' or u_pago is null)  and reparto = '" + Left(vreparto.Text, 2) + "'"
        .Refresh
        If Not .Recordset.RecordCount = 0 Then
            .Recordset.MoveFirst
            barra.Value = 0
            barra.Max = .Recordset.RecordCount
        End If
        Do Until .Recordset.EOF = True
            btemp_quebranto.Recordset.AddNew
            For i = 0 To 7
                btemp_quebranto.Recordset(i) = .Recordset(i)
            Next i
            btemp_quebranto.Recordset.Update
            .Recordset.MoveNext
            barra.Value = barra.Value + 1
        Loop
    End With
If Err Then GrabarLog "cmdfiltrar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdverarticulos_Click()
    frmPorcentaje.Show
End Sub

'-------------------------------------------------------------------------------------------------------------------
Private Sub Command7_Click()
 
    On Error Resume Next

    ' --------------------------------------------------
    If tlistado = "Para libreta" Then
        
        If Not Val(vcdesde) > 0 Then
        
        If vreparto.Text = "" Then
            MsgBox "Faltan seleccionar algunos campos importantes, para la generación de Libretas. ", vbInformation, "Mensaje ..."
            Exit Sub
        End If
        If cbolista.Text = "" Then
            MsgBox "Faltan seleccionar algunos campos importantes, para la generación de Libretas. ", vbInformation, "Mensaje ..."
            Exit Sub
        End If
        
        End If
        
        ' ----- seteo la impresora una sola vez para comenzar a imprimir  --------------------
        Printer.FontName = "Draft 10cpi"
        Printer.FontBold = False
        'printer.printQuality = -1 '---ERROR ACA SALTA EN LA SEGUNDA VUELTA
        'Printer.Height = 17280
        '------------------------------------------------------------------------------------
        MousePointer = vbHourglass
        flibreta
        MousePointer = vbDefault
        Exit Sub
    End If

    ' If tlistado = "Ficha c/ detalle" Then fcdetalle
    ' If tlistado = "Ficha s/ detalle" Then fsdetalle
    ' --------------------------------------------------

    If tlistado = "Saldos Detalles" Then
        idetallectacte fdesde, fhasta
        Exit Sub
    End If

    Dim vfiltro As String

    vfiltro = ""
    
    MousePointer = vbHourglass

    If Not (vcodigodesde = "" And vcodigohasta = "") Then
        vfiltro = vfiltro + " and codigo >= '" + vcodigodesde + "' and codigo <= '" + vcodigohasta + "'"
    End If

    If Not vlocalidad = "" Then
        vfiltro = vfiltro + " and localidad = '" + vlocalidad + "'"
    End If
    
    If Not vreparto.Text = "" Then
        vfiltro = vfiltro + " and reparto = '" + Trim(vcod_reparto) + "'"
    End If
    
    bcliente.Refresh



   If Option2 Then Mantenimiento.rsSaldo_Clientes.Filter = "saldo = 0" + vfiltro
   If Option3 Then Mantenimiento.rsSaldo_Clientes.Filter = "saldo < 0" + vfiltro
   If Option1 Then Mantenimiento.rsSaldo_Clientes.Filter = "saldo > 0" ' + vfiltro
   ' If Option4 Then Mantenimiento.rsSaldo_Clientes.filter = "id > 0 " + vfiltro

    If Not Mantenimiento.rsSaldo_Clientes.State = 1 Then
        Mantenimiento.rsSaldo_Clientes.Open
        Mantenimiento.rsSaldo_Clientes.Close
        Mantenimiento.rsSaldo_Clientes.Open
    Else
        Mantenimiento.rsSaldo_Clientes.Close
        Mantenimiento.rsSaldo_Clientes.Open
    End If

    Mantenimiento.rsSaldo_Clientes.Sort = "nombre"

    MsgBox "Prepare la impresora !", vbInformation, "Mensaje ..."

    With drClienteSaldo
      .Show
    End With
    
    Limpiar

    MousePointer = vbDefault

    If Err Then GrabarLog "Command7_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub detalle_Click()
    'If detalle.Value = 1 Then
    '    fecha.Visible = True
    'Else
    '    fecha.Visible = False
    'End If
End Sub

'-------------------------------------------------------------------------------------------------------------------
Private Sub fi_cliente(vcod_cliente As String)
    Dim renglon As Integer
    Dim vSaldoAnterior As Double
    On Error Resume Next
       
    'Miro si el cliente tiene Articulos para Imprimir en la consulta Asignacion
    bLibreta.RecordSource = "select * from libreta where CodCli = '" + vcod_cliente + "' and resta > 0 and (Fecha >= '" & strfechaMySQL(fdesde.Value) + "' and Fecha <= '" & strfechaMySQL(fhasta.Value) + "') order by Fecha ASC"
    bLibreta.Refresh
    
    'Busco el saldo anterior del cliente
    vSaldoAnterior = Val(Format(calsaldoanterior(vcod_cliente), "##########0.00"))
    
    'Si no tiene saldo anterior o no tiene articulos para imprimir entonces sale.
    If vSaldoAnterior <= 0 And bLibreta.Recordset.RecordCount = 0 Then Exit Sub
    
    'Imprime encabezado de la libreta
    fititulos bReparto_Repartidor.Recordset("codigo"), bReparto_Repartidor.Recordset("Nombre"), bReparto_Repartidor.Recordset("Direccion"), vSaldoAnterior
    
    If Not bLibreta.Recordset.RecordCount = 0 Then
        If vSaldoAnterior = 0 Then
            Printer.Print ""
            Printer.Print ""
        Else
            Printer.Print "                                                        Saldo Anterior: " & Format(vSaldoAnterior, "######0.00")
            Printer.Print ""
        End If
        
        bpagos_libretas.RecordSource = "select * from pagos_libretas where codigo ='" + vcod_cliente + "' and fecha <= '" & strfechaMySQL(fhasta.Value) + "' and fecha >= '" & strfechaMySQL(fdesde.Value) + "'"
        bpagos_libretas.Refresh
        If Not bpagos_libretas.Recordset.EOF Then
        
        renglon = 0
        Printer.Print "[Pagos efectuados:]"
        Do Until bpagos_libretas.Recordset.EOF = True
            Printer.Print ">>>>> " + Str(bpagos_libretas.Recordset("fecha")) & "    $ " & Format(num_i2(bpagos_libretas.Recordset("credito")), "########0.00")
            renglon = renglon + 1
            'fi_detalle
             bpagos_libretas.Recordset.MoveNext
        Loop
        ' -----------------------------------------------------------------------------------------------------------------------
        renglon = renglon + 1
        End If
        
        
        bLibreta.Recordset.MoveFirst
        Dim vtotalfactura As Double
        vtotalfactura = 0
        vtresta = 0
        
        ' -------------------------  Imprime pagos efectuados  ------------------------------------------------------------------
        
        Do Until bLibreta.Recordset.EOF = True
            fi_detalle
            bLibreta.Recordset.MoveNext
        Loop
        
        renglon = renglon + bLibreta.Recordset.RecordCount + 6
    Else
        Printer.Print "                                                          Saldo Anterior: " & Format(vSaldoAnterior, "######0.00")
        renglon = 5
    End If
    
    Printer.Print "--------------------------------------------------------------------------------"
    Printer.Print "                                                                Total: " & num_i(Format(vSaldoAnterior + vtresta, "#####0.00"))
    Printer.Print "--------------------------------------------------------------------------------"
    
    vtresta = 0 ' una vez impreso el saldo final se pone la variable en cero

    Dim i As Integer
    
    For i = (renglon) To 37
        Printer.Print ""
    Next i
    
    If Err Then GrabarLog "fi_cliente", Err.Number & " " & Err.Description, Me.Name
End Sub

Function fi_detalle() As Double
Dim vtotalnuevo, vadd_resta, vprecio_buscado As Double
Dim vresta_parcial, vprecio_modificado, vresta_comentario As String

vblibreta_resta = 0

    On Error Resume Next
    '------------ impresión de detalle ---------
    vcantidad = Format(bLibreta.Recordset("cantidad"), "#####0")
    vDetalle = Format(bLibreta.Recordset("descrip"), "############################################")
    vtotalnuevo = 0
    
    vprecio_modificado = "  "
    vadd_resta = 0
    
    If Not bLibreta.Recordset("pespecial") = True Then
        
        vprecio_buscado = Busca_precio(vnlista, bLibreta.Recordset("fdetalle.codigo"))
        
        vtotal = vprecio_buscado * vcantidad
        
        If Not Val(Format(vtotal, "#######0.00")) = Val(Format(bLibreta.Recordset("total"), "#######0.00")) Then
            vtotalnuevo = vtotalnuevo + vtotal
                       
            
            
            If bLibreta.Recordset("precio") = 0 Then
                vtotal = Format(bLibreta.Recordset("total"), "#########0.000")
            Else
                fmodif_fact (bLibreta.Recordset("remito")), (vtotal - bLibreta.Recordset("totaliva"))
                fmodif_fdetalle bLibreta.Recordset("fdetalle.codigo"), bLibreta.Recordset("remito"), vprecio_buscado
            End If
            
            vadd_resta = vdiferencia
            
            vprecio_modificado = " * " ' si el precio es modificado, aparece un "*" al lado del nombre del artículo
            
        End If
    Else
        vprecio_modificado = "  "
        vtotal = Format(bLibreta.Recordset("total"), "#########0.000")
    End If

    
    ' ---- si tiene un pago parcial, tiene que decirle en la leyenda lo que resta
    If bLibreta.Recordset("pagado") = "PARCIAL" Or ((bLibreta.Recordset("pagado") = "SI") And vadd_resta > 0) Then
        vresta_comentario = Format(" Pago:" + Format(Format(bLibreta.Recordset("pago"), "##0.00"), "@@@@@@@@@@"), "@@@@@@@@@@@@@@@@@")
    Else
        vresta_comentario = Format("           ", "@@@@@@@@@@@@@@@@@")
    End If
    
    If bLibreta.Recordset("pagado") = "SI" And vadd_resta < 0 Then
        display.AddItem (bLibreta.Recordset("nombre") + "  " + bLibreta.Recordset("fdetalle.codigo") + " Diferencia: " + Str(vadd_resta))
    End If
    
    
    'vpago = Format(bLibreta.Recordset("pago"), "#########0.000") PRIMER REFORMA
    If vblibreta_resta > 0 Then
        vtresta = vtresta + vblibreta_resta
    Else
        vtresta = vtresta + (bLibreta.Recordset("resta"))
    End If
    
    'Printer.Print Str(bLibreta.Recordset("fecha")) & " " & num_i2(vcantidad) & "  " & char_i(vdetalle, 28) & num_i(vtotal) & vresta_comentario & num_i(Format(vtresta + vadd_resta, "#########0.00"))
    Printer.Print Str(bLibreta.Recordset("fecha")) & " " & num_i2(vcantidad) & "  " & char_i(vDetalle, 28) & num_i(vtotal) & vresta_comentario & num_i(Format(vtresta, "#########0.00"))

    'Original lleva el pago del art. PRIMER REFORMA
    'printer.print Str(bLibreta.Recordset("fecha")) & " " & num_i2(vcantidad) & "      " & char_i(vdetalle, 28) & " " & num_i(vtotal) & "" & num_i(vpago) & " " & num_i(Format(vtresta, "#########0.00"))
    '-------------------------------------------
    fi_detalle = vtotalnuevo
    If Err Then GrabarLog "fi_detalle", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub fi_detalle_con_clientes_de_ultimo_pago()
    On Error Resume Next
    '------------ impresión de detalle ---------
    vcantidad = Format(bLibreta.Recordset("cantidad"), "#####0")
    vDetalle = Format(bLibreta.Recordset("descrip"), "############################################")
    
    'vtotal = Format(bLibreta.Recordset("total"), "#########0.000")
    btemp_quebranto.Recordset.Filter = ("Codigo = '" + Trim(bLibreta.Recordset("CodCli")) + "'")
    If Not btemp_quebranto.Recordset.EOF = True Then
        If Not bLibreta.Recordset("pespecial") = True Then
            vtotal = Val(Format(bLibreta.Recordset("total") + ((Val(Format(bLibreta.Recordset("total"), "#########0.000")) * Val(Format(bLibreta.Recordset("porcentaje"), "#########0.000"))) / 100)))
            'fmodif_fact (bLibreta.Recordset("CodCli")), (bLibreta.Recordset("Fecha")), (vtotal)
            fmodif_fact (bLibreta.Recordset("Remito")), (vtotal)
        Else
            vtotal = Format(bLibreta.Recordset("total"), "#########0.000")
        End If
    Else
        vtotal = Format(bLibreta.Recordset("total"), "#########0.000")
    End If

    'vpago = Format(bLibreta.Recordset("pago"), "#########0.000") PRIMER REFORMA
    vtresta = vtresta + bLibreta.Recordset("resta")
    
    Printer.Print Str(bLibreta.Recordset("fecha")) & " " & num_i2(vcantidad) & "      " & char_i(vDetalle, 28) & " " & num_i(vtotal) & "         " & num_i(Format(vtresta, "#########0.00"))

    'Original lleva el pago del art. PRIMER REFORMA
    'printer.print Str(bLibreta.Recordset("fecha")) & " " & num_i2(vcantidad) & "      " & char_i(vdetalle, 28) & " " & num_i(vtotal) & "" & num_i(vpago) & " " & num_i(Format(vtresta, "#########0.00"))
    '-------------------------------------------
    If Err Then GrabarLog "fi_detalle", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub fititulos(vcodigocli, _
                      vnombrecli, _
                      vdireccioncli As String, _
                      vsaldoante As Double)
    On Error Resume Next
    ' ------- titulo de las impresiones ---------
    Printer.Print "--- Sodería LA SURGENTE  -  ESTADO DE CUENTAS               -  " + strfecha2(fhasta.Value) + ""
    Printer.Print "> Cliente     : " & Trim(char_i(vcodigocli, 6)) & "-" & char_i(vnombrecli, 30) + ""
    Printer.Print "> Dirección   : " & char_i(vdireccioncli, 30)
    Printer.Print "--------------------------------------------------------------------------------"
    'printer.print "FECHA      CANTIDAD        DETALLE                    TOTAL      PAGO      SALDO"
    Printer.Print "FECHA      CANTIDAD        DETALLE                    TOTAL                SALDO"
    Printer.Print "--------------------------------------------------------------------------------"

    If Err Then GrabarLog "fititulos", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub flibreta()
    On Error Resume Next
        ' imprime los listado para adjuntar a la libreta
    With bReparto_Repartidor
    
    Dim sql As String
    sql = ""
            
            If Not vcdesde.Text = "" Then sql = " and nombre >='" + Trim(vcdesde) + "'"
            If Not vchasta.Text = "" Then sql = " and nombre <='" + Trim(vchasta) + "'"
            
    
            'Selecciono todas las personas que van a salir impresas, de la consulta Reparto_repartidor
        If Val(vcdesde) > 0 Then
        
        .RecordSource = "select * from reparto_repartidor where codigo = '" + Trim(vcdesde) + "'"
        .Refresh
        
        Else
          
            .RecordSource = "select * from reparto_repartidor where cod_reparto = '" + Trim(vcod_reparto) + "'" + sql + " order by nombre"
            .Refresh
        
        End If
    
        If .Recordset.EOF = True Then
            MsgBox "No existen datos para imprimir", vbInformation, "Error ..."
            Exit Sub
        End If
    
        lado = 2
        barra.Value = 0
        barra.Max = .Recordset.RecordCount
    
        Do Until .Recordset.EOF = True
            'If (.Recordset("codigo") = Trim(vcodigodesdei)) Or (vbdesde = True) Then
                
                'If .Recordset("codigo") = "15003" Then
                ' MsgBox "gsgfds"
                'End If
                fi_cliente (.Recordset("codigo")) ' ----- imprime un cliente en la libreta
                vbdesde = True
            'End If
            .Recordset.MoveNext
            barra.Value = barra.Value + 1
        Loop
        Printer.EndDoc
    End With
    If Err Then GrabarLog "flibreta", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub fmodif_ctacte(vremito As Long, vtotalmodif As Double)
On Error Resume Next
    With bccliente
        .RecordSource = "Select * from cuentascorrientes where remito = " & Trim(vremito)
        .Refresh
        If Not .Recordset.RecordCount = 0 Then .Recordset.MoveFirst
        .Recordset("debito") = .Recordset("debito") + vtotalmodif
        .Recordset.Update
    End With
If Err Then GrabarLog "Fmodif_ctacte", Err.Number & " " & Err.Description, Me.Name
End Sub
'------------------------------------------------------------------------------------
 
Function fmodif_fact(vremito As Long, vtotalmodif As Double) As Long
On Error Resume Next
    With bfactura
        .RecordSource = "Select * from factura where remito = " & Trim(vremito) & " order by codigo"
        .Refresh
        If .Recordset.RecordCount = 0 Then
            Exit Function
        End If
        .Recordset.MoveFirst
        .Recordset("subtotal") = .Recordset("subtotal") + vtotalmodif
        .Recordset("total") = Val(Format(.Recordset("subtotal"), "#########0.00")) + Val(Format(.Recordset("Tiva"), "#########0.00"))
        .Recordset("total_ctacte") = Val(Format(.Recordset("subtotal"), "#########0.00")) + Val(Format(.Recordset("Tiva"), "#########0.00"))
        .Recordset.Update
        fmodif_ctacte (.Recordset("remito")), (vtotalmodif)
    End With
If Err Then GrabarLog "Fmodif_fact", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub fmodif_fdetalle(vcodigoart As String, vremito As Long, vprecio As Double)
Dim vttotal, vttotal_viejo, vprecio_viejo As Double
If vprecio = 0 Then Exit Sub

    With bfdetalle
        .Refresh
        .Recordset.Filter = ("codigo = '" + vcodigoart + "' and remito = " & vremito & "")
        If Not .Recordset.EOF = True Then
            
            vprecio_viejo = .Recordset("precio")
            
            
            .Recordset("precio") = vprecio
            
            vttotal = vprecio * .Recordset("cantidad")
            vttotal_viejo = vprecio_viejo * .Recordset("cantidad")
            
            vdiferencia = vttotal - vttotal_viejo
                       
            
            
            .Recordset("totaliva") = vttotal
            .Recordset("total") = vttotal
            
            If .Recordset("total_ctacte") > 0 Then
                .Recordset("total_ctacte") = vttotal
            End If
                        
            .Recordset("resta") = vttotal - .Recordset("pago")
            vblibreta_resta = vttotal - .Recordset("pago")
                        
            If .Recordset("resta") = 0 Then .Recordset("pagado") = "SI"
            If .Recordset("resta") > 0 Then .Recordset("pagado") = "NO"
            If .Recordset("resta") < 0 Then .Recordset("pagado") = "SOBRANTE"
            
            .Recordset.Update
        Else
            Exit Sub
        End If
    End With
End Sub

Private Sub Form_Load()
    On Error Resume Next

    With bReparto_Repartidor
        .ConnectionString = pathDBMySQL
        .RecordSource = "Reparto_Repartidor"
        .Refresh
    End With
        
    With bpagos_libretas
        .ConnectionString = pathDBMySQL
        .RecordSource = "pagos_libretas"
        .Refresh
    End With
    
    With barticulos
        .ConnectionString = pathDBMySQL
        .RecordSource = "articulos"
        .Refresh
    End With

    With bLibreta
        .ConnectionString = pathDBMySQL
        .RecordSource = "Libreta"
        .Refresh
    End With

    With breparto_agrupado
        .ConnectionString = pathDBMySQL
        .RecordSource = "Reparto_Agrupado"
        .Refresh
    End With

    With bccliente
        .ConnectionString = pathDBMySQL
        .RecordSource = "CuentasCorrientes"
        .Refresh
    End With

    With bcliente
        .ConnectionString = pathDBMySQL
        .RecordSource = "Clientes"
        .Refresh
    End With

    With bsaldos_clientes
        .ConnectionString = pathDBMySQL
        .RecordSource = "saldos3"
        .Refresh
    End With

    With blistas
        .ConnectionString = pathDBMySQL
        .RecordSource = "Listas"
        .Refresh
    End With

    With bfactura
        .ConnectionString = pathDBMySQL
        .RecordSource = "Factura"
        .Refresh
    End With

    With bfdetalle
        .ConnectionString = pathDBMySQL
        .RecordSource = "Fdetalle"
        .Refresh
    End With
 
    With btemp_quebranto
        .ConnectionString = pathDBMySQL
        .RecordSource = "Temp_quebranto"
        .Refresh
    End With

    Me.Top = 1000
    Me.Left = 1300
    Me.Width = 5300 '11000 '5280
    Me.Height = 7500 '4650
    fdesde.Value = Date
    fhasta.Value = Date

    If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub idetallectacte(fdesde As Date, _
                           fhasta As Date)
    Dim condetalle As Integer
    On Error Resume Next

    If Not (vcodigodesde = "" Or vcodigohasta = "") Then
        bcliente.RecordSource = "select * from clientes where codigo >= '" + vcodigodesde + "' and codigo <= '" + vcodigohasta + "'"
    End If

    bcliente.Refresh

    condetalle = 0

    If Me.tlistado = "Ficha c/ detalle" Then
        condetalle = 1
    Else
        condetalle = 0
    End If

    Do Until bcliente.Recordset.EOF
        Mantenimiento.rsccc.Filter = "codigo = '" + bcliente.Recordset("codigo") + "' and importe = 0 and  fecha <= '" & strfechaMySQL(fhasta) + "' and fecha >= '" & strfechaMySQL(fdesde) & "'"
        Mantenimiento.rsccc.Sort = "fecha"
        Mantenimiento.rsccc.Sort = "id"
    
        Mantenimiento.rscctacte_detalle.Filter = "codigo = '" + bcliente.Recordset("codigo") + "' and importe = 0 and  fecha <= '" & strfechaMySQL(fhasta) & "' and fecha >= '" & strfechaMySQL(fdesde) & "'"
        Mantenimiento.rscctacte_detalle.Sort = "fecha"
        Mantenimiento.rscctacte_detalle.Sort = "id"
    
        If condetalle = 1 Then
            condetalle = 1
            drcuentascorrientes_detalles.Visible = False
            drcuentascorrientes_detalles.PrintReport False
        Else
        
            drcuentascorrientes.Sections("TituloEmpresa").Controls("vcliente").Caption = bcliente.Recordset("codigo") + " " + bcliente.Recordset("nombre")
            'drcuentascorrientes.Sections("TituloEmpresa").Controls("vsaldo").Caption = saldo.Caption
        
            drcuentascorrientes.Visible = False
            drcuentascorrientes.PrintReport False
        End If
        
        Mantenimiento.rsccc.Close
    
        bcliente.Recordset.MoveNext
    Loop

    If Err Then GrabarLog "idetallectacte", Err.Number & " " & Err.Description, Me.Name
End Sub
'----------------------------------------------------------------------------------
'--------------------------[ Funciones de Impresión]-------------------------------
'----------------------------------------------------------------------------------

'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------
'----------------------------------------------------------------------------------

Private Sub Limpiar()
    On Error Resume Next
    vcdesde.Text = ""
    vchasta.Text = ""
    vlocalidad.Text = ""
    vcodigodesde = ""
    vcodigohasta = ""
    cbolista.Text = ""
    vreparto.Text = ""

    If Err Then GrabarLog "limpiar", Err.Number & " " & Err.Description, Me.Name
End Sub

Public Sub vcdesde_KeyPress(KeyAscii As Integer)
    On Error Resume Next

    If KeyAscii = 13 Then
        buscacli vcdesde, "d"
    End If

    If Err Then GrabarLog "vcdesde_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub

Public Sub vchasta_KeyPress(KeyAscii As Integer)
    On Error Resume Next

    If KeyAscii = 13 Then
        buscacli vchasta, "h"
    End If

    If Err Then GrabarLog "vchasta_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub vclidesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        buscacli vclidesde, "j"
    End If
End Sub

Private Sub vlocalidad_GotFocus()
    On Error Resume Next

    With bcliente
        .RecordSource = "Select distinct localidad from clientes order by localidad"
        .Refresh

        If Not .Recordset.RecordCount = 0 Then .Recordset.MoveFirst
        vlocalidad.Clear

        Do Until .Recordset.EOF
            vlocalidad.AddItem Trim(.Recordset("Localidad"))
            .Recordset.MoveNext
        Loop

    End With

    If Err Then GrabarLog "vchasta_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub

Public Sub vlocalidad_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Command7.SetFocus
    End If

End Sub

Private Sub vreparto_GotFocus()
    On Error Resume Next

    With breparto_agrupado
        .Refresh

        If Not .Recordset.RecordCount = 0 Then .Recordset.MoveFirst
        vreparto.Clear

        Do Until .Recordset.EOF
            vreparto.AddItem Trim(.Recordset("reparto"))
            .Recordset.MoveNext
        Loop

    End With

    If Err Then GrabarLog "vreparto_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub vreparto_LostFocus()
    On Error Resume Next

    With breparto_agrupado
        .Recordset.MoveFirst
        .Recordset.Find ("reparto = '" + Trim(vreparto.Text) + "'")
        vcod_reparto = .Recordset("cod_reparto")
    End With

    If Err Then
        MsgBox "Imposible asignar reparto a este listado. Verifique los datos ingresados", vbCritical, "Error.."
        GrabarLog "vreparto_LostFocus", Err.Number & " " & Err.Description, Me.Name
        Exit Sub
    End If
    
End Sub

