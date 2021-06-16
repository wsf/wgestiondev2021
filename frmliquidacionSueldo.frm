VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmLiquidacionSueldos 
   Caption         =   "Liquidación de Sueldos por Cliente y por Rubro"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   14850
   Begin VB.CommandButton cmdImprimirUltima 
      Caption         =   "Imprimir última Liquidación (Período anterior acumulado / Actual)"
      Height          =   375
      Left            =   8370
      TabIndex        =   31
      Top             =   7560
      Width           =   4935
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "Borrar"
      Height          =   315
      Left            =   3390
      TabIndex        =   30
      Top             =   4650
      Width           =   1665
   End
   Begin MSAdodcLib.Adodc bliqui_sueldos_final 
      Height          =   375
      Left            =   6090
      Top             =   3360
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Vbprog\WGestion (La Surgente)\Datos\Wgestion.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Vbprog\WGestion (La Surgente)\Datos\Wgestion.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "liqui_sueldo_final"
      Caption         =   "bliqui_sueldos"
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
   Begin VB.Frame Frame1 
      Caption         =   "Detalles de la liquidación :"
      Height          =   765
      Left            =   4320
      TabIndex        =   24
      Top             =   1560
      Width           =   4425
      Begin VB.CommandButton cmdSDetalle 
         Caption         =   "S/ Detalle"
         Height          =   375
         Left            =   3030
         TabIndex        =   26
         Top             =   270
         Width           =   1335
      End
      Begin VB.CommandButton cmdCDetalle_Defecto 
         Caption         =   "C/Detalle - Defecto"
         Height          =   375
         Left            =   1410
         TabIndex        =   25
         Top             =   270
         Width           =   1605
      End
      Begin VB.CommandButton cmdCDetalle 
         Caption         =   "C/ Detalle"
         Enabled         =   0   'False
         Height          =   375
         Left            =   150
         TabIndex        =   27
         Top             =   270
         Width           =   1245
      End
   End
   Begin MSAdodcLib.Adodc btemp 
      Height          =   375
      Left            =   11460
      Top             =   1500
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
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
      Caption         =   "btemp"
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
   Begin MSAdodcLib.Adodc bresumen_liqui 
      Height          =   375
      Left            =   11430
      Top             =   1170
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
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
      Caption         =   "bresumen_liqui"
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
   Begin MSAdodcLib.Adodc bliqui_temp 
      Height          =   375
      Left            =   11400
      Top             =   810
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
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
      Caption         =   "bliqui_temp"
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
   Begin MSAdodcLib.Adodc bliqui_temp_final 
      Height          =   375
      Left            =   11430
      Top             =   420
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
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
      Caption         =   "bliqui_temp_final"
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
   Begin VB.CommandButton cmdLiquiActual 
      Caption         =   ">> Liquidar período actual"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2190
      TabIndex        =   23
      Top             =   1920
      Width           =   2025
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Agregar"
      Height          =   315
      Left            =   1740
      TabIndex        =   20
      Top             =   4650
      Width           =   1665
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Actualizar"
      Height          =   315
      Left            =   90
      TabIndex        =   19
      Top             =   4650
      Width           =   1665
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "Limpiar este resumen:"
      Height          =   285
      Left            =   8940
      TabIndex        =   18
      Top             =   2310
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Liquidar Sueldo"
      Height          =   375
      Left            =   13320
      TabIndex        =   17
      Top             =   7560
      Width           =   1395
   End
   Begin VB.CommandButton cmdLiquiAnterior 
      Caption         =   "Liquidar período anterior >>"
      Height          =   375
      Left            =   60
      TabIndex        =   14
      Top             =   1920
      Width           =   2145
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1995
      Left            =   90
      TabIndex        =   13
      Top             =   5490
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   3519
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
            LCID            =   1034
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
            LCID            =   1034
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
   Begin VB.ListBox liqui 
      BackColor       =   &H00EBE4E7&
      Height          =   2010
      Left            =   8940
      TabIndex        =   11
      Top             =   270
      Width           =   5715
   End
   Begin MSComctlLib.ProgressBar barra 
      Height          =   225
      Left            =   60
      TabIndex        =   9
      Top             =   1290
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSAdodcLib.Adodc bliqui_sueldos 
      Height          =   375
      Left            =   3120
      Top             =   3390
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Vbprog\WGestion (La Surgente)\Datos\Wgestion.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Vbprog\WGestion (La Surgente)\Datos\Wgestion.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "liqui_sueldos"
      Caption         =   "bliqui_sueldos"
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
   Begin MSAdodcLib.Adodc bempleados 
      Height          =   375
      Left            =   240
      Top             =   3360
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
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
      Caption         =   "bempleados"
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
      Height          =   375
      Left            =   240
      Top             =   3720
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
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
   Begin MSAdodcLib.Adodc bemp 
      Height          =   375
      Left            =   3120
      Top             =   3780
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
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
      Caption         =   "bemp"
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
   Begin VB.Frame Frame2 
      Height          =   1275
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   8715
      Begin VB.TextBox vrepartidor 
         Height          =   285
         Left            =   1170
         TabIndex        =   1
         Top             =   180
         Width           =   6765
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   285
         Left            =   4230
         TabIndex        =   4
         Top             =   510
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         Format          =   76742657
         CurrentDate     =   38023
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   285
         Left            =   6300
         TabIndex        =   5
         Top             =   540
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
         _Version        =   393216
         Format          =   76742657
         CurrentDate     =   38023
      End
      Begin MSComCtl2.DTPicker frevisar 
         Height          =   285
         Left            =   3450
         TabIndex        =   29
         Top             =   870
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   503
         _Version        =   393216
         Format          =   76742657
         CurrentDate     =   38023
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "> Revisar cobros desde:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   180
         TabIndex        =   28
         Top             =   930
         Width           =   2595
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "> Período que desea liquidar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   150
         TabIndex        =   8
         Top             =   540
         Width           =   2595
      End
      Begin VB.Label Label3 
         Caption         =   "Desde :"
         Height          =   225
         Left            =   3510
         TabIndex        =   7
         Top             =   570
         Width           =   705
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta :"
         Height          =   225
         Left            =   5670
         TabIndex        =   6
         Top             =   570
         Width           =   615
      End
      Begin VB.Label Label2 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   810
         Width           =   1905
      End
      Begin VB.Label Label1 
         Caption         =   "> Empleado :"
         Height          =   285
         Left            =   180
         TabIndex        =   2
         Top             =   210
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1995
      Left            =   60
      TabIndex        =   16
      Top             =   2640
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   3519
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   14410208
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
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
   Begin MSComCtl2.DTPicker vfecha 
      Height          =   285
      Left            =   2010
      TabIndex        =   22
      Top             =   1560
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   503
      _Version        =   393216
      Format          =   76742657
      CurrentDate     =   38023
   End
   Begin VB.Label Label7 
      Caption         =   "> Fecha de la liquidación:"
      Height          =   255
      Left            =   90
      TabIndex        =   21
      Top             =   1590
      Width           =   1905
   End
   Begin VB.Label Label6 
      Caption         =   "> Pagos por liquidación:"
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   2
      Left            =   60
      TabIndex        =   15
      Top             =   2370
      Width           =   14655
   End
   Begin VB.Label Label6 
      Caption         =   "> Resultado de la liquidación del período seleccionado:"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   5220
      Width           =   14625
   End
   Begin VB.Label Label6 
      Caption         =   "> Resumen de liquidaciones anteriores:"
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   8970
      TabIndex        =   10
      Top             =   30
      Width           =   5685
   End
End
Attribute VB_Name = "frmLiquidacionSueldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vtotal_a_cobrar As Double


Dim vcod_repartidor As String
Dim vtresta, vgtresta, vgtpago, vgtsueldo, vajuste, vgtotal As Double
Private Sub actualizacion_liquidacion()
On Error Resume Next
Dim vcod_art, vcodigo, varticulo As String


Dim vttotal, vtcdo, vtctacte, vtganancia, vtcantidad, vtcobrado, vcobrado, vporcentaje As Double
Dim vpor_art, vttganancia As Double

vcod_art = ""
vpor_art = 0

bliqui_temp.Refresh
BorrarBase "liqui_temp", pathDBMySQL
bliqui_temp.Refresh


bliqui_sueldos.RecordSource = "SELECT * FROM liqui_sueldo_defecto WHERE (fecha >= '" & strfechaMySQL(frevisar) + "' and fecha < '" & strfechaMySQL(fdesde) + "') and (repartidor = '" + Trim(vcod_repartidor) + "') order by articulos.codigo"
bliqui_sueldos.Refresh

barra.Max = bliqui_sueldos.Recordset.RecordCount
liqui.AddItem ("Cantidad de registros para procesar: " + Str(bliqui_sueldos.Recordset.RecordCount))
barra.Value = 0
vtganancia = 0

vcod_art = bliqui_sueldos.Recordset("articulos.codigo")
vpor_art = bliqui_sueldos.Recordset("ganancia")

vttganancia = 0

Do Until bliqui_sueldos.Recordset.EOF = True
    
    If Not (vcod_art = bliqui_sueldos.Recordset("articulos.codigo")) Or Not (vpor_art = bliqui_sueldos.Recordset("ganancia")) Then
        '------------ grabo en el temporal ------------------
            bliqui_temp.Recordset.AddNew
            
            bliqui_temp.Recordset("codigo") = vcodigo
            bliqui_temp.Recordset("Articulo") = varticulo
            bliqui_temp.Recordset("cantidad") = vtcantidad
            bliqui_temp.Recordset("total") = vttotal
            bliqui_temp.Recordset("porcentaje") = vporcentaje
            bliqui_temp.Recordset("cdo") = vtcdo
            bliqui_temp.Recordset("ctacte") = vtctacte
            
            vtganancia = vporcentaje * vtcobrado / 100
            
            bliqui_temp.Recordset("ganancia") = vtganancia
            
            vttganancia = vttganancia + vtganancia '  acá se acumula el sueldo general del periodo desde/hasta
            
            bliqui_temp.Recordset("cobrado") = vtcobrado
            
            bliqui_temp.Recordset.Update
        '-----------------------------------------------------
        
            vtcantidad = 0
            vttotal = 0
            
            vtcdo = 0
            vtctacte = 0
            
            vtganancia = 0
            vtcobrado = 0
            
            vporcentaje = 0
            
            vcod_art = bliqui_sueldos.Recordset("articulos.codigo")
            vpor_art = bliqui_sueldos.Recordset("ganancia")
        
    End If
    
    

    
    barra.Value = barra.Value + 1
    
    vcodigo = bliqui_sueldos.Recordset("articulos.codigo")
    varticulo = bliqui_sueldos.Recordset("ÚltimoDeDescrip")
   ' vporcentaje = bliqui_sueldos.Recordset("ÚltimoDeporcentaje")
    
    vtcobrado = vtcobrado + bliqui_sueldos.Recordset("SumaDepago1")
    vtcantidad = vtcantidad + bliqui_sueldos.Recordset("sumadecantidad")
    vttotal = vttotal + bliqui_sueldos.Recordset("sumadetotal")
    vtcdo = vtcdo + bliqui_sueldos.Recordset("SumaDetotal_cdo")
    vtctacte = vtctacte + bliqui_sueldos.Recordset("SumaDetotal_ctacte")
    
    '----- calculo cual es la ganancia del repartidor por artículo -------
    vtganancia = vtganancia + bliqui_sueldos.Recordset("SumaDeSueldo")
        
    
    vporcentaje = bliqui_sueldos.Recordset("Ganancia")
    
    bliqui_sueldos.Recordset.MoveNext
Loop

 ' ---------------- grabo el ultimo en el temporal ----------------------------------
     '------------ grabo en el temporal ------------------
            bliqui_temp.Recordset.AddNew
            
            bliqui_temp.Recordset("codigo") = vcodigo
            bliqui_temp.Recordset("Articulo") = varticulo
            bliqui_temp.Recordset("cantidad") = vtcantidad
            bliqui_temp.Recordset("total") = vttotal
            bliqui_temp.Recordset("porcentaje") = vporcentaje
            bliqui_temp.Recordset("cdo") = vtcdo
            bliqui_temp.Recordset("ctacte") = vtctacte
            
            vtganancia = vporcentaje * vtcobrado / 100
            
            bliqui_temp.Recordset("ganancia") = vtganancia
            
            vttganancia = vttganancia + vtganancia '  acá se acumula el sueldo general del periodo desde/hasta
            
            bliqui_temp.Recordset("cobrado") = vtcobrado
            
            bliqui_temp.Recordset.Update
        '-----------------------------------------------------
    

' --- anotación de los resultados ----------------------------------------
liqui.AddItem ("-------------------------------------")
liqui.AddItem ("> Valores globales:")
liqui.AddItem ("  Ganancia acumulada desde el principio: " + Format(vttganancia, "########0.00"))
liqui.AddItem ("  Pago acumulado desde el principio    : " + Format(vgtpago, "########0.00"))
 

vajuste = vttganancia - vgtpago
liqui.AddItem ("  Importe a pagar por ajuste de los meses anteriores : " + Format(vajuste, "#########0.00"))
' -------------------------------------------------------------------------------------------------------------


'------- guardo los datos en la tabla resumen_liqui -----
bresumen_liqui.Refresh
bresumen_liqui.Recordset.AddNew
bresumen_liqui.Recordset("fecha") = vfecha
bresumen_liqui.Recordset("repartidor") = vcod_repartidor
bresumen_liqui.Recordset("nombre") = Me.vrepartidor
bresumen_liqui.Recordset("hasta") = Me.fhasta
bresumen_liqui.Recordset("pago") = Val(Format(vajuste, "#######0.00"))
bresumen_liqui.Recordset("comentario") = " Importe por cobro de meses anteriores"

bresumen_liqui.Recordset.Update


End Sub

Private Sub cmdAgregar_Click()
On Error Resume Next

    With bresumen_liqui
        .Refresh
        .Recordset.AddNew

        .Recordset("fecha") = vfecha
        .Recordset("repartidor") = vcod_repartidor
        .Recordset("nombre") = Me.vrepartidor
        
        .Recordset.Update
    End With

If Err Then GrabarLog "cmdAgregar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdBorrar_Click()
On Error Resume Next
    
    If InputBox("Ingrese la palabra 'borrar', para eliminar el registro.", "Mensaje ...") = "borrar" Then
        Borrar bresumen_liqui, False
    End If

If Err Then GrabarLog "cmdBorrar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdCDetalle_Click()
On Error Resume Next
Dim vcod_art, vcodigo, varticulo As String


Dim vttotal, vtcdo, vtctacte, vtganancia, vtcantidad, vtcobrado, vcobrado, vporcentaje As Double

vcod_art = ""

bliqui_temp.Refresh
BorrarBase "liqui_temp", pathDBMySQL
bliqui_temp.Refresh


bliqui_sueldos.RecordSource = "select * from liqui_sueldos where fecha >= '" & strfechaMySQL(fdesde) + "' and fecha <= '" & strfechaMySQL(fhasta) + "' and cod_repartidor = '" + Trim(vcod_repartidor) + "' order by codigo"
bliqui_sueldos.Refresh

barra.Max = bliqui_sueldos.Recordset.RecordCount

barra.Value = 0

vcod_art = bliqui_sueldos.Recordset("codigo")

Do Until bliqui_sueldos.Recordset.EOF
    
    If Not vcod_art = bliqui_sueldos.Recordset("codigo") Then
        '------------ grabo en el temporal ------------------
            bliqui_temp.Recordset.AddNew
            
            bliqui_temp.Recordset("codigo") = vcodigo
            bliqui_temp.Recordset("Articulo") = varticulo
            bliqui_temp.Recordset("cantidad") = vtcantidad
            bliqui_temp.Recordset("total") = vttotal
            bliqui_temp.Recordset("porcentaje") = vporcentaje
            bliqui_temp.Recordset("cdo") = vtcdo
            bliqui_temp.Recordset("ctacte") = vtctacte
            bliqui_temp.Recordset("ganancia") = vtganancia
            bliqui_temp.Recordset("cobrado") = vtcobrado
            
            bliqui_temp.Recordset.Update
        '-----------------------------------------------------
        
            vtcantidad = 0
            vttotal = 0
            vtcdo = 0
            vtctacte = 0
            vtganancia = 0
            vtcobrado = 0
            
            vcod_art = bliqui_sueldos.Recordset("codigo")
        
    End If
    
    barra.Value = barra.Value + 1
    
    vcodigo = bliqui_sueldos.Recordset("codigo")
    varticulo = bliqui_sueldos.Recordset("ÚltimoDeDescrip")
    vporcentaje = bliqui_sueldos.Recordset("ÚltimoDeporcentaje")
    
    vtcobrado = vtcobrado + bliqui_sueldos.Recordset("SumaDepago1")
    vtcantidad = vtcantidad + bliqui_sueldos.Recordset("sumadecantidad")
    vttotal = vttotal + bliqui_sueldos.Recordset("sumadetotal")
    vtcdo = vtcdo + bliqui_sueldos.Recordset("SumaDetotal_cdo")
    vtctacte = vtctacte + bliqui_sueldos.Recordset("SumaDetotal_ctacte")
    vtganancia = vtganancia + bliqui_sueldos.Recordset("ganancia")
    
    bliqui_sueldos.Recordset.MoveNext
Loop

     If Mantenimiento.rsLiqui.State = 1 Then
        Mantenimiento.rsLiqui.Close
        Mantenimiento.rsLiqui.Open
    Else
        Mantenimiento.rsLiqui.Open
        Mantenimiento.rsLiqui.Close
        Mantenimiento.rsLiqui.Open
    End If
    
    
    drliqui_sueldos.Sections("TituloEmpresa").Controls("vrepartidor").Caption = Me.vrepartidor
    drliqui_sueldos.Sections("TituloEmpresa").Controls("vfdesde").Caption = Str(fdesde.Value)
    drliqui_sueldos.Sections("TituloEmpresa").Controls("vfhasta").Caption = Str(fhasta.Value)
    
    drliqui_sueldos.Sections("TituloEmpresa").Controls("vrepartidor").Caption = Me.vrepartidor
    drliqui_sueldos.Show
    
If Err Then GrabarLog "cmdCDetalle_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdCDetalle_Defecto_Click()
On Error Resume Next
Dim vcod_art, vcodigo, varticulo As String


Dim vttotal, vtcdo, vtctacte, vtganancia, vtcantidad, vtcobrado, vcobrado, vporcentaje As Double
Dim vpor_art As Double

vcod_art = ""
vpor_art = 0

bliqui_temp.Refresh
BorrarBase "liqui_temp", pathDBMySQL
bliqui_temp.Refresh


bliqui_sueldos.RecordSource = "select * from liqui_sueldo_defecto where fecha >= '" & strfechaMySQL(fdesde) + "' and fecha <= '" & strfechaMySQL(fhasta) + "' and repartidor = '" + Trim(vcod_repartidor) + "' order by articulos.codigo"
bliqui_sueldos.Refresh

barra.Max = bliqui_sueldos.Recordset.RecordCount

barra.Value = 0

vcod_art = bliqui_sueldos.Recordset("articulos.codigo")
vpor_art = bliqui_sueldos.Recordset("ganancia")

Do Until bliqui_sueldos.Recordset.EOF = True
    
    If Not (vcod_art = bliqui_sueldos.Recordset("articulos.codigo")) Or Not (vpor_art = bliqui_sueldos.Recordset("ganancia")) Then
        '------------ grabo en el temporal ------------------
            bliqui_temp.Recordset.AddNew
            
            bliqui_temp.Recordset("codigo") = vcodigo
            bliqui_temp.Recordset("Articulo") = varticulo
            bliqui_temp.Recordset("cantidad") = vtcantidad
            bliqui_temp.Recordset("total") = vttotal
            bliqui_temp.Recordset("porcentaje") = vporcentaje
            bliqui_temp.Recordset("cdo") = vtcdo
            bliqui_temp.Recordset("ctacte") = vtctacte
            bliqui_temp.Recordset("ganancia") = vtganancia
            bliqui_temp.Recordset("cobrado") = vtcobrado
            
            bliqui_temp.Recordset.Update
        '-----------------------------------------------------
        
            vtcantidad = 0
            vttotal = 0
            
            vtcdo = 0
            vtctacte = 0
            
            vtganancia = 0
            vtcobrado = 0
            
            vporcentaje = 0
            
            vcod_art = bliqui_sueldos.Recordset("articulos.codigo")
            vpor_art = bliqui_sueldos.Recordset("ganancia")
        
    End If
    
    barra.Value = barra.Value + 1
    
    vcodigo = bliqui_sueldos.Recordset("articulos.codigo")
    varticulo = bliqui_sueldos.Recordset("ÚltimoDeDescrip")
   ' vporcentaje = bliqui_sueldos.Recordset("ÚltimoDeporcentaje")
    
    vtcobrado = vtcobrado + bliqui_sueldos.Recordset("SumaDepago1")
    vtcantidad = vtcantidad + bliqui_sueldos.Recordset("sumadecantidad")
    vttotal = vttotal + bliqui_sueldos.Recordset("sumadetotal")
    vtcdo = vtcdo + bliqui_sueldos.Recordset("SumaDetotal_cdo")
    vtctacte = vtctacte + bliqui_sueldos.Recordset("SumaDetotal_ctacte")
    vtganancia = vtganancia + bliqui_sueldos.Recordset("SumaDeSueldo")
    
    vporcentaje = bliqui_sueldos.Recordset("Ganancia")
    
    bliqui_sueldos.Recordset.MoveNext
Loop

     If Mantenimiento.rsLiqui.State = 1 Then
        Mantenimiento.rsLiqui.Close
        Mantenimiento.rsLiqui.Open
    Else
        Mantenimiento.rsLiqui.Open
        Mantenimiento.rsLiqui.Close
        Mantenimiento.rsLiqui.Open
    End If
    
    
    drliqui_sueldos.Sections("TituloEmpresa").Controls("vrepartidor").Caption = Me.vrepartidor
    drliqui_sueldos.Sections("TituloEmpresa").Controls("vfdesde").Caption = Str(fdesde.Value)
    drliqui_sueldos.Sections("TituloEmpresa").Controls("vfhasta").Caption = Str(fhasta.Value)
    
    drliqui_sueldos.Sections("TituloEmpresa").Controls("vrepartidor").Caption = Me.vrepartidor
    drliqui_sueldos.Show
    
If Err Then GrabarLog "cmdCDetalle_Defecto_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdImprimirUltima_Click()
On Error Resume Next
 
    With Mantenimiento.rsliqui_temp_final
        If Not .State = 1 Then .Open
        .Close
        .Open
    End With
    
    With drLiquidacionSueldosFinal
        
        .Sections("TituloEmpresa").Controls("vrepartidor").Caption = vrepartidor
        .Sections("TituloEmpresa").Controls("vfdesde").Caption = Str(fdesde.Value)
        .Sections("TituloEmpresa").Controls("vfhasta").Caption = Str(fhasta.Value)
        .Sections("TituloEmpresa").Controls("vrepartidor").Caption = vrepartidor
    
        .Sections("sección5").Controls("vanterior").Caption = Format(vajuste, "########0.00")
        .Sections("sección5").Controls("vcobro").Caption = Format(vtotal_a_cobrar, "########0.00")
        .Sections("sección5").Controls("vtotal").Caption = Format((vajuste + vtotal_a_cobrar), "########0.00")
    
        .Show
    End With

If Err Then GrabarLog "cmdImprimirUltima_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdLimpiar_Click()
On Error Resume Next

    liqui.Clear

If Err Then GrabarLog "cmdLimpiar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdLiquiActual_Click()

Dim vcod_art, vcodigo, varticulo As String

Dim vttotal, vtcdo, vtctacte, vtganancia, vtcantidad, vtcobrado, vcobrado, vporcentaje As Double
Dim vpor_art As Double

On Error Resume Next

vcod_art = ""
vpor_art = 0

BorrarBase "liqui_temp", pathDBMySQL
bliqui_temp.Refresh

    With bliqui_sueldos_final
        .RecordSource = "select * from liqui_sueldo_defecto where fecha >= '" & strfechaMySQL(fdesde) + "' and fecha <= '" & strfechaMySQL(fhasta) + "' and repartidor = '" + Trim(vcod_repartidor) + "' order by articulos.codigo"
        .Refresh

        barra.Max = .Recordset.RecordCount
        barra.Value = 0

        vcod_art = .Recordset("articulos.codigo").Value
        vpor_art = .Recordset("ganancia").Value
    

        vtotal_a_cobrar = 0
        .Recordset.MoveFirst
    End With

Dim vtotal_vendido As Double

vtotal_vendido = 0
vtotal_a_cobrar = 0

Do Until bliqui_sueldos_final.Recordset.EOF = True
    
    If Not (vcod_art = bliqui_sueldos_final.Recordset("articulos.codigo").Value) Or Not (vpor_art = bliqui_sueldos_final.Recordset("ganancia").Value) Then
        '------------ grabo en el temporal ------------------
            bliqui_temp.Recordset.AddNew
            
            bliqui_temp.Recordset("codigo").Value = vcodigo
            bliqui_temp.Recordset("Articulo").Value = varticulo
            bliqui_temp.Recordset("cantidad").Value = vtcantidad
            bliqui_temp.Recordset("total").Value = vttotal
            
            
            
            bliqui_temp.Recordset("porcentaje").Value = vporcentaje
            bliqui_temp.Recordset("cdo").Value = vtcdo
            bliqui_temp.Recordset("ctacte").Value = vtctacte
            bliqui_temp.Recordset("ganancia").Value = (vporcentaje * vtcobrado) / 100
            
            vtotal_a_cobrar = vtotal_a_cobrar + bliqui_temp.Recordset("ganancia").Value
            vtotal_vendido = vtotal_vendido + (vporcentaje * vttotal) / 100
            
            bliqui_temp.Recordset("cobrado").Value = vtcobrado
            
            bliqui_temp.Recordset.Update
        '-----------------------------------------------------
        
            vtcantidad = 0
            vttotal = 0
            
            vtcdo = 0
            vtctacte = 0
            
            vtganancia = 0
            vtcobrado = 0
            
            vporcentaje = 0
            
            vcod_art = bliqui_sueldos_final.Recordset("articulos.codigo").Value
            vpor_art = bliqui_sueldos_final.Recordset("ganancia").Value
        
    End If
    
    barra.Value = barra.Value + 1
    
    vcodigo = bliqui_sueldos_final.Recordset("articulos.codigo").Value
    varticulo = bliqui_sueldos_final.Recordset("ÚltimoDeDescrip").Value
   ' vporcentaje = bliqui_sueldos.Recordset("ÚltimoDeporcentaje")
    
    vtcobrado = vtcobrado + bliqui_sueldos_final.Recordset("SumaDepago1").Value
    vtcantidad = vtcantidad + bliqui_sueldos_final.Recordset("sumadecantidad").Value
    vttotal = vttotal + bliqui_sueldos_final.Recordset("sumadetotal").Value
    vtcdo = vtcdo + bliqui_sueldos_final.Recordset("SumaDetotal_cdo").Value
    vtctacte = vtctacte + bliqui_sueldos_final.Recordset("SumaDetotal_ctacte").Value
'    vtganancia = vtganancia + bliqui_sueldos.Recordset("SumaDeSueldo")
    
    vporcentaje = bliqui_sueldos_final.Recordset("Ganancia").Value
    
    bliqui_sueldos_final.Recordset.MoveNext
Loop


 '---------- grabo el ultimo renglon de la liqui --------------------
        '------------ grabo en el temporal ------------------
            bliqui_temp.Recordset.AddNew
            
            bliqui_temp.Recordset("codigo").Value = vcodigo
            bliqui_temp.Recordset("Articulo").Value = varticulo
            bliqui_temp.Recordset("cantidad").Value = vtcantidad
            bliqui_temp.Recordset("total").Value = vttotal
            
            
            
            bliqui_temp.Recordset("porcentaje").Value = vporcentaje
            bliqui_temp.Recordset("cdo").Value = vtcdo
            bliqui_temp.Recordset("ctacte").Value = vtctacte
            bliqui_temp.Recordset("ganancia").Value = (vporcentaje * vtcobrado) / 100
            
            vtotal_a_cobrar = vtotal_a_cobrar + bliqui_temp.Recordset("ganancia").Value
            vtotal_vendido = vtotal_vendido + (vporcentaje * vttotal) / 100
            
            bliqui_temp.Recordset("cobrado").Value = vtcobrado
            
            bliqui_temp.Recordset.Update
  
 '-----------------------------------------------------------------
 
 'Me.bliqui_temp_final.Refresh
 '
 '   Unload mantenimiento
 '   Load mantenimiento
 '
 'Me.bliqui_temp_final.Refresh
 '
 '
 '  If mantenimiento.rsLiqui.State = 1 Then
 '       mantenimiento.rsLiqui.Close
 '       mantenimiento.rsLiqui.Open
 '   Else
 '       mantenimiento.rsLiqui.Open
 '       mantenimiento.rsLiqui.Close
 '       mantenimiento.rsLiqui.Open
 '   End If
 '
 '   MsgBox "La liquidación fue efectuada correctamente.", vbInformation, "Mensaje ..."
 '
 '   With drliqui_sueldos_final
 '       .Sections("TituloEmpresa").Controls("vrepartidor").Caption = vrepartidor
 '       .Sections("TituloEmpresa").Controls("vfdesde").Caption = Str(fdesde.Value)
 '       .Sections("TituloEmpresa").Controls("vfhasta").Caption = Str(fhasta.Value)
 '
 '       .Sections("TituloEmpresa").Controls("vrepartidor").Caption = vrepartidor
 '
 '       .Sections("sección5").Controls("vanterior").Caption = Format(vajuste, "########0.00")
 '       .Sections("sección5").Controls("vcobro").Caption = Format(vtotal_a_cobrar, "########0.00")
 '       .Sections("sección5").Controls("vtotal").Caption = Format((vajuste + vtotal_a_cobrar), "########0.00")
 '
 '       .Show
 '   End With
    
    cmdLiquiActual.Enabled = False
    
    
    ' -------- ingreso el cobro a la tabla resumen_liqui --------------
    With bresumen_liqui
        .Refresh
        .Recordset.AddNew
        
        .Recordset("fecha").Value = vfecha
        .Recordset("repartidor").Value = vcod_repartidor
        .Recordset("nombre").Value = Me.vrepartidor
        .Recordset("desde").Value = fdesde
        .Recordset("hasta").Value = fhasta
        .Recordset("pago").Value = vtotal_a_cobrar
        .Recordset("total").Value = vtotal_vendido
        .Recordset("resta").Value = vtotal_vendido - vtotal_a_cobrar
        .Recordset("comentario").Value = "sín el acumulativo"

        .Recordset.Update
    End With
    '------------------------------------------------------------------
    If vConfigGral.vIncluyeContabilidad = True Then
        With frmAsientosAlta
            .Show
            .ZOrder (0)
            .txtCuentaVieneDe.Text = Me.Caption
        End With
    End If
    
If Err Then GrabarLog "cmdLiquiActual_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdLiquiAnterior_Click()
On Error Resume Next

    cobros_anteriores
    actualizacion_liquidacion
    'cobro_actual
    cmdLiquiActual.Enabled = True

If Err Then GrabarLog "cmdLiquiAnterior_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdSDetalle_Click() 'Sin detalle
On Error Resume Next

    BorrarBase "temp", pathDBMySQL
    
    If vrepartidor.Text = "" Then
        bempleados.Refresh
        bempleados.RecordSource = "Select * from empleados order by codigo"
        bempleados.Refresh
        bempleados.Recordset.MoveFirst

        Do Until bempleados.Recordset.EOF
            Liqui_repartidor fdesde, fhasta, bempleados.Recordset("codigo")
            bempleados.Recordset.MoveNext
        Loop

    Else
        Liqui_repartidor fdesde, fhasta, Val(vcod_repartidor)
    End If
    
    vcod_repartidor = ""
    
    Unload Mantenimiento
    Load Mantenimiento

    If Mantenimiento.rstemp.State = 1 Then
        Mantenimiento.rstemp.Close
        Mantenimiento.rstemp.Open
    Else
        Mantenimiento.rstemp.Open
        Mantenimiento.rstemp.Close
        Mantenimiento.rstemp.Open
    End If
    
    Mantenimiento.rstemp.Sort = "Codigo Asc"
    drliqui_sueldo2.Sections("TituloEmpresa").Controls("fdesde").Caption = fdesde
    drliqui_sueldo2.Sections("TituloEmpresa").Controls("fhasta").Caption = fhasta
    drliqui_sueldo2.Show
    MousePointer = vbDefault
    
    vrepartidor.Text = ""
    'fdesde.Value = Date
    'fhasta.Value = Date
   
If Err Then GrabarLog "cmdSDetalle_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cobros_anteriores()
' recorro resumen_liqui y veo lo que falta cobrar

    With bresumen_liqui
        .RecordSource = "select * from Resumen_liqui where repartidor = '" + vcod_repartidor + "' order by fecha ASC"
        .Refresh

        vgtresta = 0
        vgtpago = 0
        vgtsueldo = 0
        vgtotal = 0

        If Not .Recordset.EOF Then .Recordset.MoveFirst


        Do Until .Recordset.EOF = True
            vgtresta = vgtresta + Format(.Recordset("resta").Value, "######0.00")
            vgtotal = vgtotal + Format(.Recordset("total").Value, "######0.00")
            vgtpago = vgtpago + Format(.Recordset("pago").Value, "######0.00")
            .Recordset.MoveNext
        Loop
    
        vtresta = vgtotal - vgtpago

        liqui.AddItem ("> Resumen de liquidaciones anteriores: ")
        liqui.AddItem ("        Sueldo a cobrar :" + Str(vgtotal))
        liqui.AddItem ("        Total Cobrado   :" + Str(vgtpago))
        liqui.AddItem ("        Total Restante  :" + Str(vgtotal - vgtpago))

    End With
    
If Err Then GrabarLog "", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub Command5_Click()
Me.bliqui_temp_final.Refresh
End Sub

Private Sub Form_Load()
On Error Resume Next

    With bliqui_temp
        .ConnectionString = pathDBMySQL
        .RecordSource = "liqui_temp"
        .Refresh
    End With
    With bliqui_temp_final
        .ConnectionString = pathDBMySQL
        .RecordSource = "liqui_temp"
        .Refresh
    End With
    With bresumen_liqui
        .ConnectionString = pathDBMySQL
        .RecordSource = "resumen_liqui"
        .Refresh
    End With
    With bemp
        .ConnectionString = pathDBMySQL
        .RecordSource = "Select * from Liqui_Sueldo2"
        .Refresh
    End With
    With btemp
        .ConnectionString = pathDBMySQL
        .RecordSource = "Temp"
        .Refresh
    End With
    With bempleados
        .ConnectionString = pathDBMySQL
        .RecordSource = "empleados"
        .Refresh
    End With
    With bfdetalle
        .ConnectionString = pathDBMySQL
        .RecordSource = "fdetalle"
        .Refresh
    End With

    cmdCDetalle.Enabled = True
    
    With Me
        .Top = 0
        .Left = 0
        .height = 8760
        .width = 14970
        .KeyPreview = True
    End With
    frevisar.Value = strfechaMySQL("01/03/2007")
    vfecha.Value = Date
    fdesde.Value = Date
    fhasta.Value = Date
    
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Liqui_repartidor(vdesde As Date, _
                             vhasta As Date, _
                             vempleado As Integer)
    Dim vcodigo As Integer
    Dim vsueldo As Double
    Dim vnombre, vfecha As String
        
On Error Resume Next
    
    With bemp
        .RecordSource = "Select * from Liqui_Sueldo2 where Fecha >= '" & strfechaMySQL(vdesde) + "' and Fecha <= '" & strfechaMySQL(vhasta) + "' and Codigo = '" + Trim(Str(vempleado)) + "' order by Fecha"
        .Refresh

        If .Recordset.EOF Then Exit Sub
        
        .Recordset.MoveFirst

        Do Until .Recordset.EOF = True
            vsueldo = vsueldo + .Recordset("sumadesueldo")
            vfecha = .Recordset("fecha")
            vnombre = .Recordset("Nombre")
            vcodigo = .Recordset("codigo")
            
            .Recordset.MoveNext
        Loop
    
    End With
    
    With btemp
        .Refresh
        .Recordset.AddNew
        
        .Recordset("codigo") = vcodigo
        .Recordset("nombre") = vnombre
        .Recordset("saldo") = vsueldo
        .Recordset("Fecha") = vfecha
        
        .Recordset.Update
    End With

If Err Then GrabarLog "Liqui_repartidor", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub vrepartidor_KeyPress(Keyascii As Integer)
On Error Resume Next


    If Keyascii = 13 Then
        With bemp
            .RecordSource = "select * from empleados where codigo like '%" + vrepartidor.Text + "%' or nombre like '%" + vrepartidor.Text + "%'"
            .Refresh

            If Not .Recordset.EOF Then
                vrepartidor.Text = .Recordset("nombre")
                vcod_repartidor = .Recordset("codigo")
            
                bresumen_liqui.RecordSource = "select * from resumen_liqui where repartidor = '" + Trim(vcod_repartidor) + "' ORDER BY FECHA ASC"
                bresumen_liqui.Refresh
                If Not bresumen_liqui.Recordset.EOF Then bresumen_liqui.Recordset.MoveLast
            End If
    
        End With
    End If

If Err Then GrabarLog "vrepartidor_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub


