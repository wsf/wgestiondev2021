VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLibreta 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Libreta"
   ClientHeight    =   6495
   ClientLeft      =   150
   ClientTop       =   405
   ClientWidth     =   12060
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   12060
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc bsaldo 
      Height          =   330
      Left            =   60
      Top             =   7830
      Width           =   3495
      _ExtentX        =   6165
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
      Caption         =   "bsaldo"
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
   Begin TabDlg.SSTab TabLibreta 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   11245
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Generación"
      TabPicture(0)   =   "frmLibreta.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl(4)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Barra_Detalle"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Barra_Clientes"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "frmCliente"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fecha"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboReparto"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cboLista"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cboTipoListado"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdEjecutar"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lstEventos"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdLibretasEspeciales"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdImprimir"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "bpagopormes"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Clientes Libreta"
      TabPicture(1)   =   "frmLibreta.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DgTemp_Libreta"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Detalle Libretas"
      TabPicture(2)   =   "frmLibreta.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DgTemp_LibretaDetalle"
      Tab(2).Control(1)=   "lblnota"
      Tab(2).ControlCount=   2
      Begin MSAdodcLib.Adodc bpagopormes 
         Height          =   345
         Left            =   480
         Top             =   5070
         Visible         =   0   'False
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   609
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
         Caption         =   "bpagopormes"
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
         Height          =   5715
         Left            =   6180
         TabIndex        =   24
         Top             =   240
         Width           =   5625
         Begin VB.ListBox saldos 
            Height          =   4545
            Left            =   180
            TabIndex        =   25
            Top             =   780
            Width           =   5295
         End
         Begin VB.Label Label1 
            Caption         =   "ANTENCIÓN !!!!! Verifique desde el módulo Cuentas Corrientes si lo saldos de los siguientes clientes son correctos "
            ForeColor       =   &H00808080&
            Height          =   735
            Left            =   150
            TabIndex        =   26
            Top             =   270
            Width           =   5265
         End
      End
      Begin MSDataGridLib.DataGrid DgTemp_Libreta 
         Bindings        =   "frmLibreta.frx":0054
         Height          =   6075
         Left            =   -75000
         TabIndex        =   20
         Top             =   0
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   10716
         _Version        =   393216
         BackColor       =   -2147483626
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
         ColumnCount     =   7
         BeginProperty Column00 
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
         BeginProperty Column01 
            DataField       =   "Codigo_num"
            Caption         =   "Codigo_num"
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
            DataField       =   "Nombre"
            Caption         =   "Nombre"
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
            DataField       =   "Direccion"
            Caption         =   "Dirección"
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
            DataField       =   "SaldoAnterior"
            Caption         =   "S. Anterior"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$"" #,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "SaldoTotal"
            Caption         =   "S. Total"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$"" #,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column06 
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
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   4919.811
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1409.953
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   0
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir U. Listado"
         Height          =   375
         Left            =   3720
         TabIndex        =   15
         Top             =   2280
         Width           =   2415
      End
      Begin VB.CommandButton cmdLibretasEspeciales 
         Caption         =   "> Imprimir Libretas Especiales"
         Height          =   375
         Left            =   3570
         TabIndex        =   18
         Top             =   5070
         Width           =   2535
      End
      Begin VB.ListBox lstEventos 
         Height          =   2205
         Left            =   120
         TabIndex        =   17
         Top             =   2760
         Width           =   6015
      End
      Begin VB.CommandButton cmdEjecutar 
         Caption         =   "Generar Listado"
         Height          =   375
         Left            =   1440
         TabIndex        =   16
         Top             =   2280
         Width           =   2295
      End
      Begin VB.ComboBox cboTipoListado 
         BackColor       =   &H80000018&
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
         ItemData        =   "frmLibreta.frx":0070
         Left            =   1185
         List            =   "frmLibreta.frx":007D
         TabIndex        =   11
         Text            =   "Para libreta"
         Top             =   1080
         Width           =   2655
      End
      Begin VB.ComboBox cboLista 
         Height          =   315
         ItemData        =   "frmLibreta.frx":00B3
         Left            =   1185
         List            =   "frmLibreta.frx":00D2
         TabIndex        =   10
         Top             =   1800
         Width           =   2655
      End
      Begin VB.ComboBox cboReparto 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   1185
         TabIndex        =   9
         Top             =   1455
         Width           =   2655
      End
      Begin VB.Frame fecha 
         Caption         =   "Rango de Fecha"
         ForeColor       =   &H00000080&
         Height          =   1065
         Left            =   4200
         TabIndex        =   6
         Top             =   1080
         Width           =   1875
         Begin MSComCtl2.DTPicker fdesde 
            Height          =   285
            Left            =   240
            TabIndex        =   7
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Format          =   58458113
            CurrentDate     =   39295
         End
         Begin MSComCtl2.DTPicker fhasta 
            Height          =   285
            Left            =   240
            TabIndex        =   8
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Format          =   58458113
            CurrentDate     =   39325
         End
      End
      Begin VB.Frame frmCliente 
         Caption         =   "> Cliente:"
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5895
         Begin VB.TextBox vcdesde 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1200
            TabIndex        =   3
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox vchasta 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   4200
            TabIndex        =   2
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "> C. Desde:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   5
            Top             =   285
            Width           =   840
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "> C. Hasta:"
            Height          =   195
            Index           =   1
            Left            =   3240
            TabIndex        =   4
            Top             =   285
            Width           =   795
         End
      End
      Begin MSComctlLib.ProgressBar Barra_Clientes 
         Height          =   225
         Left            =   120
         TabIndex        =   19
         Top             =   5700
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSDataGridLib.DataGrid DgTemp_LibretaDetalle 
         Bindings        =   "frmLibreta.frx":00FA
         Height          =   5535
         Left            =   -75000
         TabIndex        =   21
         Top             =   0
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   9763
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "Codigo_Cliente"
            Caption         =   "Codigo_Cliente"
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
         BeginProperty Column02 
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
         BeginProperty Column03 
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
         BeginProperty Column04 
            DataField       =   "Total"
            Caption         =   "Total"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$"" #,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Saldo"
            Caption         =   "Saldo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$"" #,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column06 
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
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2415.118
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   3929.953
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   0
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ProgressBar Barra_Detalle 
         Height          =   135
         Left            =   120
         TabIndex        =   23
         Top             =   5580
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label lblnota 
         Caption         =   "* Cualquier Cambio realizado sobre esta Grilla NO TENDRA CAMBIOS sobre las ctacte."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   -74880
         TabIndex        =   22
         Top             =   5640
         Width           =   9615
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "> Lista"
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
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "> Tipo:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
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
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1005
      End
   End
   Begin MSAdodcLib.Adodc bpagos_libretas 
      Height          =   330
      Left            =   0
      Top             =   6720
      Width           =   3495
      _ExtentX        =   6165
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
   Begin MSAdodcLib.Adodc bsaldos_clientes 
      Height          =   330
      Left            =   3480
      Top             =   6720
      Width           =   3495
      _ExtentX        =   6165
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
   Begin MSAdodcLib.Adodc bLibreta 
      Height          =   330
      Left            =   6960
      Top             =   6720
      Width           =   3495
      _ExtentX        =   6165
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
   Begin MSAdodcLib.Adodc bTemp_LibretaDetalle 
      Height          =   330
      Left            =   6960
      Top             =   7080
      Width           =   3495
      _ExtentX        =   6165
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
      Caption         =   "bTemp_LibretaDetalle"
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
   Begin MSAdodcLib.Adodc bTemp_Libreta 
      Height          =   330
      Left            =   3480
      Top             =   7080
      Width           =   3495
      _ExtentX        =   6165
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
      Caption         =   "bTemp_Libreta"
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
      Left            =   6960
      Top             =   7440
      Width           =   3495
      _ExtentX        =   6165
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
   Begin MSAdodcLib.Adodc bfdetalle 
      Height          =   330
      Left            =   3480
      Top             =   7440
      Width           =   3495
      _ExtentX        =   6165
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
   Begin MSAdodcLib.Adodc bccliente 
      Height          =   330
      Left            =   0
      Top             =   7440
      Width           =   3495
      _ExtentX        =   6165
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
      Left            =   0
      Top             =   7080
      Width           =   3495
      _ExtentX        =   6165
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
   Begin MSAdodcLib.Adodc barticulos 
      Height          =   330
      Left            =   0
      Top             =   7000
      Width           =   3500
      _ExtentX        =   6165
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
End
Attribute VB_Name = "frmLibreta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vcod_reparto, vDetalle As String
Dim vcantidad As Integer
Dim vnlista As Integer
Dim vtresta, vblibreta_resta, vtotal, vdiferencia As Double
Dim sql_especiales As String
Dim vEspeciales As Boolean
Function Analiza_Codigo(unosolo As Boolean) As Boolean
Dim i As Integer
On Error Resume Next
    
    sql_especiales = ""
    
    i = 0
    
    If (unosolo = False) Then
        Do Until i = (lstEventos.ListCount)
        
            lstEventos.ListIndex = i

            sql_especiales = sql_especiales + " or (Codigo = '" + Parser_Linea(lstEventos.List(lstEventos.ListIndex)) + "')"
            
            i = i + 1
        
        Loop
    
    Else
        
        sql_especiales = sql_especiales + " or (Codigo = '" + Parser_Linea(lstEventos.List(lstEventos.ListIndex)) + "')"
    
    End If

    If Not sql_especiales = "" Then
        Analiza_Codigo = True
    Else
        Analiza_Codigo = False
    End If
    
If Err Then GrabarLog "Analiza_Codigo", Err.Number & " " & Err.Description, Me.Name
End Function
Function Busca_precio(ByRef vnlista As Integer, vcodigoart As String) As Double
    With barticulos
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "Select * from articulos"
        .Refresh
        .Recordset.Find ("Codigo = '" + vcodigoart + "'")
        If Not .Recordset.EOF = True Then
            Busca_precio = Val(Format(.Recordset("Pventa" & Val(vnlista)), "######0.00"))
        Else
            Busca_precio = 0
        End If
    End With
End Function
Function calsaldoanterior(vcliente As String) As Double
    Dim vtotal, saldo1, saldo2, vvsaldo As Double

    Dim vfhasta As Date
    On Error Resume Next
    vfhasta = fdesde.Value ''  ojo que cambio el vhasta por vfdesde

' esta es igual a la consulta bpagopormes


bpagopormes.RecordSource = "SELECT Sum(fdetalle.resta) AS resta, cuentascorrientes.anomes, cuentascorrientes.Codigo, cuentascorrientes.Noimputar FROM cuentascorrientes INNER JOIN (Factura INNER JOIN fdetalle ON Factura.Remito = fdetalle.Remito) ON cuentascorrientes.Remito = fdetalle.Remito GROUP BY cuentascorrientes.anomes, cuentascorrientes.Codigo, cuentascorrientes.Noimputar HAVING (((cuentascorrientes.Codigo)='" + Trim(vcliente) + "') AND ((cuentascorrientes.Noimputar)=0)) order by anomes"
'bpagopormes.RecordSource = "SELECT cuentascorrientes.Codigo, Sum(cuentascorrientes.Debito) AS SumaDeDebito, Sum(cuentascorrientes.Credito) AS SumaDeCredito, [SumaDeDebito]-[SumadDeCredito] AS saldo, cuentascorrientes.anomes, Where (((cuentascorrientes.Noimputar) <> True)) Or (((cuentascorrientes.Debito) > 0)) GROUP BY cuentascorrientes.Codigo, cuentascorrientes.anomes HAVING (((cuentascorrientes.Codigo)='" + vcliente + "')) ordey by anomes"
bpagopormes.Refresh


vvsaldo = 0

Do Until bpagopormes.Recordset.EOF

If Val(bpagopormes.Recordset("anomes")) < Val((Trim(Right(Me.fdesde, 4) + Trim(Left(strfechaMySQL(Me.fdesde), 2))))) Then
    'saldos.AddItem (Str(bpagopormes.Recordset(2)))
    vvsaldo = vvsaldo + bpagopormes.Recordset("resta")
End If

bpagopormes.Recordset.MoveNext
Loop
        
        If bpagopormes.Recordset.RecordCount = 0 Then
            calsaldoanterior = 0
        Else
            calsaldoanterior = Val(Format(vvsaldo, "#######0.00"))
        End If
    
    If Err Then GrabarLog "calsaldoanterior", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub cboLista_Click()
    On Error Resume Next

    Select Case Trim(cboLista.Text)

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

    CargarCombo "Listas", "Lista", cboLista, False

If Err Then GrabarLog "cboLista_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cboReparto_Click()
On Error Resume Next

    vcod_reparto = Left(cboReparto.Text, 2)

If Err Then GrabarLog "cboLista_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cboreparto_GotFocus()
On Error Resume Next

    CargarCombo "clireparto", "descrip", cboReparto, False

If Err Then GrabarLog "cboReparto_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdEjecutar_Click()
On Error Resume Next
    
    '0 - Depura las Tablas Temporales
    '1 - Filtra Clientes para tirar la libreta
    '2 - Graba en el Temporal
    '3 - Llama al Data Environoment y al Data Report

    '0-
    BorrarBase "Temp_Libreta", pathDBMySQL
    BorrarBase "Temp_LibretaDetalle", pathDBMySQL

    vEspeciales = False
    
    '1-
    If Filtra_Clientes = True Then
    
        '2-
    
        '3
        Imprime_Reporte '(False)
    
    End If
    
If Err Then GrabarLog "cmdEjecutar_Click", Err.Number & "  " & Err.Description, Me.Name
End Sub
Private Sub cmdImprimir_Click()
On Error Resume Next

    Call Imprime_Reporte '(False)

If Err Then GrabarLog "", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdLibretasEspeciales_Click()
On Error Resume Next
    
    BorrarBase "Temp_Libreta", pathDBMySQL
    BorrarBase "Temp_LibretaDetalle", pathDBMySQL
    
    vEspeciales = True
    
    If Not Analiza_Codigo(False) = True Then Exit Sub
    
    With bReparto_Repartidor
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "Select * from Reparto_Repartidor where 1=2 " + sql_especiales
        .Refresh
        If .Recordset.EOF = True Then Exit Sub
        
        Barra_Clientes.Value = 0
        Barra_Clientes.Max = .Recordset.RecordCount
        
        .Recordset.MoveFirst
        
        Do Until .Recordset.EOF = True
            DoEvents
            
            fi_cliente (.Recordset("Codigo").Value)
            .Recordset.MoveNext
        
            Barra_Clientes.Value = Barra_Clientes.Value + 1
        Loop
    
    End With
    
    Call Imprime_Reporte

If Err Then GrabarLog "cmdLibretasEspeciales_Click", Err.Number & " " & Err.Description, Caption
End Sub

Private Sub DgTemp_Libreta_DblClick()
On Error Resume Next
    
    With bTemp_LibretaDetalle
        If Not (bTemp_Libreta.Recordset.EOF = True) And Not (bTemp_Libreta.Recordset.BOF = True) Then
            
            .RecordSource = "Select * from temp_LibretaDetalle where (Codigo_Cliente = '" + bTemp_Libreta.Recordset("Codigo").Value + "') and not (Fecha is Null) order by ID ASC"
            .Refresh
            
            If Not .Recordset.EOF = True Then .Recordset.MoveLast
                
            TabLibreta.Tab = 2
            Caption = "Cliente : " & bTemp_Libreta.Recordset("Codigo").Value & " - " & bTemp_Libreta.Recordset("Nombre").Value & " / " & "S. Anterior : " & bTemp_Libreta.Recordset("SaldoAnterior")
            
        End If
    
    End With
    
If Err Then GrabarLog "DgLibretas_DblClick", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub DgTemp_Libreta_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next
    
    OrdenarDataGrid ColIndex, bTemp_Libreta.Recordset, DgTemp_Libreta

If Err Then GrabarLog "DgTemp_Libreta_HeadClick", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub DgTemp_LibretaDetalle_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next
    
    OrdenarDataGrid ColIndex, bTemp_LibretaDetalle.Recordset, DgTemp_LibretaDetalle

If Err Then GrabarLog "DgTemp_LibretaDetalle_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub fi_cliente(vcod_cliente As String)
    Dim renglon As Integer
    Dim vSaldoAnterior As Double
    On Error Resume Next
       
    'Miro si el cliente tiene Articulos para Imprimir en la consulta Asignacion
    With bLibreta
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM libreta WHERE CodCli = '" & vcod_cliente & "' and resta > 0 and (Fecha >= '" & strfechaMySQL(fdesde.Value) + "' and Fecha <= '" & strfechaMySQL(fhasta.Value) + "') order by Fecha ASC"
        .Refresh
    End With
    
    'Busco el saldo anterior del cliente
    vSaldoAnterior = Val(Format(calsaldoanterior(vcod_cliente), "##########0.00"))
    
    'Si no tiene saldo anterior o no tiene articulos para imprimir entonces sale.
    If vSaldoAnterior <= 0 And bLibreta.Recordset.RecordCount = 0 Then Exit Sub
    
    'Imprime encabezado de la libreta
    fititulos bReparto_Repartidor.Recordset("codigo"), bReparto_Repartidor.Recordset("Nombre"), bReparto_Repartidor.Recordset("Direccion"), vSaldoAnterior
    
    renglon = 0
    
    Barra_Detalle.Value = 0

    If Not bLibreta.Recordset.RecordCount = 0 Then
        
        With bpagos_libretas
            If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
            .RecordSource = "select * from pagos_libretas where codigo ='" + vcod_cliente + "' and fecha <= '" & strfechaMySQL(fhasta.Value) + "' and fecha >= '" & strfechaMySQL(fdesde.Value) + "'"
            .Refresh
            
            

            If Not .Recordset.RecordCount = 0 Then
                Barra_Detalle.Max = bLibreta.Recordset.RecordCount + .Recordset.RecordCount
                Do Until .Recordset.EOF = True
                    
                    bTemp_LibretaDetalle.Recordset.AddNew
                    
                    bTemp_LibretaDetalle.Recordset("Codigo_Cliente").Value = vcod_cliente
                    bTemp_LibretaDetalle.Recordset("Fecha").Value = ">>>>>>>" & Str(.Recordset("fecha").Value) & "    $ " & Val(Format(.Recordset("credito").Value, "######0.00"))
                    
                    bTemp_LibretaDetalle.Recordset.Update
                    
                    renglon = renglon + 1
                
                    .Recordset.MoveNext
                    Barra_Detalle.Value = Barra_Detalle.Value + 1
                Loop
            End If
        
        End With
        
        bLibreta.Recordset.MoveFirst
        
        Dim vtotalfactura As Double
        vtotalfactura = 0
        vtresta = 0
        
        ' -------------------------  Imprime pagos efectuados  ------------------------------------------------------------------
        
        Do Until bLibreta.Recordset.EOF = True
            fi_detalle
            bLibreta.Recordset.MoveNext
            Barra_Detalle.Value = Barra_Detalle.Value + 1
        Loop
        
        renglon = renglon + bLibreta.Recordset.RecordCount
    End If

    'bTemp_Libreta.Recordset("SaldoTotal").Value = Str(Trim(Val(Format(vsaldoanterior + vtresta, "#######0.00"))))
    bTemp_Libreta.Recordset("SaldoTotal").Value = Str(Trim(Val(Format(vSaldoAnterior + vtresta, "#######0.00"))))
    
    ' doing ------- verifico si el saldo es el mismo del las ctacte -------------
    bsaldo.RecordSource = "SELECT * FROM saldos_clientes WHERE codigo = '" + vcod_cliente + "'"
    bsaldo.Refresh
    
    If Not bsaldo.Recordset.EOF Then
        If Val(bTemp_Libreta.Recordset("SaldoTotal").Value) - bsaldo.Recordset("expr1") > 0.001 Then
            saldos.AddItem (vcod_cliente + "Libreta: " + Trim(bTemp_Libreta.Recordset("SaldoTotal").Value) + " Real: " + Trim(bsaldo.Recordset("expr1")))
        End If
    End If
    
    '--------------------------------------------------------------------
    
    bTemp_Libreta.Recordset.Update
    
    vtresta = 0 ' una vez impreso el saldo final se pone la variable en cero

    Dim i As Integer
    
    For i = (renglon + 1) To 19
        With bTemp_LibretaDetalle
            .Recordset.AddNew
            .Recordset("Codigo_Cliente") = vcod_cliente
            .Recordset("Detalle") = ""
            .Recordset.Update
        End With
    Next i
    
    If (renglon > 19) And (vEspeciales = False) Then
        
        'CONTROLAR LOS CLIENTES CON MAS MOVIMIENTOS Y PASARLOS A OTRO LADO
        BorrarBase "Temp_LibretaDetalle where Codigo_Cliente = '" + vcod_cliente + "'", pathDBMySQL
        BorrarBase "Temp_Libreta where codigo = '" + vcod_cliente + "'", pathDBMySQL
        
        'No tocar el "-" porque el parser saca el codigo de ahí.
        lstEventos.AddItem bTemp_Libreta.Recordset("Codigo") & "-" & bTemp_Libreta.Recordset("Nombre")
    End If
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
        'display.AddItem (bLibreta.Recordset("nombre") + "  " + bLibreta.Recordset("fdetalle.codigo") + " Diferencia: " + Str(vadd_resta))
    End If
    
    
    If vblibreta_resta > 0 Then
        vtresta = vtresta + vblibreta_resta
    Else
        vtresta = vtresta + (bLibreta.Recordset("resta"))
    End If

    'Printer.Print Str(bLibreta.Recordset("fecha")) & " " & num_i2(vcantidad) & "  " & char_i(vdetalle, 28) & num_i(vtotal) & vresta_comentario & num_i(Format(vtresta, "#########0.00"))
    With bTemp_LibretaDetalle
        .Refresh
        
        .Recordset.AddNew
        
        .Recordset("Codigo_Cliente").Value = bLibreta.Recordset("CodCli")
        .Recordset("Fecha").Value = strfecha2(bLibreta.Recordset("fecha"))
        .Recordset("Cantidad").Value = vcantidad
        .Recordset("Detalle").Value = vDetalle
        .Recordset("Total").Value = Trim(Str(vtotal))
        .Recordset("Saldo").Value = Trim(Str(vtresta))
        
        .Recordset.Update
    
    
    End With
    

    fi_detalle = vtotalnuevo
    If Err Then GrabarLog "fi_detalle", Err.Number & " " & Err.Description, Me.Name
End Function
Function Filtra_Clientes() As Boolean
Dim sql As String
    On Error Resume Next
        
    With bReparto_Repartidor
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        
        sql = ""
        lstEventos.Clear
            
        If Not vcdesde.Text = "" Then sql = " and Codigonum >= " & Trim(vcdesde)
        If Not vchasta.Text = "" Then sql = sql + " and Codigonum <= " & Trim(vchasta)
        
        If Val(vcdesde) > 0 Then
            
            .RecordSource = "select * from reparto_repartidor where codigo = '" + Trim(vcdesde) + "'"
            .Refresh

        Else
             
            .RecordSource = "select * from reparto_repartidor where cod_reparto = '" + Trim(vcod_reparto) + "'" + sql + " order by nombre"
            .Refresh
                  
        End If
    
        If .Recordset.EOF = True Then
            MsgBox "No existen datos para imprimir", vbInformation, "Error ..."
            Filtra_Clientes = False
            Exit Function
        End If
    
        Barra_Clientes.Value = 0
        Barra_Clientes.Max = .Recordset.RecordCount
    
        Do Until .Recordset.EOF = True
            DoEvents
            fi_cliente (.Recordset("codigo"))
            .Recordset.MoveNext
            Barra_Clientes.Value = Barra_Clientes.Value + 1
        Loop
        
        Filtra_Clientes = True
        
    End With
    If Err Then GrabarLog "flibreta", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub fititulos(vcodigocli, _
                      vnombrecli, _
                      vdireccioncli As String, _
                      vsaldoante As Double)
    On Error Resume Next
    
    With bTemp_Libreta
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "Select * from Temp_Libreta"
        .Refresh
        
        .Recordset.AddNew
        
        .Recordset("Codigo").Value = vcodigocli
        .Recordset("Codigo_Num").Value = Val(vcodigocli)
        .Recordset("Nombre").Value = vnombrecli
        .Recordset("Direccion").Value = vdireccioncli
        .Recordset("SaldoAnterior").Value = Val(Format(vsaldoante, "#########0.00"))
        
    End With
    
    If Err Then GrabarLog "fititulos", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub fmodif_ctacte(vremito As Long, vtotalmodif As Double)
On Error Resume Next
    With bccliente
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "Select * from cuentascorrientes where remito = " & Trim(vremito)
        .Refresh
        If Not .Recordset.RecordCount = 0 Then .Recordset.MoveFirst
        .Recordset("debito") = .Recordset("debito") + vtotalmodif
        .Recordset.Update
    End With
If Err Then GrabarLog "Fmodif_ctacte", Err.Number & " " & Err.Description, Me.Name
End Sub
Function fmodif_fact(vremito As Long, vtotalmodif As Double) As Long
On Error Resume Next
    With bfactura
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
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

On Error Resume Next


If vprecio = 0 Then Exit Sub

    With bfdetalle
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "Fdetalle"
        .Refresh
        .Recordset.filter = ("codigo = '" + vcodigoart + "' and remito = " & vremito & "")
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
If Err Then GrabarLog "fmodif_fdetalle", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then MsgBox "OK"
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Load()
On Error Resume Next
    
    With TabLibreta
        .Tab = 0
    End With

    With bReparto_Repartidor
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "Reparto_Repartidor"
        .Refresh
    End With
        
  
    
    
    With bTemp_Libreta
        .ConnectionString = pathDBMySQL
        .RecordSource = "Temp_Libreta"
        .Refresh
    End With
    
    With bTemp_LibretaDetalle
        .ConnectionString = pathDBMySQL
        .RecordSource = "Temp_LibretaDetalle"
        .Refresh
    End With
    
    With bsaldo
        .ConnectionString = pathDBMySQL
        .RecordSource = "saldos_clientes"
        .Refresh
    End With
    
    
    With bpagopormes
        .ConnectionString = pathDBMySQL
        .RecordSource = "Pagopormes"
        .Refresh
    End With
    
    
If Err Then GrabarLog "Form_Load", Err.Number & "  " & Err.Description, Me.Name
End Sub
Private Sub Imprime_Reporte() 'Llamada al Data Report
On Error Resume Next

    
    If Not Printer.PaperSize = vbPRPSLetter Then
        MsgBox "El papel configurado no corresponde al Tamaño CARTA"
        Exit Sub
    End If
    
    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox "     Prepare la Impresora     ", vbInformation, "Mensaje ..."
    
    With Mantenimiento.rsLibreta
        
        If Not .State = 1 Then
            .Open
            .Close
            .Open
        Else
            .Close
            .Open
        End If
        
        .filter = ("ID > -10000")
        .Sort = "Nombre ASC, Codigo_num ASC"
        
        If .RecordCount = 0 Then
            MsgBox "No existen Datos para generar libretas para clientes!! ", vbInformation, "Mensaje ..."
            Exit Sub
        End If
    End With
    
    With drLibreta
        If vEspeciales = True Then .Sections("Libreta_Footer").ForcePageBreak = rptPageBreakAfter
            
        .Sections("Libreta_Header").Controls("lblfecha").Caption = Date
        .Show
    End With
    
If Err Then GrabarLog "Imprime_Reporte", Err.Number & "  " & Err.Description, Me.Name
End Sub

Private Sub lstEventos_DblClick()
On Error Resume Next

    BorrarBase "Temp_Libreta", pathDBMySQL
    BorrarBase "Temp_LibretaDetalle", pathDBMySQL
    
    vEspeciales = True
    
    If Not Analiza_Codigo(True) = True Then Exit Sub


    With bReparto_Repartidor
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "Select * from Reparto_Repartidor where 1=2 " + sql_especiales
        .Refresh
        If Not .Recordset.RecordCount = 0 Then .Recordset.MoveFirst
        
        Do Until .Recordset.EOF = True
            DoEvents
            
            fi_cliente (.Recordset("Codigo").Value)
            .Recordset.MoveNext
        
        
        Loop
    
    End With

    Call Imprime_Reporte
    
If Err Then GrabarLog "lstEventos_DblClick", Err.Number & " " & Err.Description, Me.Name
End Sub
Function Parser_Linea(vlinea As String) As String
Dim vCodigo As String
Dim i As Integer
    
    i = 1
    Do Until vCodigo = "-"
        
        vCodigo = Mid(vlinea, i, 1)
        i = i + 1
    
    Loop
    
    Parser_Linea = Left(vlinea, i - 2)

End Function

Private Sub TabLibreta_Click(PreviousTab As Integer)
    
    Height = 6825
    Select Case TabLibreta.Tab
    
        Case 0
            Width = 6345
        Case 1
            Width = 10350
        Case 2
        
    
    End Select

    
End Sub

Private Sub vcdesde_KeyPress(Tecla As Integer)

    If Tecla = 13 Then vchasta.SetFocus
    
End Sub

