VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmQuebrantos 
   Caption         =   "Listado de quebrantos y Últimos Pagos"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   10710
   Begin VB.Frame Frame1 
      Height          =   2445
      Left            =   0
      TabIndex        =   12
      Top             =   5550
      Width           =   10545
      Begin MSAdodcLib.Adodc bempleado 
         Height          =   330
         Left            =   3360
         Top             =   1680
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
         Caption         =   "bempleado"
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
      Begin MSAdodcLib.Adodc bquebranto 
         Height          =   330
         Left            =   3360
         Top             =   2040
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
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   1770
         TabIndex        =   15
         Top             =   1740
         Width           =   1035
      End
      Begin MSAdodcLib.Adodc bcliente 
         Height          =   330
         Left            =   3360
         Top             =   960
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
      Begin VB.CommandButton cmdPasarQuebranto 
         Caption         =   "Pasar a quebranto:"
         Height          =   375
         Left            =   180
         TabIndex        =   16
         Top             =   1740
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.TextBox vdividir 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1080
         TabIndex        =   13
         Top             =   375
         Width           =   2055
      End
      Begin MSAdodcLib.Adodc bccliente 
         Height          =   330
         Left            =   3360
         Top             =   1320
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
      Begin MSDataGridLib.DataGrid DgEmpleados 
         Bindings        =   "quebrantos.frx":0000
         Height          =   1845
         Left            =   3360
         TabIndex        =   14
         Top             =   480
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   3254
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   -2147483631
         HeadLines       =   2
         RowHeight       =   15
         RowDividerStyle =   4
         FormatLocked    =   -1  'True
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
         ColumnCount     =   12
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
         BeginProperty Column02 
            DataField       =   "Direccion"
            Caption         =   "Direccion"
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
            DataField       =   "Localidad"
            Caption         =   "Localidad"
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
            DataField       =   "Telefono"
            Caption         =   "Telefono"
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
            DataField       =   "Iva"
            Caption         =   "Iva"
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
            DataField       =   "Cuit"
            Caption         =   "Cuit"
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
            DataField       =   "Credito"
            Caption         =   "Credito"
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
            DataField       =   "Responsable"
            Caption         =   "Responsable"
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
            DataField       =   "Ibrutos"
            Caption         =   "Ibrutos"
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
         BeginProperty Column11 
            DataField       =   "Quebranto"
            Caption         =   "Quebranto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2594.835
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1695.118
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   14.74
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
            BeginProperty Column09 
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column11 
               Alignment       =   1
               ColumnWidth     =   1080
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "> Cantidad de registro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   225
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1965
      End
      Begin VB.Label vcregistro 
         Alignment       =   1  'Right Justify
         Height          =   225
         Left            =   2100
         TabIndex        =   19
         Top             =   960
         Width           =   945
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Quebrantos de los CLIENTES:"
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
         Height          =   255
         Left            =   3330
         TabIndex        =   18
         Top             =   210
         Width           =   7125
      End
      Begin VB.Label lbldividir 
         AutoSize        =   -1  'True
         Caption         =   "% a Dividir:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   435
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdPagos 
      Caption         =   "Ejecutar Pagos efectuados"
      Height          =   435
      Left            =   6870
      TabIndex        =   8
      Top             =   1000
      Width           =   3375
   End
   Begin ComctlLib.ProgressBar barra 
      Height          =   285
      Left            =   -60
      TabIndex        =   5
      Top             =   5280
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CheckBox chkQuebranto 
      BackColor       =   &H80000004&
      Caption         =   "sin quebrantos"
      Height          =   255
      Left            =   6720
      TabIndex        =   4
      Top             =   210
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid DgQuebranto 
      Bindings        =   "quebrantos.frx":0018
      Height          =   3555
      Left            =   -60
      TabIndex        =   3
      Top             =   1710
      Width           =   10605
      _ExtentX        =   18706
      _ExtentY        =   6271
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      RowDividerStyle =   4
      FormatLocked    =   -1  'True
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
      ColumnCount     =   8
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
      BeginProperty Column02 
         DataField       =   "U_venta"
         Caption         =   "U. Venta"
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
         Caption         =   "U. Pago"
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
         DataField       =   "ENombre"
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
         DataField       =   "Empleados.Codigo"
         Caption         =   "Empleados.Codigo"
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
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2099.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1590.236
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker vfecha 
      Height          =   315
      Left            =   4260
      TabIndex        =   2
      Top             =   150
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   556
      _Version        =   393216
      Format          =   58916865
      CurrentDate     =   38744
   End
   Begin VB.CommandButton cmdEjecutar 
      Caption         =   "Ejecutar  !"
      Height          =   315
      Left            =   9000
      TabIndex        =   0
      Top             =   180
      Width           =   1275
   End
   Begin MSComCtl2.DTPicker vpdesde 
      Height          =   315
      Left            =   720
      TabIndex        =   6
      Top             =   1080
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      _Version        =   393216
      Format          =   58916865
      CurrentDate     =   38744
   End
   Begin MSComCtl2.DTPicker vphasta 
      Height          =   315
      Left            =   3000
      TabIndex        =   9
      Top             =   1080
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      _Version        =   393216
      Format          =   58916865
      CurrentDate     =   38744
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000015&
      Caption         =   "> Listados de Pagos relizados en un determinado  período:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   570
      Width           =   10485
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "> Hasta:"
      Height          =   255
      Left            =   2190
      TabIndex        =   10
      Top             =   1120
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "> Desde: "
      Height          =   255
      Left            =   -90
      TabIndex        =   7
      Top             =   1120
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "> Ing.  última fecha de pago con saldo deudor :"
      Height          =   255
      Left            =   90
      TabIndex        =   1
      Top             =   150
      Width           =   4155
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      Height          =   525
      Left            =   -30
      Top             =   60
      Width           =   10545
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   1095
      Left            =   -30
      Top             =   510
      Width           =   10545
   End
End
Attribute VB_Name = "frmQuebrantos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub A_Quebranto(vempleado As String, _
                        vsaldo As Double)
    On Error Resume Next

    With bempleado
        .RecordSource = "SELECT * FROM empleados WHERE (codigo = '" + vempleado + "')"
        .Refresh

        If .Recordset.EOF = True Then
            MsgBox "Ocurrió un problema al querer pasar a quebranto para el empleado: " & vempleado, vbCritical, "Quebranto ..."
            Exit Sub
        End If

    
        .Recordset("quebranto").Value = Val(Format(.Recordset("quebranto").Value, "####0.00")) + (vsaldo * (Val(vdividir) / 100))
        .Recordset.Update
    
    End With

If Err Then GrabarLog "A_Quebranto", Err.Number & " " & Err.Description, Me.Name
End Sub
Function CalcularSaldo(vCodigoCliente As String) As Double
    On Error Resume Next

    With bccliente 'Saldos_clientes
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT Max(cuentascorrientes.Fecha) AS ÚltimoDeFecha, cuentascorrientes.Codigo, Max(Clientes.Nombre) AS Nombre, Sum(cuentascorrientes.Debito) AS SumaDeDebito, Sum(cuentascorrientes.Credito) AS SumaDeCredito,  Sum(cuentascorrientes.Debito) - Sum(cuentascorrientes.credito) AS Expr1, Clientes.Direccion, clientes.idtipoiva, Clientes.Cuit, Clientes.codigonum, Clientes.idReparto FROM cuentascorrientes INNER JOIN Clientes ON cuentascorrientes.Codigo = Clientes.Codigo Where (((CuentasCorrientes.Noimputar) <> True) And ((CuentasCorrientes.fecha) < '" & strfechaMySQL(Date) + "')) GROUP BY cuentascorrientes.Codigo, Clientes.Direccion, clientes.idtipoiva, Clientes.Cuit, Clientes.codigonum, Clientes.idReparto HAVING (((cuentascorrientes.Codigo) = '" + CalcularSaldo + "'))"
        .Refresh

        If Not .Recordset.RecordCount = 0 Then
            CalcularSaldo = .Recordset("expr1").Value
        Else
            CalcularSaldo = 0
        End If
    End With
    
    If Err Then GrabarLog "CalcularSaldo", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub cmdEjecutar_Click()
Dim sql As String
On Error Resume Next
    
    sql = ""

    If chkQuebranto.Value = 1 Then
        sql = sql + " and (condicion <> 'Quebranto' or condicion is Null) and (u_pago <= '" + strfecha2(vfecha.Value) + "')"
    Else
        sql = sql + " and u_pago <= '" + strfecha2(vfecha.Value) + "'"
    End If

    With bquebranto
        .RecordSource = "SELECT * FROM quebranto WHERE 1=1 " + sql + " ORDER BY codigo"
        .Refresh
        
        vcregistro.Caption = .Recordset.RecordCount
        
    End With

    With bcliente
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM clientes WHERE 1=1" + sql
        .Refresh
        
        If Not .Recordset.EOF = True Then .Recordset.MoveFirst
    End With
    
    cmdPasarQuebranto.Visible = True

    If Err Then GrabarLog "cmdEjecutar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdPasarQuebranto_Click()
    On Error Resume Next

    If Trim(vdividir.Text) = "" Then
        MsgBox "Ingrese el % de división para asignar al repartidor", vbInformation, "Mensaje ..."
        Exit Sub
    End If

    With bcliente
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM clientes ORDER BY codigonum ASC"
        .Refresh

        If Not .Recordset.RecordCount = 0 Then Exit Sub
        
        .Recordset.MoveFirst
        
        barra.Max = .Recordset.RecordCount
        barra.Value = 0

        Do Until .Recordset.EOF = True
            .Recordset("condicion") = "Quebranto"
            .Recordset("Saldo") = CalcularSaldo(.Recordset("codigo").Value)
            
            .Recordset.MoveNext
            barra.Value = barra.Value + 1
        Loop

    End With
    
    With bquebranto
        .Refresh
    
        If .Recordset.RecordCount = 0 Then Exit Sub
        
        .Recordset.MoveFirst
        
        barra.Max = .Recordset.RecordCount
        barra.Value = 0

        Do Until .Recordset.EOF = True
            A_Quebranto .Recordset("empleados.codigo").Value, .Recordset("saldo").Value
            .Recordset.MoveNext
            barra.Value = barra.Value + 1
        Loop

        MsgBox "Se ha pasado a quebrante a todos los clientes seleccionado !", vbInformation, "Quebranto..."
    End With
    
    With bempleado
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM empleados"
        .Refresh
    End With
    
    If vConfigGral.vIncluyeContabilidad = True Then
        With frmAsientosAlta
            .Show
            .ZOrder (0)
            .txtCuentaVieneDe.Text = Me.Caption
        End With
    End If
    
If Err Then GrabarLog "cmdPasarQuebranto_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdImprimir_Click()
    On Error Resume Next
    
    Unload Mantenimiento
    Load Mantenimiento

    MsgBox "   Prepare la Impresora   ", vbInformation, "Mensaje ..."
    
    With Mantenimiento.rsquebranto
    
        If Not .State = 1 Then .Open
        .Close
        .Open
    
        If chkQuebranto.Value = 1 Then
            .filter = " ([condicion] <> 'Quebranto') AND ([u_pago] <= '" & strfechaMySQL(vfecha.Value) + "') AND ([saldo] > 0)"
        Else
            .filter = "([u_pago] <= '" + strfecha2(vfecha) + "') AND ([saldo] > 0)"
        End If
        
        .Sort = "nombre ASC"
    End With
    
    With drquebrantos
        .Show
    End With
    If Err Then GrabarLog "cmdImprimir_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdPagos_Click()
  On Error Resume Next

    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox "   Prepare la Impresora   ", vbInformation, "Mensaje ..."
    
    With Mantenimiento.rsquebranto
    
        If Not .State = 1 Then .Open
        .Close
        .Open
    
        .filter = "([u_pago] >= '" + strfecha2(vpdesde.Value) + "') AND ([u_pago] <= '" + strfecha2(vphasta.Value) + "') AND ([credito] > 0)"
        .Sort = "nombre ASC"
    
    End With
    
    With drquebrantos
        .Sections("TituloEmpresa").Controls("vfechas").Caption = "Desde : " & vpdesde.Value & " hasta: " & vphasta.Value
        .Show
    End With

    If Err Then GrabarLog "cmdPagos_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub DgEmpleados_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    
    OrdenarDataGrid ColIndex, bempleado.Recordset, DgEmpleados
    
    If Err Then GrabarLog "DgEmpleados_HeadClick", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub DgQuebranto_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    
    OrdenarDataGrid ColIndex, bquebranto.Recordset, DgQuebranto
    
    If Err Then GrabarLog "DgQuebranto_HeadClick", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Load()
    On Error Resume Next

    With bempleado
        .ConnectionString = pathDBMySQL
        .RecordSource = "Empleados"
        .Refresh
    End With

    With bquebranto
        .ConnectionString = pathDBMySQL
        .RecordSource = "Select * from quebranto"
        .Refresh
    End With
    
    With Me
        .Height = 8700
        .Width = 10830
        .Top = 50
    End With
    
    vpdesde.Value = Date
    vphasta.Value = Date

    If Err Then GrabarLog "Form_load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
