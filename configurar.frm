VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmConfigurar 
   Caption         =   "Módulo de configuración"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   11355
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "configurar.frx":0000
      Height          =   4095
      Left            =   5400
      TabIndex        =   16
      Top             =   600
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   7223
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
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   405
      Left            =   3840
      TabIndex        =   15
      Top             =   4410
      Width           =   1155
   End
   Begin MSAdodcLib.Adodc bconfigura 
      Height          =   375
      Left            =   240
      Top             =   4440
      Visible         =   0   'False
      Width           =   2640
      _ExtentX        =   4657
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
      Caption         =   "bconfigura"
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
      Caption         =   "Datos de Economía :"
      ForeColor       =   &H00000080&
      Height          =   1245
      Left            =   240
      TabIndex        =   10
      Top             =   1950
      Width           =   4755
      Begin VB.TextBox viva 
         DataField       =   "iva"
         DataSource      =   "bconfigura"
         Height          =   315
         Left            =   1920
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox vdolar 
         DataField       =   "dolar"
         DataSource      =   "bconfigura"
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Top             =   690
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "I.V.A. % :"
         Height          =   225
         Left            =   150
         TabIndex        =   14
         Top             =   390
         Width           =   1545
      End
      Begin VB.Label Label5 
         Caption         =   "Precio Dólar u$s :"
         Height          =   225
         Left            =   150
         TabIndex        =   13
         Top             =   720
         Width           =   1665
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de la Empresa :"
      ForeColor       =   &H00000080&
      Height          =   1485
      Left            =   210
      TabIndex        =   3
      Top             =   390
      Width           =   4785
      Begin VB.TextBox vnombre 
         DataField       =   "Nombre"
         DataSource      =   "bconfigura"
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   300
         Width           =   3435
      End
      Begin VB.TextBox vdireccion 
         DataField       =   "Direccion"
         DataSource      =   "bconfigura"
         Height          =   315
         Left            =   1200
         TabIndex        =   5
         Top             =   630
         Width           =   3435
      End
      Begin VB.TextBox vtelefono 
         DataField       =   "Telefono"
         DataSource      =   "bconfigura"
         Height          =   315
         Left            =   1200
         TabIndex        =   4
         Top             =   960
         Width           =   3435
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre :"
         Height          =   225
         Left            =   90
         TabIndex        =   9
         Top             =   360
         Width           =   1665
      End
      Begin VB.Label Label2 
         Caption         =   "Dirección :"
         Height          =   225
         Left            =   90
         TabIndex        =   8
         Top             =   690
         Width           =   1665
      End
      Begin VB.Label Label3 
         Caption         =   "Teléfono :"
         Height          =   225
         Left            =   90
         TabIndex        =   7
         Top             =   1020
         Width           =   1665
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Opciones de Configuración"
      ForeColor       =   &H00000080&
      Height          =   1035
      Left            =   240
      TabIndex        =   1
      Top             =   3270
      Width           =   4755
      Begin VB.CheckBox Check1 
         Caption         =   "Mostrar detalle de Factura Compra"
         DataField       =   "Fcompra"
         DataSource      =   "bconfigura"
         Height          =   285
         Left            =   930
         TabIndex        =   2
         Top             =   450
         Width           =   3105
      End
   End
   Begin VB.Label Label6 
      Caption         =   "Datos de Configuración"
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
      Height          =   225
      Left            =   1620
      TabIndex        =   0
      Top             =   60
      Width           =   2085
   End
End
Attribute VB_Name = "frmConfigurar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
