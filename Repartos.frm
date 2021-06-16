VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmRepartos 
   Caption         =   "Confección de repartos"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   16860
   Begin MSAdodcLib.Adodc bpagopormes 
      Height          =   330
      Left            =   6240
      Top             =   8040
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
   Begin MSAdodcLib.Adodc bconfigura 
      Height          =   330
      Left            =   3120
      Top             =   8010
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
   Begin MSAdodcLib.Adodc bcomentario 
      Height          =   330
      Left            =   9120
      Top             =   7680
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
      Caption         =   "bcomentario"
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
   Begin VB.TextBox cli_extra 
      Height          =   285
      Left            =   3660
      TabIndex        =   42
      Top             =   2490
      Width           =   1005
   End
   Begin VB.ListBox problemas 
      Height          =   645
      Left            =   7170
      TabIndex        =   37
      Top             =   2580
      Width           =   4965
   End
   Begin MSAdodcLib.Adodc bccliente 
      Height          =   330
      Left            =   120
      Top             =   8010
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
   Begin MSAdodcLib.Adodc bcliente 
      Height          =   330
      Left            =   9105
      Top             =   7320
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
   Begin MSAdodcLib.Adodc bsaldos_clientes 
      Height          =   330
      Left            =   3105
      Top             =   6600
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
   Begin MSAdodcLib.Adodc bclirepartos 
      Height          =   330
      Left            =   6120
      Top             =   7320
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
      Caption         =   "bclirepartos"
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
   Begin MSAdodcLib.Adodc barticulos_Ganancia 
      Height          =   330
      Left            =   90
      Top             =   6960
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
      Caption         =   "barticulos_Ganancia"
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
      Left            =   3120
      Top             =   7680
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
   Begin MSAdodcLib.Adodc barticulos_Clientes 
      Height          =   330
      Left            =   90
      Top             =   6600
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
      Caption         =   "barticulos_Clientes"
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
      Left            =   3105
      Top             =   7320
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
   Begin MSAdodcLib.Adodc bfactura 
      Height          =   330
      Left            =   90
      Top             =   7320
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
   Begin MSAdodcLib.Adodc bultima_compra 
      Height          =   330
      Left            =   3105
      Top             =   6960
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
      Caption         =   "bultima_compra"
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
   Begin MSAdodcLib.Adodc bimprime_Reparto 
      Height          =   330
      Left            =   9120
      Top             =   6600
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
      Caption         =   "bimprime_Reparto"
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
   Begin MSAdodcLib.Adodc bCliRep 
      Height          =   330
      Left            =   6120
      Top             =   6600
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
      Caption         =   "bCliRep"
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
   Begin MSAdodcLib.Adodc bclientes_repartos 
      Height          =   330
      Left            =   6120
      Top             =   7680
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
      Caption         =   "bclientes_repartos"
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
   Begin VB.Timer Espera_tecla 
      Enabled         =   0   'False
      Left            =   12150
      Top             =   6630
   End
   Begin MSAdodcLib.Adodc bdevol 
      Height          =   330
      Left            =   9120
      Top             =   6960
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
      Caption         =   "bdevol"
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
   Begin MSAdodcLib.Adodc breparto 
      Height          =   330
      Left            =   6105
      Top             =   6960
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
      Caption         =   "breparto"
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
   Begin MSAdodcLib.Adodc bgral 
      Height          =   330
      Left            =   120
      Top             =   7680
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
      Caption         =   "bgral"
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
   Begin VB.Frame Frame4 
      Caption         =   "[Datos de los artículos asociado al Reparto:]"
      ForeColor       =   &H00A5715F&
      Height          =   1455
      Left            =   7170
      TabIndex        =   12
      Top             =   3300
      Width           =   4965
      Begin VB.TextBox valias 
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   720
         TabIndex        =   27
         Top             =   990
         Width           =   1815
      End
      Begin VB.CommandButton Command9 
         Height          =   435
         Left            =   4500
         Picture         =   "Repartos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   810
         Width           =   375
      End
      Begin VB.CommandButton Command8 
         Height          =   435
         Left            =   4110
         Picture         =   "Repartos.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   810
         Width           =   405
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "Confirmar"
         Height          =   285
         Left            =   2730
         TabIndex        =   15
         Top             =   1050
         Width           =   1245
      End
      Begin VB.TextBox varticulo 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   60
         TabIndex        =   14
         Top             =   660
         Width           =   4005
      End
      Begin VB.Label Label8 
         Caption         =   "> Alias:"
         Height          =   255
         Left            =   60
         TabIndex        =   26
         Top             =   1050
         Width           =   675
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Buscar Siguiente"
         Height          =   405
         Left            =   4170
         TabIndex        =   23
         Top             =   300
         Width           =   705
      End
      Begin VB.Label Label4 
         Caption         =   "> Ingresar artículo asociado al reparto: "
         Height          =   315
         Left            =   60
         TabIndex        =   13
         Top             =   420
         Width           =   3285
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1755
      Left            =   0
      TabIndex        =   6
      Top             =   3030
      Width           =   6795
      Begin VB.CheckBox chkfact 
         Caption         =   "Generar Facturas con Detalles"
         Height          =   255
         Left            =   1560
         TabIndex        =   35
         Top             =   1440
         Width           =   2535
      End
      Begin VB.ComboBox vlocalidad 
         BackColor       =   &H80000016&
         Height          =   315
         ItemData        =   "Repartos.frx":0884
         Left            =   1680
         List            =   "Repartos.frx":089D
         TabIndex        =   25
         Text            =   "Wheelwright"
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox vcod_reparto 
         BackColor       =   &H80000016&
         Height          =   345
         Left            =   1680
         TabIndex        =   20
         Top             =   570
         Width           =   1335
      End
      Begin VB.TextBox vrepartidor 
         BackColor       =   &H80000016&
         Height          =   315
         Left            =   4680
         TabIndex        =   17
         Top             =   960
         Width           =   1875
      End
      Begin VB.ComboBox Reparto 
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
         Left            =   1680
         TabIndex        =   7
         Top             =   180
         Width           =   4900
      End
      Begin MSComCtl2.DTPicker vfecha 
         Height          =   345
         Left            =   4920
         TabIndex        =   8
         Top             =   570
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   609
         _Version        =   393216
         Format          =   66191361
         CurrentDate     =   36927
      End
      Begin VB.Label vccomentarios 
         Height          =   255
         Left            =   4920
         TabIndex        =   47
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "> Codigo del reparto:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1470
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "> Repartidor :"
         Height          =   195
         Left            =   3600
         TabIndex        =   18
         Top             =   1005
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "> Localidad:"
         Height          =   195
         Left            =   720
         TabIndex        =   11
         Top             =   1005
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "> Fecha del reparto:"
         Height          =   195
         Left            =   3360
         TabIndex        =   10
         Top             =   600
         Width           =   1425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "> Nombre del reparto:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1530
      End
   End
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   8070
      TabIndex        =   5
      Top             =   5190
      Width           =   4185
      Begin VB.CommandButton cmdImprime_Comentarios 
         Caption         =   "Comentarios"
         Enabled         =   0   'False
         Height          =   555
         Left            =   1680
         Picture         =   "Repartos.frx":08DF
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1065
      End
      Begin VB.CommandButton cmdUltimo 
         Caption         =   "Último"
         Height          =   555
         Left            =   880
         Picture         =   "Repartos.frx":09E1
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   800
      End
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "Imprimir"
         Enabled         =   0   'False
         Height          =   555
         Left            =   0
         Picture         =   "Repartos.frx":0AE3
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   885
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir "
         Height          =   555
         Left            =   3480
         Picture         =   "Repartos.frx":0BE5
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   675
      End
      Begin VB.CommandButton cmdAnula 
         Caption         =   "Anula"
         Height          =   555
         Left            =   2760
         Picture         =   "Repartos.frx":0CE7
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   705
      End
   End
   Begin VB.Frame Frame3 
      Height          =   705
      Left            =   0
      TabIndex        =   0
      Top             =   5190
      Width           =   7905
      Begin VB.CommandButton cmdCarga_Comentario 
         Caption         =   "Cargar Comentario"
         Height          =   495
         Left            =   4350
         Picture         =   "Repartos.frx":1219
         TabIndex        =   2
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton cmdBorrar 
         Caption         =   "Borrar"
         Height          =   495
         Left            =   3270
         Picture         =   "Repartos.frx":131B
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton cmdMostrarTodos 
         Caption         =   "Mostrar todos"
         Height          =   495
         Left            =   2190
         Picture         =   "Repartos.frx":141D
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "Consultar"
         Height          =   495
         Left            =   1110
         Picture         =   "Repartos.frx":151F
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         Height          =   495
         Left            =   30
         Picture         =   "Repartos.frx":1621
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid plani 
      Bindings        =   "Repartos.frx":1723
      Height          =   2355
      Left            =   60
      TabIndex        =   16
      Top             =   0
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   4154
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483626
      ForeColor       =   0
      HeadLines       =   2
      RowHeight       =   15
      RowDividerStyle =   6
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "Id"
         Caption         =   "Id"
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
         DataField       =   "Cod_Reparto"
         Caption         =   "Cod_Reparto"
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
         DataField       =   "Reparto"
         Caption         =   "Reparto"
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
         DataField       =   "Cod_Articulo"
         Caption         =   "Cod_Articulo"
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
         DataField       =   "Alias"
         Caption         =   "Alias"
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
         DataField       =   "Articulo"
         Caption         =   "Articulo"
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
      BeginProperty Column08 
         DataField       =   "Cod_Repartidor"
         Caption         =   "Cod_Repartidor"
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
         DataField       =   "Repartidor"
         Caption         =   "Repartidor"
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
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1349.858
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1844.787
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   3644.788
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2055.118
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1709.858
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox dibu 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      Picture         =   "Repartos.frx":173A
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   30
      Top             =   90
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   150
      Left            =   2790
      Top             =   150
   End
   Begin MSComctlLib.ProgressBar p 
      Height          =   255
      Left            =   30
      TabIndex        =   43
      Top             =   5010
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar barra2 
      Height          =   45
      Left            =   60
      TabIndex        =   44
      Top             =   4920
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   79
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label14 
      Caption         =   "> Clientes que deben ser verificados: "
      ForeColor       =   &H00A5715F&
      Height          =   255
      Left            =   7140
      TabIndex        =   45
      Top             =   2370
      Width           =   4965
   End
   Begin VB.Label Label13 
      Caption         =   "> Generar Reparto de un solo cliente. Ingresar nro:"
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   30
      TabIndex        =   41
      Top             =   2520
      Width           =   3585
   End
   Begin VB.Label Label12 
      Caption         =   "> de:"
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   3420
      TabIndex        =   40
      Top             =   2790
      Width           =   405
   End
   Begin VB.Label Label11 
      Caption         =   "> Cantidad de Clientes Procesados:"
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   30
      TabIndex        =   39
      Top             =   2790
      Width           =   2565
   End
   Begin VB.Label vclirep 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   195
      Left            =   2640
      TabIndex        =   38
      Top             =   2790
      Width           =   735
   End
   Begin VB.Label vcli 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   195
      Left            =   3870
      TabIndex        =   36
      Top             =   2790
      Width           =   645
   End
   Begin VB.Label Label10 
      Caption         =   "Espere un momento por favor !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   585
      Left            =   1200
      TabIndex        =   29
      Top             =   1410
      Width           =   10665
   End
   Begin VB.Label Label9 
      Caption         =   "Confeccionando planilla de Reparto ...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   585
      Left            =   1230
      TabIndex        =   28
      Top             =   570
      Width           =   10665
   End
End
Attribute VB_Name = "frmRepartos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vfcliente As String
Dim nremito As Long
Dim sql As String
Dim pr As Integer
Dim vcod_articulo, vcod_repartidor As String
Dim vEscape As Boolean
Dim vfecha_uc As String
Dim vganancia As Double
Dim vrepcompleto As Integer
Dim urenglon As Integer
Dim vcodigoant As String 'Para control de la variable vcodigo (Codigo del Cliente)
Private Sub BorrarFactura(ncliente As String)
    On Error Resume Next

    With bfactura
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "Select * from Factura where Codigo = '" + ncliente + "' and Fecha = '" & strfechaMySQL(vfecha) + "' and (total = 0 and total_ctacte = 0) order by codigo"
        .Refresh
    
        If Not .Recordset.RecordCount = 0 Then .Recordset.MoveFirst

        Do Until .Recordset.EOF
            BorrarFDetalle .Recordset("remito").Value
            .Recordset.Delete
            .Recordset.MoveNext
        Loop

    End With

    If Err Then GrabarLog "BorrarFactura (" & ncliente & ")", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub BorrarFDetalle(nremito As String)
    On Error Resume Next

    Call BorrarBase("fdetalle WHERE (remito = " & Val(nremito) & ")", pathDBMySQL)
    
    If Err Then GrabarLog "BorrarFDetalle (" & nremito & ")", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub BorrarReparto()
    On Error Resume Next

    With bcliente
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "Select * from clientes where reparto = '" + Trim(vcod_reparto) + "' order by reparto"
        .Refresh

        If Not .Recordset.RecordCount = 0 Then .Recordset.MoveFirst

        Do Until .Recordset.EOF
            BorrarFactura (.Recordset("Codigo").Value)
            .Recordset.MoveNext
        Loop

    End With

    If Err Then GrabarLog "BorrarReparto", Err.Number & " " & Err.Description, Me.Name
End Sub
Function calsaldoanterior(vfhasta As Date, vCodigoCliente As String) As Double
    Dim vtotal As Double
    On Error Resume Next

    With bsaldos_clientes
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT Max(cuentascorrientes.Fecha) AS ÚltimoDeFecha, cuentascorrientes.Codigo, Max(Clientes.Nombre) AS Nombre, Sum(cuentascorrientes.Debito) AS SumaDeDebito, Sum(cuentascorrientes.Credito) AS SumaDeCredito, Sum(cuentascorrientes.Debito)- Sum(cuentascorrientes.credito) AS Expr1, Clientes.Direccion, clientes.idtipoiva, Clientes.Cuit, Clientes.codigonum, Clientes.idReparto FROM cuentascorrientes INNER JOIN Clientes ON cuentascorrientes.Codigo = Clientes.Codigo Where (((CuentasCorrientes.Noimputar) = false ) And ((CuentasCorrientes.fechaInput) <= '" & strfechaMySQL(vfhasta) + "')) GROUP BY cuentascorrientes.Codigo, Clientes.Direccion, clientes.idtipoiva, Clientes.Cuit, Clientes.codigonum, Clientes.idReparto HAVING (((cuentascorrientes.Codigo)='" + vCodigoCliente + "'))"
        .Refresh

        If .Recordset.RecordCount = 0 Then
            calsaldoanterior = 0
        Else
            calsaldoanterior = .Recordset("expr1").Value
        End If
    End With

    If Err Then GrabarLog "calsaldoanterior", Err.Number & " " & Err.Description, Me.Name
End Function

Private Sub cmdAnula_Click() 'Anula Reparto
    On Error Resume Next
    fimpresion
'    If MsgBox(" ¿ Desea Anular el reparto ? ", vbYesNo, "Mensaje ...") = vbYes Then
'        bclirep.Refresh

'        If bclirep.Recordset.EOF Then Exit Sub
'        bclirep.Recordset.Delete
'        bclirep.Recordset.Update
'    End If
    
'    If Err Then
'        grabarlog "cmdanularep_Click", Err.Number & " " & Err.Description, Me.Name
'        MsgBox "Error en cmdanularep", vbInformation, "ERROR", Me.Name
'    End If

End Sub

Private Sub cmdBorrar_Click() 'Borra Articulos
    On Error Resume Next
    
    Borrar breparto, True
    
    If Err Then
        GrabarLog "cmdBorrar_Click", Err.Number & " " & Err.Description, Me.Name
        MsgBox "No se pudo completar la operacion-Command1", vbInformation, "Mensaje ..."
    End If

End Sub
Private Sub cmdCarga_Comentario_Click()
Dim vcomentario As String
On Error Resume Next
    
    With bcomentario
        vcomentario = InputBox("Ingrese el comentario a cargar en el reparto..: ", "Mensaje ...", "Hugo Dice..: ")
        
            If Not vcomentario = "" Then
                If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
                .RecordSource = "Temp_Comentarios"
                .Refresh
                
                .Recordset.AddNew
                
                .Recordset("Codigo") = ""
                .Recordset("Nombre") = ""
                .Recordset("Comentario") = vcomentario
                .Recordset("Cod_Reparto") = ""
                .Recordset("Repartidor") = ""
            
                .Recordset.Update
        
            End If
        
    End With
    
If Err Then GrabarLog "cmdCarga_Comentario_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdConfirmar_Click() 'Agrega nuevo articulo
    On Error Resume Next

    With breparto

        If valias.Text = "" Or varticulo.Text = "" Then
            MsgBox "Ingrese un articulo o alias", vbInformation, "Mensaje ..."
            varticulo.SetFocus
            Exit Sub
        End If
        
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "Select * from reparto where cod_reparto = '" + vcod_reparto + "' and articulo = '" + varticulo + "' and Cod_repartidor = '" + vcod_repartidor + "'"
        .Refresh
    
        If Not .Recordset.RecordCount = 0 Then
            MsgBox "El articulo ya esta cargado en este reparto", vbInformation, "Mensaje ..."
            varticulo.SetFocus
            Exit Sub
        End If

        .Recordset.AddNew
        .Recordset("reparto") = Reparto.Text
        .Recordset("cod_reparto") = vcod_reparto.Text
        .Recordset("articulo") = varticulo.Text
        .Recordset("cod_articulo") = Val(vcod_articulo)
        .Recordset("localidad") = vlocalidad.Text
        .Recordset("alias") = valias.Text
        .Recordset("repartidor") = vrepartidor.Text
        .Recordset("cod_repartidor") = vcod_repartidor
        .Recordset("fecha") = vfecha
        .Recordset.Update
    
        Limpiar
    
        plani.BackColor = &H80000018
        cmdConsultar_Click
        .Refresh

    End With

    
If Err Then GrabarLog "cmdConfirmar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdConsultar_Click()
    On Error Resume Next
    
    plani.BackColor = &H80000018
    cmdGenerar.Enabled = True
    sql = ""

    If Not Reparto.Text = "" Then sql = sql + " and (reparto like '%" + Trim(Reparto.Text) + "%')"
    If Not vrepartidor.Text = "" Then sql = sql + " and (repartidor = '" + Trim(vrepartidor.Text) + "')"
    If Not (vlocalidad.Text = "" Or vlocalidad = "Todas") Then sql = sql + " and (localidad like '%" + Trim(vlocalidad.Text) + "%')"

    With breparto
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "select * from reparto where 1=1 " + sql + " order by ID ASC"
        .Refresh
    End With
        
    Mantenimiento.rsRepartoI.filter = "id > 0 " + sql

    FiltrarComentarios (sql)
    
    
If Err Then GrabarLog "cmdConsultar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdGenerar_Click()
    On Error Resume Next
    
    Dim j As Long
    
    'plani.Visible = False

    ' vcod_repartidor = bgral.Recordset("codigo")
    '------------ borro la base de datos temporal que contiene el reparto ---------------
    BorrarBase "imprime_reparto", pathDBMySQL
    
    With bimprime_Reparto
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM Imprime_Reparto"
        .Refresh
    
       If .Recordset.RecordCount > 0 Then
            MsgBox "Intente nuevamene borrar el reparto. El sistema no está disponible en este instante", vbCritical
            Exit Sub
        End If
    End With
    '------------------------------------------------------------------------------------
    
    '-----------Guardo datos para el boton imprimir
    Call GuardarUDato(0, Reparto.Text, vrepartidor.Text, vfecha.Value)

    '---------- filtro los clientes del repartidor -----------------

    With bclientes_repartos
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
            If Val(cli_extra.Text) > 0 Then
                .RecordSource = "SELECT * FROM clientes WHERE (codigo ='" + Trim(cli_extra.Text) + "')"
            Else
                .RecordSource = "SELECT * FROM clientes WHERE (reparto = '" + Trim(vcod_reparto.Text) + "' and pasivo = 'NO')"
            End If
            .Refresh
    End With
    '-----------------------------------------------------------------
    bimprime_Reparto.Refresh
    
    'bimprime_Reparto.Refresh
    p.Refresh
    p.Max = bclientes_repartos.Recordset.RecordCount
    p.Value = 0
    vcli.Caption = Str(bclientes_repartos.Recordset.RecordCount)
    
    Dim vrespuesta As Byte
    
    vrespuesta = MsgBox("¿ Confirma la confección del listado con Reportes ?", vbInformation + vbYesNoCancel)
        
    If vrespuesta = 2 Then
        MousePointer = vbDefault
        Exit Sub
    End If
    
    'Me.Width = 14790
    
    Call GuardarUDato(1, "", "", vfecha.Value)
    
    
    If vcod_repartidor = "" Then
        MsgBox "No tiene ningun repartidor ingresado", vbInformation, "Mensaje ..."
        MousePointer = vbDefault
        Exit Sub
    End If
    
    MousePointer = vbHourglass
    
    cmdImprime_Comentarios.Enabled = True
    
    j = 0
    bimprime_Reparto.Refresh
    vclirep.Caption = "0"
    
    BorrarBase "Temp_Comentarios", pathDBMySQL
    With bclientes_repartos
    
    
        Do Until .Recordset.EOF = True
            DoEvents
                
            freparto .Recordset("codigo"), Trim(.Recordset("nombre")), Trim(.Recordset("Direccion")), .Recordset("Credito")
            
            
            j = j + 1
            
            If Not j = Val(vclirep.Caption) Then
                problemas.AddItem (Trim(.Recordset("nombre")))
                j = Val(vclirep.Caption)
                problemas.Refresh
            End If
            .Recordset.MoveNext
            p.Value = p.Value + 1
            
            If vEscape = True Then Exit Do
    
        Loop
    
    End With
    
    MousePointer = vbDefault
    
    If Val(vcli.Caption) = Val(vclirep.Caption) Then
        
        MsgBox "El Reparto fue confeccionado correctamente ", vbInformation, "Mensaje"
    Else
        MsgBox "Faltaron procesar " + Str(Val(vcli.Caption) - Val(vclirep.Caption)), vbCritical, "Mensaje"
    End If

    'plani.Visible = True
    If vrespuesta = 6 Then
        With Mantenimiento.rsImprime_Reparto
        
            If .State = 1 Then
                .Close
                .Open
            Else
                .Open
                .Close
                .Open
            End If

            .Sort = "Dirección ASC, Codigo ASC"

        End With
        
        'If breparto.Recordset.RecordCount <= 10 Then
        '    With drReparto2
        '        .Sections(2).Controls("Nom_Reparto").Caption = Trim(Reparto.Text)
        '        .Sections(2).Controls("Nom_repartidor").Caption = Trim(vrepartidor.Text)
        '        .Sections(2).Controls("sfecha_confeccion").Caption = "(" & vfecha.Value & ")"
        '        .Show
        '    End With
        'Else 'Para mas de 10 articulos
        '    With drReparto
        '        .Sections(2).Controls("Nom_Reparto").Caption = Trim(Reparto.Text)
        '        .Sections(2).Controls("Nom_repartidor").Caption = Trim(vrepartidor.Text)
        '        .Sections(2).Controls("sfecha_confeccion").Caption = "(" & vfecha.Value & ")"
        '        .Show
        '    End With
        'End If
    End If
    
    If vrespuesta = 7 Then
        fimpresion
    End If
    
    ControlarReparto

    plani.Visible = True
    cli_extra.Text = ""


    
If Err Then GrabarLog "cmdGenerar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdImprime_Comentarios_Click()
On Error Resume Next
  
    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox " Prepare la Impresora ...!", vbInformation, "Mensaje ..."
    
    With Mantenimiento.rsComentarios
        If .State = 0 Then .Open
        .Close
        .Open
    End With
    
    With drComentario
        .Sections(2).Controls("sfecha").Caption = vfecha.Value
        .Sections(2).Controls("sreparto").Caption = Reparto.Text
        .Sections(2).Controls("srepartidor").Caption = vrepartidor.Text
        .Show
    End With


If Err Then GrabarLog "cmdImprime_Comentarios_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdMostrarTodos_Click()
    On Error Resume Next
    
    With breparto
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM reparto ORDER BY reparto ASC"
        .Refresh
        
    End With
    
    plani.BackColor = &HFFFFFF
    
    If Err Then GrabarLog "cmdMostrarTodos_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdSalir_Click()
On Error Resume Next

    Unload Me

If Err Then GrabarLog "cmdSalir_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdUltimo_Click()
    On Error Resume Next
    
    With Mantenimiento.rsImprime_Reparto
        If Not .State = 1 Then .Open
        .Close
        .Open

        .Sort = "Dirección ASC, Codigo ASC"
    
    End With
    
    With bconfigura
        .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM configura"
        .Refresh
    End With
    
    'If breparto.Recordset.RecordCount <= 10 Then
    '    With drReparto2
    '        .Sections(2).Controls("Nom_Reparto").Caption = Trim(bconfigura.Recordset("ufreparto").Value)
    '        .Sections(2).Controls("Nom_repartidor").Caption = Trim(bconfigura.Recordset("urepartidor").Value)
    '        .Sections(2).Controls("sfecha_confeccion").Caption = Trim(bconfigura.Recordset("ufecha_reparto").Value)
    '        .Show
    '    End With
    'Else
    '    With drReparto
    '        .Sections(2).Controls("Nom_Reparto").Caption = Trim(bconfigura.Recordset("ufreparto").Value)
    '        .Sections(2).Controls("Nom_repartidor").Caption = Trim(bconfigura.Recordset("urepartidor").Value)
    '        .Sections(2).Controls("sfecha_confeccion").Caption = Trim(bconfigura.Recordset("ufecha_reparto").Value)
    '        .Show
    '    End With
    'End If
    
If Err Then GrabarLog "cmdUltimo_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Command8_Click()
    On Error Resume Next
    
    bgral.Recordset.MovePrevious
    varticulo = bgral.Recordset("descrip").Value
    vcod_articulo = bgral.Recordset("codigo").Value
 
If Err Then GrabarLog "Command8_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub Command9_Click()
    On Error Resume Next
    bgral.Recordset.MoveNext
    varticulo = bgral.Recordset("descrip").Value
    vcod_articulo = bgral.Recordset("codigo").Value

    If Err Then
        GrabarLog "Command9_Click", Err.Number & " " & Err.Description, Me.Name
        MsgBox "No se pudo completar la operacion-Command9", vbInformation, "Mensaje ..."
    End If

End Sub
Private Sub ControlarReparto()
    On Error Resume Next
    Dim vcont, vremito, vnremito As Long
    
    With bfdetalle
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM fdetalle WHERE (Fecha = '" & strfecha2(vfecha.Value) & "') ORDER BY remito ASC"
        .Refresh

        If Not .Recordset.RecordCount = 0 Then .Recordset.MoveFirst

        Do Until .Recordset.EOF
            vnremito = .Recordset("remito").Value

            If vremito = vnremito Then
                vremito = .Recordset("remito").Value
                vcont = vcont + 1
            Else
                vremito = .Recordset("remito").Value
                vcont = 1
            End If

            If vcont > 50 Then
                MsgBox "ERROR en la confección del reparto" & vbCrLf & "Por favor revise las operaciones y reahagalo", vbExclamation, "Mensaje ..."
                Exit Sub
            End If

            .Recordset.MoveNext
        Loop

    End With

    If Err Then GrabarLog "ControlarReparto", Err.Number & " " & Err.Description, Me.Name
End Sub
Function espacio_i(vchar As Variant, vespacio As Integer) As String
    Dim i As Integer

    If vchar = 0 Then
        espacio_i = Space(vespacio)
    Else
        espacio_i = vchar & Space(vespacio - Len(vchar))
    End If

End Function
Function BuscarComentario(ByRef vCodigo As String) As String
On Error Resume Next

    Dim connComentario As New ADODB.Connection
    Dim rsComentario As New ADODB.Recordset
    Dim sqlComentario As String
    
    With connComentario
        .ConnectionString = pathDBMySQL
        .Open
    End With
    
    sqlComentario = "SELECT * FROM comentarios WHERE (codigo = '" & vCodigo & "')"
    
    With rsComentario
        Call .Open(sqlComentario, connComentario, adOpenStatic, adLockPessimistic)
        
        If Not .EOF = True Then
            
            If .Fields("Repeticiones").Value > 0 Then
                BuscarComentario = Trim(.Fields("Comentario").Value)
            
                If .Fields("Repeticiones").Value > 0 Then
                    .Fields("Repeticiones").Value = .Fields("Repeticiones").Value - 1
                    .Update
                End If
            Else
                'En el caso que no tenga mas Repeticiones
                BuscarComentario = ""
            End If
        Else
            'En el caso que no tenga COMENTARIO
            BuscarComentario = ""
        End If
    
    End With
    
    sqlComentario = ""
    
    rsComentario.Close
    Set rsComentario = Nothing
    
    connComentario.Close
    Set connComentario = Nothing
    
If Err Then GrabarLog "BuscarComentario", Err.Number & " " & Err.Description, Caption
End Function

Function fdevol(vCodigo As String, vcod_cliente) As Integer

    '- calcula los envaces adedudados por este cliente --------------------
    '- utiliza una consulta q junta fdetalle (donde están todas las ventas y las devoluciones)-----

    On Error Resume Next

    With bdevol
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM devol WHERE Codigo = '" + vcod_cliente + "' and DetalleCodigo = '" + vCodigo + "'"
        .Refresh

        If Not .Recordset.EOF Then
            fdevol = .Recordset("sumaDecantidad") - .Recordset("sumaDeImpuesto1")
        Else
            fdevol = 0
        End If
    
    End With
    
    If Err Then GrabarLog Left("fdevol " & vCodigo & vcod_cliente, 49), Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub FiltrarComentarios(ByRef sql As String)
On Error Resume Next
    
    With bcomentario
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM comentarios WHERE 1=1"
        .Refresh
                    
        If Not .Recordset.EOF = True Then .Recordset.MoveFirst
            
    End With
    
    vccomentarios.Caption = bcomentario.Recordset.RecordCount

If Err Then GrabarLog "FiltrarComentarios", Err.Number & " " & Err.Description, Me.Name
End Sub

'--------------------------------------------------------------------
Private Sub fimpresion()
    Dim j, numpagina, vcantarticulos As Integer

    ' ----- seteo la impresora una sola vez para comenzar a imprimir  --------------------
    Printer.FontName = "Draft 10cpi"
    Printer.FontBold = False
    'Printer.PrintQuality = -1 '---ERROR ACA SALTA EN LA SEGUNDA VUELTA
    'Printer.Height = 17280
    '------------------------------------------------------------------------------------
    breparto.Refresh

    With bimprime_Reparto
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "Select * from imprime_reparto order by Dirección"
        .Refresh

        If Not .Recordset.RecordCount = 0 Then .Recordset.MoveFirst

        Do Until .Recordset.EOF
            numpagina = Val(numpagina) + 1
            
            'Encabezado de Hoja
            Printer.Print "Reparto: " & char_i(Reparto.Text, 61) & strfecha2(vfecha.Value)
            Printer.Print "--------------------------------------------------------------------------------"
            
            fimprimedetalle (breparto.Recordset.RecordCount)
            urenglon = urenglon + 2
            
            For j = urenglon To 82
                Printer.Print ""
            Next j
            
            'Número de Hoja
            Printer.Print "                                                                         " & num_i2(numpagina)
            Printer.Print ""
            Printer.Print ""
            urenglon = 0
          
        Loop

        Printer.EndDoc
    End With

End Sub

Private Sub fimprimedetalle(vDetalle As Integer)
    Dim i, j As Integer
    Dim vialias As String

    With bimprime_Reparto

        If vDetalle <= 10 Then

            For i = 1 To 9

                If bimprime_Reparto.Recordset.EOF = True Then
                    Exit For
                End If

                vialias = espacio_i(.Recordset("uc1"), 5) & espacio_i(.Recordset("e1"), 2) & espacio_i(.Recordset("uc2"), 5) & espacio_i(.Recordset("e2"), 2) & espacio_i(.Recordset("uc3"), 5) & espacio_i(.Recordset("e3"), 2) & espacio_i(.Recordset("uc4"), 5) & espacio_i(.Recordset("e4"), 2) & espacio_i(.Recordset("uc5"), 5) & espacio_i(.Recordset("e5"), 2) & espacio_i(.Recordset("uc6"), 5) & espacio_i(.Recordset("e6"), 2) & espacio_i(.Recordset("uc7"), 5) & espacio_i(.Recordset("e7"), 2) & espacio_i(.Recordset("uc8"), 5) & espacio_i(.Recordset("e8"), 2) & espacio_i(.Recordset("uc9"), 5) & espacio_i(.Recordset("e9"), 2) & espacio_i(.Recordset("uc10"), 5) & espacio_i(.Recordset("e10"), 1)
            
                Printer.Print char_i(Trim(Str(.Recordset("codigo"))), 5) + "- " + char_i(.Recordset("Nombre"), 30) + "U. PAGO" + "   SALDO" + "  CREDITO" + "   S.ACTUAL" + "   F | C"
                Printer.Print char_i(.Recordset("dirección"), 38) & Day(Val(Format(.Recordset("upago"), "######0"))) & "/" & Month(Val(Format(.Recordset("upago"), "######0"))) & "    " & Format(num_i2(.Recordset("saldoanterior")), "#####0.00") + "     20.00" + num_i(.Recordset("saldo"))
                Printer.Print "U.COMPRA"
                Printer.Print espacio_i(Format(.Recordset("ucompra"), "dd/mm/yy"), 11) & vialias
                vialias = ""
                Printer.Print espacio_i("Pago", 11) & char_i(.Recordset("a1"), 6) & " " & char_i(.Recordset("a2"), 6) & " " & char_i(.Recordset("a3"), 6) & " " & char_i(.Recordset("a4"), 6) & " " & char_i(.Recordset("a5"), 6) & " " & char_i(.Recordset("a6"), 6) & " " & char_i(.Recordset("a7"), 6) & " " & char_i(.Recordset("a8"), 6) & " " & char_i(.Recordset("a9"), 6) & " " & char_i(.Recordset("a10"), 6)
                Printer.Print ""
                Printer.Print "Comentario:" & .Recordset("Comentarios")
                Printer.Print "--------------------------------------------------------------------------------"
                bimprime_Reparto.Recordset.MoveNext

                urenglon = urenglon + 8
            Next i

            For j = 0 To 7
                Printer.Print ""
                urenglon = urenglon + 1
            Next j

        Else

            For i = 1 To 8

                If bimprime_Reparto.Recordset.EOF = True Then
                    Exit For
                End If

                vialias = espacio_i(.Recordset("uc1"), 5) & espacio_i(.Recordset("e1"), 2) & espacio_i(.Recordset("uc2"), 5) & espacio_i(.Recordset("e2"), 2) & espacio_i(.Recordset("uc3"), 5) & espacio_i(.Recordset("e3"), 2) & espacio_i(.Recordset("uc4"), 5) & espacio_i(.Recordset("e4"), 2) & espacio_i(.Recordset("uc5"), 5) & espacio_i(.Recordset("e5"), 2) & espacio_i(.Recordset("uc6"), 5) & espacio_i(.Recordset("e6"), 2) & espacio_i(.Recordset("uc7"), 5) & espacio_i(.Recordset("e7"), 2) & espacio_i(.Recordset("uc8"), 5) & espacio_i(.Recordset("e8"), 2) & espacio_i(.Recordset("uc9"), 5) & espacio_i(.Recordset("e9"), 2) & espacio_i(.Recordset("uc10"), 5) '& .Recordset("e10")
                Printer.Print char_i(.Recordset("codigo"), 5) + "- " + char_i(.Recordset("Nombre"), 30) + "U. PAGO" + "   SALDO" + "  CREDITO" + "   S.ACTUAL" + "   F | C"
                Printer.Print char_i(.Recordset("dirección"), 38) & Day(.Recordset("upago")) & "/" & Month(.Recordset("upago")) & "    " & Format(num_i2(.Recordset("saldoanterior")), "#####0.00") + "     20.00" + num_i(.Recordset("saldo"))
                Printer.Print "U.COMPRA"
                Printer.Print espacio_i(Format(.Recordset("ucompra"), "dd/mm/yy"), 11) & vialias
                vialias = ""
                Printer.Print espacio_i("Pago", 11) & char_i(.Recordset("a1"), 6) & " " & char_i(.Recordset("a2"), 6) & " " & char_i(.Recordset("a3"), 6) & " " & char_i(.Recordset("a4"), 6) & " " & char_i(.Recordset("a5"), 6) & " " & char_i(.Recordset("a6"), 6) & " " & char_i(.Recordset("a7"), 6) & " " & char_i(.Recordset("a8"), 6) & " " & char_i(.Recordset("a9"), 6) & " " & char_i(.Recordset("a10"), 6)
                vialias = espacio_i(.Recordset("uc11"), 5) & espacio_i(.Recordset("e11"), 2) & espacio_i(.Recordset("uc12"), 5) & espacio_i(.Recordset("e12"), 2) & espacio_i(.Recordset("uc13"), 5) & espacio_i(.Recordset("e13"), 2) & espacio_i(.Recordset("uc14"), 5) & espacio_i(.Recordset("e14"), 2) & espacio_i(.Recordset("uc15"), 5) & espacio_i(.Recordset("e15"), 2) & espacio_i(.Recordset("uc16"), 5) & espacio_i(.Recordset("e16"), 2) & espacio_i(.Recordset("uc17"), 5) & espacio_i(.Recordset("e17"), 2) & espacio_i(.Recordset("uc18"), 5) & espacio_i(.Recordset("e18"), 2) & espacio_i(.Recordset("uc19"), 5) & espacio_i(.Recordset("e19"), 2) & espacio_i(.Recordset("uc20"), 5) & .Recordset("e20")
                Printer.Print Space(11) & vialias
                vialias = ""
                Printer.Print Space(11) & char_i(.Recordset("a11"), 6) & " " & char_i(.Recordset("a12"), 6) & " " & char_i(.Recordset("a13"), 6) & " " & char_i(.Recordset("a14"), 6) & " " & char_i(.Recordset("a15"), 6) & " " & char_i(.Recordset("a16"), 6) & " " & char_i(.Recordset("a17"), 6) & " " & char_i(.Recordset("a18"), 6) & " " & char_i(.Recordset("a19"), 6) & " " & char_i(.Recordset("a20"), 6)
                Printer.Print ""
                Printer.Print "Comentario:" & .Recordset("Comentarios")
                Printer.Print "--------------------------------------------------------------------------------"
                bimprime_Reparto.Recordset.MoveNext
                urenglon = urenglon + 10
            Next i

        End If

    End With

End Sub

Function fiuc(vCodigo As String) As Double
    On Error Resume Next

    With bultima_compra
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "select * from Ultima_Compra_mod where Codigo_Cliente = '" + vCodigo + "' order by Fecha ASC"
        .Refresh
    
        If .Recordset.EOF = True Then
            fiuc = 0
        Else
            fiuc = .Recordset("últimodetotal")
        End If
        
    End With
    
    If Err Then GrabarLog Left("fiuc: " & vCodigo, 49), Err.Number & " " & Err.Description, Me.Name
End Function


Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        If MsgBox("¿Esta seguro que desea parar/detener este proceso?", vbInformation + vbYesNo, "Mensaje ...") = vbYes Then
            vEscape = True
        End If
    End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Load()
    On Error Resume Next
   
   With bpagopormes
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM pagopormes"
        .Refresh
    End With
    
   
    With bfdetalle
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM fdetalle"
        .Refresh
    End With
    
    With Me
        .Height = 6375
        .Width = 12270
        .Left = 30
        .Top = 30
    End With
    
    vfecha.Value = Date
    KeyPreview = True
    vrepcompleto = 0
      
    cmdCarga_Comentario.Enabled = True
      
    If Err Then GrabarLog "Form_load", Err.Number & " " & Err.Description, Me.Name
    
End Sub

Function fprecio_articulo(vcod_articulo As String, vcod_cliente As String) As Double

    On Error Resume Next
    ' ir a mirar en la tabla donde esté relacionada un cliente a un artículo y ver el precio. En el caso q no exista, poner el precio q tiene en la lista
    With barticulos_Clientes
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "select * from articulos_clientes where codigo_cliente = '" + vcod_cliente + "' and articulo = '" + vcod_articulo + "'"
        .Refresh
    
        If Not .Recordset.EOF Then
            fprecio_articulo = .Recordset("precio").Value
        Else
            barticulo.Refresh
            barticulo.Recordset.Find ("codigo = '" + vcod_articulo + "'")
            fprecio_articulo = barticulo.Recordset("pventa1").Value
        End If
    End With
    If Err Then GrabarLog Left("fprecio_articulo " & vcod_articulo & vcod_cliente, 49), Err.Number & " " & Err.Description, Me.Name
End Function
Function fremito() As Long
On Error Resume Next
                
    With bfactura
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM factura"
        .Refresh
    
        If .Recordset.EOF = True Then
            
            nremito = 1
            fremito = 1
        
        Else
        
            .Recordset.Sort = "remito"
            .Recordset.MoveLast
            nremito = .Recordset("remito") + 1
            fremito = .Recordset("remito") + 1
        
        End If
    
    End With
    
If Err Then GrabarLog "fremito", Err.Number & " " & Err.Description, Me.Name
End Function

Private Sub freparto(vCodigo As String, _
                     vnombre As String, _
                     vdireccion As String, _
                     vcredito As Double)
    
    On Error Resume Next
    Dim i As Integer
    Dim vfup As String
    Dim vvsaldoanterior, vvsaldoactual  As Double
    
    breparto.Refresh
    i = 0
    
    bimprime_Reparto.Recordset.AddNew
    
    Dim viuc As Double
    
    barra2.Max = breparto.Recordset.RecordCount
    barra2.Value = 0
    Do Until breparto.Recordset.EOF = True
        
        DoEvents
        i = i + 1

        
        If i <= 20 Then
            'Exit Sub
            
            'Para no guardar "i - 1" veces en los campos que no son necesarios
            'If Not (vcodigo = vcodigoant) Then
            barra2.Value = barra2.Value + 1
            
            'If Not i > 1 Then ' estos datos son siempre lo mismos en el registro.
               
            'End If
            'bimprime_Reparto.Recordset("credito") = vcredito
            
            
            viuc = fiuc(vCodigo)
            'End If
            'bimprime_Reparto.Recordset("comentario") = fcc(vcodigo) ' fun q busca el comentario en la tabla cliente
            
            bimprime_Reparto.Recordset("a" + Trim(Str(i))) = breparto.Recordset("alias") 'nombre del artículo
            bimprime_Reparto.Recordset("e" + Trim(Str(i))) = fdevol(breparto.Recordset("Cod_Articulo"), vCodigo) ' envases en mora
            bimprime_Reparto.Recordset("uc" + Trim(Str(i))) = fuc(breparto.Recordset("Cod_Articulo"), vCodigo) ' ultima compra
        
            If Not vfecha_uc = "" Then bimprime_Reparto.Recordset("ucompra") = vfecha_uc 'Format(vfecha_uc)
            vfup = fup(vCodigo)
            
            If Not vfup = "" Then bimprime_Reparto.Recordset("upago") = vfup 'Format(fup(vcodigo))
            
           If Not i > 1 Then
                
                bimprime_Reparto.Recordset("codigo").Value = vCodigo
                bimprime_Reparto.Recordset("dirección").Value = vdireccion
                bimprime_Reparto.Recordset("nombre").Value = vnombre
                bimprime_Reparto.Recordset("saldo").Value = fsaldo(vCodigo)
                'bimprime_Reparto.Recordset("saldoAnterior") = Format(calsaldoanterior(date - Left(date, 2), vcodigo), "######0.00")
                
                
               vvsaldoanterior = (fsaldoAnterior(vCodigo))
               
                bimprime_Reparto.Recordset("saldoAnterior").Value = vvsaldoanterior
                bimprime_Reparto.Recordset("comentarios").Value = BuscarComentario(vCodigo)
                
                '-----------Procedimiento para guardar los comentarios------------
                
                '-----------------------------------------------------------------
        
        
            End If
        End If
        
        If chkfact.Value = 1 Then
            With barticulo
                If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
                .RecordSource = "SELECT * FROM Articulos WHERE (codigo = '" & Trim(breparto.Recordset("cod_articulo").Value) + "')"
                .Refresh
            End With
                
            ' ------- busco la ganancia que tiene el artículo ----------------------
            With barticulos_Ganancia
                .RecordSource = "SELECT * FROM Articulos_Ganancia WHERE (CodEmp = '" & vcod_repartidor & "') and (CodCli = '" + vCodigo + "') and (Codrub =   '" & Trim(barticulo.Recordset("rubro").Value) & "'"
                .Refresh
                            
                If Not .Recordset.EOF = True Then
                    'En el caso que tenga asignado porcentaje
                    vganancia = .Recordset("porcentaje").Value
                    
                Else
                    ' en el caso q no tengo asignado rubro, cliente, empleado porcentaje
                    vganancia = Val(Format(barticulo.Recordset("ganancia_vendedor"), "#######0.00"))
                    
                End If
            
            End With
            '-----------------------------------------------------------------------
                        
            ' ----- en este lugar se graba un registro en la tabla fdetalle -------------
        
            
            With bfdetalle
                .Recordset.AddNew
                .Recordset("fecha") = vfecha.Value
                .Recordset("detalle") = breparto.Recordset("Articulo").Value
                .Recordset("codigo") = barticulo.Recordset("codigo").Value
                .Recordset("precio") = fprecio_articulo(breparto.Recordset("cod_Articulo").Value, vCodigo)
                .Recordset("remito") = fremito  ' me tiene q tirar el numero del ultimo remito + 1 (ojo q este es el mismo en este bucle)
                .Recordset("Ganancia") = vganancia
            
                'bfdetalle.Recordset("sueldo") = (vganancia * bfdetalle.Recordset("precio")) / 100 'Se calcula en la factura
            
                .Recordset.Update
            End With
            '----------------------------------------------------------------------------
        End If

        breparto.Recordset.MoveNext
    Loop

    bimprime_Reparto.Recordset.Update
   
   vclirep.Caption = bimprime_Reparto.Recordset.RecordCount
   vclirep.Refresh
   
    'vfcliente = ""
    ' ---------  en este lugar tengo q guardar un registro en la tabla factura ------------
    If chkfact.Value = 1 Then
        
        With bfactura

            
            .Recordset.AddNew
            .Recordset("fecha") = vfecha
            .Recordset("remito") = nremito
            .Recordset("nombre") = vnombre
            .Recordset("Domicilio") = vdireccion
            .Recordset("codigo") = vCodigo
            .Recordset("iva") = tipoiva(vCodigo)
            .Recordset("cod_repartidor") = vcod_repartidor   'bclientes_repartos.Recordset("cod_empleado")
            .Recordset("repartidor") = vrepartidor.Text
            .Recordset("localidad") = vlocalidad.Text
        
                 
            .Recordset.Update
        End With
            
    End If

    If Err Then
        GrabarLog Left("freparto " & vCodigo & vnombre & vdireccion, 49), Err.Number & " " & Err.Description, Me.Name
        problemas.AddItem (vCodigo)
        problemas.Refresh
    End If
End Sub
Function fsaldo(vCodigo As String) As Double
    On Error Resume Next
    
    With bsaldos_clientes
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "select * from saldos_clientes where codigo = '" + vCodigo + "'"
        .Refresh
        
        If .Recordset.EOF = True Then
        
        
            fsaldo = 0
        Else
            fsaldo = .Recordset(1) - .Recordset(2)
        End If
    
    End With
    If Err Then GrabarLog Left("fsaldo " & vCodigo, 49), Err.Number & " " & Err.Description, Me.Name
End Function

Function fsaldoAnterior(vCodigo As String) As Double ' saldo anterior calculada con pagopormes ''panic panic
Dim vimporte As Double
vimporte = 0
'bpagopormes.RecordSource = "select * from pagopormes where codigo = '" + vcodigo + "'"
 bpagopormes.RecordSource = "SELECT  Sum(fdetalle.resta) AS resta, cuentascorrientes.anomes, Max(cuentascorrientes.Saldo_PPM) AS ÚltimoDeSaldo_PPM, cuentascorrientes.Codigo, cuentascorrientes.Noimputar FROM cuentascorrientes INNER JOIN (Factura INNER JOIN fdetalle ON Factura.Remito = fdetalle.Remito) ON cuentascorrientes.Remito = fdetalle.Remito GROUP BY cuentascorrientes.anomes, cuentascorrientes.Codigo, cuentascorrientes.Noimputar HAVING (((cuentascorrientes.Codigo)='" + Trim(vCodigo) + "') AND ((cuentascorrientes.Noimputar)=0));"
 bpagopormes.Refresh

If bpagopormes.Recordset.EOF Then


Else
    problemas.AddItem ("Codigo: " + Str(vCodigo))
Do Until Val(bpagopormes.Recordset("anomes")) >= Val((Trim(Right(Date, 4) + Trim(Left(strfechaMySQL(Date), 2)))))
    problemas.AddItem (Trim(bpagopormes.Recordset("anomes")) + " " + Str(vimporte))
    vimporte = bpagopormes.Recordset("resta") + vimporte
bpagopormes.Recordset.MoveNext

If bpagopormes.Recordset.EOF Then
    fsaldoAnterior = vimporte
Exit Function
End If
Loop

End If

fsaldoAnterior = vimporte


End Function
Function fuc(vcod_art As String, vcod_cli As String)
    Dim vultimaf As Date
    '--------- verifica un articulo de un cliente; si c encuentra en la última compra realizada por el cliente. ---------

    On Error Resume Next
    'bultima_compra.Refresh

    If bultima_compra.Recordset.EOF Then
        vfecha_uc = ""
        Exit Function
    End If

    'bultima_compra.RecordSource = "Select * From Ultima_Compra where Codigo = '" + vcod_cli + "' order by UFecha ASC"
    
    bultima_compra.RecordSource = "Select * From Ultima_Compra_mod where Codigo_Cliente = '" + vcod_cli + "' order by Fecha ASC"
    bultima_compra.Refresh

    If bultima_compra.Recordset.EOF Then
        fuc = 0
        Exit Function
    End If
    bultima_compra.Recordset.MoveLast ' me voy a la ultima compra del cliente
    
    vfecha_uc = bultima_compra.Recordset("Fecha") ' guardo la fecha de la ultima compra en una variable global
    vultimaf = Format(vfecha_uc)
                
    'bultima_compra.RecordSource = "select * from ultima_compra where factura.codigo = '" + Trim(vcod_cli) + "' and ÚltimoDeCodigo1 = '" + Trim(vcod_art) + "' and Ufecha = '" & strfechaMySQl(vultimaf) + "'"
    
    'ULTIMO UPDATE
    'bultima_compra.RecordSource = "select * from ultima_compra where factura.codigo = '" + Trim(vcod_cli) + "' and Ufecha = '" & strfechaMySQl(vultimaf) + "'"
    
    bultima_compra.RecordSource = "SELECT * FROM Ultima_Compra_mod WHERE Codigo_Cliente = '" + Trim(vcod_cli) + "' AND fecha = '" & strfechaMySQL(vultimaf) + "'"
    bultima_compra.Refresh
    
    'bultima_compra.Recordset.Find "ÚltimoDeCodigo1 = '" + Trim(vcod_art) + "'"
    bultima_compra.Recordset.Find "Codigo_Articulos = '" + Trim(vcod_art) + "'"
    
    If Not bultima_compra.Recordset.EOF Then
        fuc = bultima_compra.Recordset("últimodecantidad")
    Else
        fuc = 0
    End If

    If Err Then GrabarLog Left("fuc " & vcod_art & vcod_cli, 49), Err.Number & " " & Err.Description, Me.Name
End Function

Function fup(vcod_cliente As String) As String 'Funcion para el ultimo pago

    With bccliente
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "Select * from cuentascorrientes where codigo = '" + vcod_cliente + "' and NoImputar = True order by Fecha"
        .Refresh
        If Not .Recordset.RecordCount = 0 Then
            .Recordset.MoveLast
            fup = .Recordset("fecha")
        Else
            fup = ""
        End If

    End With

End Function

Private Sub Limpiar()
    On Error Resume Next
    varticulo = ""
    valias = ""
    varticulo.SetFocus

    If Err Then GrabarLog "Limpiar", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub plani_BeforeColUpdate(ByVal ColIndex As Integer, _
                                  OldValue As Variant, _
                                  Cancel As Integer)
    breparto.Recordset.Update
End Sub

Private Sub Reparto_GotFocus()
    On Error Resume Next
    
    CargarCombo "clireparto", "descrip", Reparto, False
    
    If Err Then GrabarLog "Reparto_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub Reparto_KeyPress(KeyAscii As Integer)
    On Error Resume Next

    If KeyAscii = 13 Then vcod_reparto.SetFocus

    If Err Then GrabarLog "Reparto_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub Reparto_LostFocus()
    On Error Resume Next
    
    With bclirepartos
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM clireparto ORDER BY descrip"
        .Refresh
        
        .Recordset.Find ("descrip = '" + Trim(Reparto.Text) + "'")

        If .Recordset.EOF Then
            MsgBox "Usted ha ingresado un reparto inexistente." + Chr(13) + "Debe ingresarlo desde el mantenimiento de reparto", vbCritical, "Error..."
            Reparto.SetFocus
        Else
            vcod_reparto.Text = .Recordset("nreparto")
        End If
    
    End With
    
    If Err Then GrabarLog "Reparto_LostFocus", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next

    If dibu.Visible Then
        dibu.Visible = False
    Else
        dibu.Visible = True
    End If

    If Err Then GrabarLog "Timer1_Timer", Err.Number & " " & Err.Description, Me.Name
End Sub

Function tipoiva(vCodigo As String) As String
    On Error Resume Next
    With bcliente
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM clientes"
        .Refresh
        
        
        .Recordset.Find ("codigo = '" + vCodigo + "'")
        tipoiva = .Recordset("iva")
    
    End With

    If Err Then GrabarLog Left("tipoiva " & vCodigo, 49), Err.Number & " " & Err.Description, Me.Name
End Function

Private Sub valias_KeyPress(KeyAscii As Integer)
    On Error Resume Next

    If KeyAscii = 13 Then
        cmdConfirmar.SetFocus
    End If

    If Err Then GrabarLog "valias_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub varticulo_keypress(KeyAscii As Integer)
    On Error Resume Next

    If KeyAscii = 13 Then
    
        With bgral
            If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
            .RecordSource = "select * from articulos where codigo = '" + varticulo.Text + "'"  'or descrip like '%" + varticulo.Text + "%'"
            .Refresh
    
            If Not .Recordset.EOF Then
                varticulo.Text = .Recordset("descrip")
                vcod_articulo = .Recordset("codigo")
            End If
        
        End With
        
        valias.SetFocus
        valias.SelStart = 0
        valias.SelLength = Len(valias)
    End If

    If Err Then GrabarLog "varticulo_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub vcod_reparto_KeyPress(KeyAscii As Integer)
    On Error Resume Next

    If KeyAscii = 13 Then
        vlocalidad.SetFocus
    End If

    If Err Then GrabarLog "vcod_reparto_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub vlocalidad_KeyPress(KeyAscii As Integer)
    On Error Resume Next

    If KeyAscii = 13 Then
        vrepartidor.SetFocus
    End If

    If Err Then GrabarLog "vlocalidad_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub vrepartidor_KeyPress(KeyAscii As Integer)
    On Error Resume Next

    If KeyAscii = 13 Then
        
        With bgral
            If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
            .RecordSource = "select * from empleados where codigo like '%" + vrepartidor.Text + "%' or nombre like '%" + vrepartidor.Text + "%'"
            .Refresh
    
            If Not .Recordset.EOF Then
                vrepartidor.Text = .Recordset("nombre")
                vcod_repartidor = .Recordset("codigo")
            Else
                frmEmpleados.Show
                frmEmpleados.TabEmpleados.Tab = 1
            End If
        End With
        varticulo.SetFocus
    End If

    If Err Then GrabarLog "vrepartidor_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub



