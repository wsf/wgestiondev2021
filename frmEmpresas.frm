VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmEmpresas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administracion de Empresas"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.PushButton cmdAcciones 
      Height          =   420
      Index           =   0
      Left            =   2400
      TabIndex        =   8
      Top             =   4440
      Width           =   1815
      _Version        =   851968
      _ExtentX        =   3201
      _ExtentY        =   732
      _StockProps     =   79
      Caption         =   "Aceptar"
      BackColor       =   -2147483638
      Appearance      =   6
      Picture         =   "frmEmpresas.frx":0000
   End
   Begin XtremeSuiteControls.PushButton cmdAcciones 
      Height          =   420
      Index           =   1
      Left            =   4320
      TabIndex        =   9
      Top             =   4440
      Width           =   1815
      _Version        =   851968
      _ExtentX        =   3201
      _ExtentY        =   732
      _StockProps     =   79
      Caption         =   "Volver/Salir"
      BackColor       =   -2147483638
      Appearance      =   6
      Picture         =   "frmEmpresas.frx":6862
   End
   Begin XtremeSuiteControls.FlatEdit txtDatosUsuario 
      Height          =   345
      Index           =   0
      Left            =   2250
      TabIndex        =   0
      Top             =   440
      Width           =   3855
      _Version        =   851968
      _ExtentX        =   6800
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   -2147483643
      Appearance      =   4
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDatos 
      Height          =   345
      Index           =   1
      Left            =   2250
      TabIndex        =   1
      Top             =   920
      Width           =   3855
      _Version        =   851968
      _ExtentX        =   6800
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   -2147483643
      MaxLength       =   8
      Appearance      =   4
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDatos 
      Height          =   345
      Index           =   3
      Left            =   2250
      TabIndex        =   3
      Top             =   1880
      Width           =   3855
      _Version        =   851968
      _ExtentX        =   6800
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   -2147483643
      Appearance      =   4
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDatos 
      Height          =   345
      Index           =   4
      Left            =   2250
      TabIndex        =   4
      Top             =   2360
      Width           =   3855
      _Version        =   851968
      _ExtentX        =   6800
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   -2147483643
      Appearance      =   4
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDatos 
      Height          =   345
      Index           =   5
      Left            =   2250
      TabIndex        =   5
      Top             =   2840
      Width           =   3855
      _Version        =   851968
      _ExtentX        =   6800
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   -2147483643
      Appearance      =   4
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDatos 
      Height          =   345
      Index           =   6
      Left            =   2250
      TabIndex        =   6
      Top             =   3320
      Width           =   3855
      _Version        =   851968
      _ExtentX        =   6800
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   -2147483643
      Appearance      =   4
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDatos 
      Height          =   345
      Index           =   7
      Left            =   2250
      TabIndex        =   7
      Top             =   3800
      Width           =   3855
      _Version        =   851968
      _ExtentX        =   6800
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   -2147483643
      Appearance      =   4
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDatos 
      Height          =   345
      Index           =   2
      Left            =   2250
      TabIndex        =   2
      Top             =   1400
      Width           =   3855
      _Version        =   851968
      _ExtentX        =   6800
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   -2147483643
      Appearance      =   4
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit lblTitulo 
      Height          =   345
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   6255
      _Version        =   851968
      _ExtentX        =   11033
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Text            =   "Ingrese los datos de la Nueva Empresa"
      BackColor       =   12648447
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   3
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label lblDatosServidor 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "> ID del Servidor  :"
      Height          =   195
      Index           =   8
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   2205
   End
   Begin VB.Label lblDatosServidor 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "> Nombre de la Servidor  :"
      Height          =   195
      Index           =   7
      Left            =   0
      TabIndex        =   17
      Top             =   435
      Width           =   2205
   End
   Begin VB.Label lblDatosServidor 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "> Direccion del Servidor  :"
      Height          =   195
      Index           =   6
      Left            =   0
      TabIndex        =   16
      Top             =   915
      Width           =   2205
   End
   Begin VB.Label lblDatos 
      Alignment       =   1  'Right Justify
      Caption         =   "> Website :"
      Height          =   195
      Index           =   7
      Left            =   30
      TabIndex        =   14
      Top             =   3840
      Width           =   2205
   End
   Begin VB.Label lblDatos 
      Alignment       =   1  'Right Justify
      Caption         =   "> E-Mail :"
      Height          =   195
      Index           =   6
      Left            =   30
      TabIndex        =   13
      Top             =   3360
      Width           =   2205
   End
   Begin VB.Label lblDatos 
      Alignment       =   1  'Right Justify
      Caption         =   "> C.U.I.T :"
      Height          =   195
      Index           =   2
      Left            =   0
      TabIndex        =   12
      Top             =   1440
      Width           =   2200
   End
   Begin VB.Label lblDatos 
      Alignment       =   1  'Right Justify
      Caption         =   "> Alias de la Empresa  :"
      Height          =   195
      Index           =   1
      Left            =   30
      TabIndex        =   11
      Top             =   960
      Width           =   2200
   End
   Begin VB.Label lblDatosUsuario 
      Alignment       =   1  'Right Justify
      Caption         =   "> Nombre de la Empresa  :"
      Height          =   195
      Index           =   0
      Left            =   30
      TabIndex        =   10
      Top             =   480
      Width           =   2200
   End
End
Attribute VB_Name = "frmEmpresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

