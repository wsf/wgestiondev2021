VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.ShortcutBar.v13.0.0.Demo.ocx"
Object = "{9746E3DA-06E1-4D26-9CE4-D9F6411A9C70}#1.0#0"; "SMGA_OcxTxt2009.ocx"
Begin VB.Form frmBancosMovimientosAlta 
   Caption         =   "Alta de moviemntcoo de Caja/Banco. "
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   2175
      Left            =   0
      TabIndex        =   6
      Top             =   420
      Width           =   7545
      _Version        =   851968
      _ExtentX        =   13309
      _ExtentY        =   3836
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin Aplisoft_CajasDeTexto.TxF vfecha 
         Height          =   285
         Left            =   1020
         TabIndex        =   7
         Top             =   240
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label lblFecha 
         Height          =   165
         Left            =   150
         TabIndex        =   8
         Top             =   330
         Width           =   735
         _Version        =   851968
         _ExtentX        =   1296
         _ExtentY        =   291
         _StockProps     =   79
         Caption         =   "> Fecha:"
         Alignment       =   1
      End
   End
   Begin VB.PictureBox PicInferior 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      Picture         =   "frmBancosMovimientosAlta.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   7590
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2640
      Width           =   7590
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   0
         Left            =   5340
         TabIndex        =   2
         Top             =   120
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Aplicar"
         Appearance      =   6
         Picture         =   "frmBancosMovimientosAlta.frx":50B3
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   1
         Left            =   6450
         TabIndex        =   3
         Top             =   120
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Cerrar"
         Appearance      =   6
         Picture         =   "frmBancosMovimientosAlta.frx":54BA
      End
      Begin VB.Label lblWGestion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WGESTION 2010"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Index           =   0
         Left            =   50
         TabIndex        =   5
         Top             =   150
         Width           =   1770
      End
      Begin VB.Label lblWGestion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WGESTION 2010"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Index           =   1
         Left            =   75
         TabIndex        =   4
         Top             =   170
         Width           =   1770
      End
   End
   Begin XtremeShortcutBar.ShortcutCaption vcajabanco 
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7545
      _Version        =   851968
      _ExtentX        =   13309
      _ExtentY        =   767
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmBancosMovimientosAlta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
