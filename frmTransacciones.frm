VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmTransacciones 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   9300
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   825
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      _Version        =   851968
      _ExtentX        =   16325
      _ExtentY        =   1455
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.FlatEdit vnrointerno 
         Height          =   285
         Left            =   1980
         TabIndex        =   2
         Top             =   300
         Width           =   6735
         _Version        =   851968
         _ExtentX        =   11880
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin VB.Label Label1 
         Caption         =   "Nro Interno:"
         Height          =   225
         Left            =   300
         TabIndex        =   1
         Top             =   330
         Width           =   1335
      End
   End
   Begin XtremeSuiteControls.PushButton cmdFiltroMovimientos 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   9285
      _Version        =   851968
      _ExtentX        =   16378
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Buscar la transacción asociada"
      Appearance      =   6
      Picture         =   "frmTransacciones.frx":0000
   End
End
Attribute VB_Name = "frmTransacciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
