VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{52463EDA-D668-43B6-8D47-4FA8035EF04A}#1.0#0"; "PhotoWSF.ocx"
Begin VB.Form frmProveedoresAlta 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alta de Proveedores"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.ComboBox vtipoProveedor 
      Height          =   315
      Left            =   4980
      TabIndex        =   80
      Top             =   60
      Width           =   4005
      _Version        =   851968
      _ExtentX        =   7064
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "Proveedor"
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   315
      Left            =   2220
      TabIndex        =   79
      Top             =   1200
      Width           =   315
      _Version        =   851968
      _ExtentX        =   556
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "..."
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit vpersonas 
      Height          =   345
      Left            =   2580
      TabIndex        =   78
      Top             =   1200
      Width           =   6405
      _Version        =   851968
      _ExtentX        =   11298
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   -2147483643
   End
   Begin PhotoWSF.Photo phtProveedor 
      Height          =   1515
      Left            =   9090
      TabIndex        =   68
      Top             =   30
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2672
      BorderColor     =   14737632
      BorderColor     =   14737632
      BackStyle       =   0
   End
   Begin XtremeSuiteControls.PushButton pbCarga 
      Height          =   315
      Index           =   0
      Left            =   10590
      TabIndex        =   30
      Top             =   90
      Width           =   315
      _Version        =   851968
      _ExtentX        =   556
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "..."
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton pbCierraFoto 
      Height          =   315
      Left            =   10590
      TabIndex        =   32
      Top             =   420
      Visible         =   0   'False
      Width           =   315
      _Version        =   851968
      _ExtentX        =   556
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "X"
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtAlta 
      Height          =   315
      Index           =   1
      Left            =   2580
      TabIndex        =   1
      Top             =   450
      Width           =   6405
      _Version        =   851968
      _ExtentX        =   11298
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit txtAlta 
      Height          =   315
      Index           =   2
      Left            =   2580
      TabIndex        =   2
      Top             =   810
      Width           =   6405
      _Version        =   851968
      _ExtentX        =   11298
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.TabControl TabAlta 
      Height          =   3615
      Left            =   0
      TabIndex        =   36
      Top             =   1620
      Width           =   10965
      _Version        =   851968
      _ExtentX        =   19341
      _ExtentY        =   6376
      _StockProps     =   68
      Color           =   8
      ItemCount       =   4
      Item(0).Caption =   "Ficha"
      Item(0).ControlCount=   30
      Item(0).Control(0)=   "txtFicha(0)"
      Item(0).Control(1)=   "lblFicha(0)"
      Item(0).Control(2)=   "lblFicha(1)"
      Item(0).Control(3)=   "txtFicha(1)"
      Item(0).Control(4)=   "lblFicha(2)"
      Item(0).Control(5)=   "txtFicha(2)"
      Item(0).Control(6)=   "txtFicha(3)"
      Item(0).Control(7)=   "lblFicha(3)"
      Item(0).Control(8)=   "txtFicha(4)"
      Item(0).Control(9)=   "lblFicha(4)"
      Item(0).Control(10)=   "lblFicha(5)"
      Item(0).Control(11)=   "lblFicha(6)"
      Item(0).Control(12)=   "lblFicha(7)"
      Item(0).Control(13)=   "txtFicha(5)"
      Item(0).Control(14)=   "pbCarga(1)"
      Item(0).Control(15)=   "txtFicha(6)"
      Item(0).Control(16)=   "txtFicha(7)"
      Item(0).Control(17)=   "lblFicha(8)"
      Item(0).Control(18)=   "txtFicha(8)"
      Item(0).Control(19)=   "txtFicha(9)"
      Item(0).Control(20)=   "pbCarga(2)"
      Item(0).Control(21)=   "txtFicha(10)"
      Item(0).Control(22)=   "txtFicha(11)"
      Item(0).Control(23)=   "lblFicha(9)"
      Item(0).Control(24)=   "pbCarga(3)"
      Item(0).Control(25)=   "vporcentajeVendedor"
      Item(0).Control(26)=   "lblFicha(10)"
      Item(0).Control(27)=   "Pus"
      Item(0).Control(28)=   "vdesvendedor"
      Item(0).Control(29)=   "vcodvendedor"
      Item(1).Caption =   "Datos Comerciales"
      Item(1).ControlCount=   20
      Item(1).Control(0)=   "lblDatosComerciales(1)"
      Item(1).Control(1)=   "lblDatosComerciales(0)"
      Item(1).Control(2)=   "lblDatosComerciales(2)"
      Item(1).Control(3)=   "lblDatosComerciales(3)"
      Item(1).Control(4)=   "lblDatosComerciales(4)"
      Item(1).Control(5)=   "txtDatosComerciales(0)"
      Item(1).Control(6)=   "txtDatosComerciales(1)"
      Item(1).Control(7)=   "txtDatosComerciales(2)"
      Item(1).Control(8)=   "txtDatosComerciales(3)"
      Item(1).Control(9)=   "txtDatosComerciales(4)"
      Item(1).Control(10)=   "pbCarga(4)"
      Item(1).Control(11)=   "txtDatosComerciales(5)"
      Item(1).Control(12)=   "pbCarga(5)"
      Item(1).Control(13)=   "txtDatosComerciales(6)"
      Item(1).Control(14)=   "txtDatosComerciales(7)"
      Item(1).Control(15)=   "pbCarga(6)"
      Item(1).Control(16)=   "txtDatosComerciales(8)"
      Item(1).Control(17)=   "pbCarga(7)"
      Item(1).Control(18)=   "lblDatosComerciales(5)"
      Item(1).Control(19)=   "txtDatosComerciales(9)"
      Item(2).Caption =   "Otros Datos"
      Item(2).ControlCount=   16
      Item(2).Control(0)=   "lblOtrosDatos(0)"
      Item(2).Control(1)=   "dtpAlta"
      Item(2).Control(2)=   "lblOtrosDatos(1)"
      Item(2).Control(3)=   "lblOtrosDatos(2)"
      Item(2).Control(4)=   "txtOtrosDatos(0)"
      Item(2).Control(5)=   "lblOtrosDatos(3)"
      Item(2).Control(6)=   "txtOtrosDatos(1)"
      Item(2).Control(7)=   "txtOtrosDatos(2)"
      Item(2).Control(8)=   "lblOtrosDatos(4)"
      Item(2).Control(9)=   "dtpFechaNacimiento"
      Item(2).Control(10)=   "txtOtrosDatos(3)"
      Item(2).Control(11)=   "lblOtrosDatos(5)"
      Item(2).Control(12)=   "pbCarga(8)"
      Item(2).Control(13)=   "txtOtrosDatos(4)"
      Item(2).Control(14)=   "txtOtrosDatos(5)"
      Item(2).Control(15)=   "lblOtrosDatos(6)"
      Item(3).Caption =   "Observaciones"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "txtObservaciones"
      Begin XtremeSuiteControls.FlatEdit vcodvendedor 
         Height          =   315
         Left            =   2400
         TabIndex        =   85
         Top             =   2160
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton Pus 
         Height          =   315
         Left            =   3360
         TabIndex        =   84
         Top             =   2160
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit vporcentajeVendedor 
         Height          =   345
         Left            =   2400
         TabIndex        =   82
         Top             =   2880
         Width           =   5925
         _Version        =   851968
         _ExtentX        =   10451
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtOtrosDatos 
         Height          =   315
         Index           =   0
         Left            =   -67600
         TabIndex        =   25
         Top             =   1080
         Visible         =   0   'False
         Width           =   5895
         _Version        =   851968
         _ExtentX        =   10398
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         MaxLength       =   50
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosComerciales 
         Height          =   315
         Index           =   2
         Left            =   -67600
         TabIndex        =   17
         Top             =   1080
         Visible         =   0   'False
         Width           =   5895
         _Version        =   851968
         _ExtentX        =   10398
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   1
         Left            =   4080
         TabIndex        =   45
         Tag             =   "CodigoPostal"
         Top             =   1080
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   2
         Left            =   5400
         TabIndex        =   5
         Top             =   1080
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   0
         Left            =   2400
         TabIndex        =   3
         Top             =   720
         Width           =   5895
         _Version        =   851968
         _ExtentX        =   10398
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   1
         Left            =   2400
         TabIndex        =   4
         Top             =   1080
         Width           =   1575
         _Version        =   851968
         _ExtentX        =   2778
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   3
         Left            =   2400
         TabIndex        =   6
         Top             =   1440
         Width           =   2655
         _Version        =   851968
         _ExtentX        =   4683
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   4
         Left            =   5640
         TabIndex        =   7
         Top             =   1440
         Width           =   2655
         _Version        =   851968
         _ExtentX        =   4683
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   5
         Left            =   2400
         TabIndex        =   8
         Top             =   1800
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   6
         Left            =   4560
         TabIndex        =   9
         Top             =   1800
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   7
         Left            =   6800
         TabIndex        =   10
         Top             =   1800
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtObservaciones 
         Height          =   3015
         Left            =   -69760
         TabIndex        =   29
         Top             =   480
         Visible         =   0   'False
         Width           =   8535
         _Version        =   851968
         _ExtentX        =   15055
         _ExtentY        =   5318
         _StockProps     =   77
         BackColor       =   -2147483643
         MultiLine       =   -1  'True
         ScrollBars      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosComerciales 
         Height          =   315
         Index           =   3
         Left            =   -67600
         TabIndex        =   18
         Top             =   1440
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
         MaxLength       =   3
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   5
         Left            =   -66640
         TabIndex        =   53
         Tag             =   "TipoCliente"
         Top             =   1440
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosComerciales 
         Height          =   315
         Index           =   4
         Left            =   -66160
         TabIndex        =   19
         Top             =   1440
         Visible         =   0   'False
         Width           =   4455
         _Version        =   851968
         _ExtentX        =   7858
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosComerciales 
         Height          =   315
         Index           =   5
         Left            =   -67600
         TabIndex        =   20
         Top             =   1800
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
         MaxLength       =   3
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   6
         Left            =   -66640
         TabIndex        =   55
         Tag             =   "Actividad"
         Top             =   1800
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosComerciales 
         Height          =   315
         Index           =   6
         Left            =   -66160
         TabIndex        =   21
         Top             =   1800
         Visible         =   0   'False
         Width           =   4455
         _Version        =   851968
         _ExtentX        =   7858
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.DateTimePicker dtpAlta 
         Height          =   315
         Left            =   -67600
         TabIndex        =   57
         Top             =   720
         Visible         =   0   'False
         Width           =   1575
         _Version        =   851968
         _ExtentX        =   2778
         _ExtentY        =   556
         _StockProps     =   68
         CustomFormat    =   "01012009"
         Format          =   1
         CurrentDate     =   40284
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFechaNacimiento 
         Height          =   315
         Left            =   -63280
         TabIndex        =   59
         Top             =   720
         Visible         =   0   'False
         Width           =   1575
         _Version        =   851968
         _ExtentX        =   2778
         _ExtentY        =   556
         _StockProps     =   68
         CustomFormat    =   "01012009"
         Format          =   1
         CurrentDate     =   40284
      End
      Begin XtremeSuiteControls.FlatEdit txtOtrosDatos 
         Height          =   315
         Index           =   1
         Left            =   -67600
         TabIndex        =   26
         Top             =   1440
         Visible         =   0   'False
         Width           =   5895
         _Version        =   851968
         _ExtentX        =   10398
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         MaxLength       =   50
      End
      Begin XtremeSuiteControls.FlatEdit txtOtrosDatos 
         Height          =   315
         Index           =   2
         Left            =   -67600
         TabIndex        =   27
         Top             =   1800
         Visible         =   0   'False
         Width           =   5895
         _Version        =   851968
         _ExtentX        =   10398
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         MaxLength       =   50
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   8
         Left            =   8550
         TabIndex        =   11
         Top             =   2190
         Visible         =   0   'False
         Width           =   195
         _Version        =   851968
         _ExtentX        =   344
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   2
         Left            =   9600
         TabIndex        =   65
         Tag             =   "Vendedor"
         Top             =   2220
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   9
         Left            =   8820
         TabIndex        =   12
         Top             =   2190
         Visible         =   0   'False
         Width           =   675
         _Version        =   851968
         _ExtentX        =   1191
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosComerciales 
         Height          =   315
         Index           =   0
         Left            =   -67600
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
         MaxLength       =   3
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   4
         Left            =   -66640
         TabIndex        =   66
         Tag             =   "TipoIva"
         Top             =   720
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosComerciales 
         Height          =   315
         Index           =   1
         Left            =   -66160
         TabIndex        =   16
         Top             =   720
         Visible         =   0   'False
         Width           =   4455
         _Version        =   851968
         _ExtentX        =   7858
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosComerciales 
         Height          =   315
         Index           =   7
         Left            =   -67600
         TabIndex        =   22
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
         MaxLength       =   3
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   7
         Left            =   -66640
         TabIndex        =   67
         Tag             =   "Lista"
         Top             =   2160
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosComerciales 
         Height          =   315
         Index           =   8
         Left            =   -66160
         TabIndex        =   23
         Top             =   2160
         Visible         =   0   'False
         Width           =   4455
         _Version        =   851968
         _ExtentX        =   7858
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   10
         Left            =   2400
         TabIndex        =   13
         Top             =   2520
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   3
         Left            =   3360
         TabIndex        =   69
         Tag             =   "Reparto"
         Top             =   2520
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   11
         Left            =   3900
         TabIndex        =   14
         Top             =   2520
         Width           =   4395
         _Version        =   851968
         _ExtentX        =   7752
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosComerciales 
         Height          =   315
         Index           =   9
         Left            =   -67600
         TabIndex        =   24
         Top             =   2520
         Visible         =   0   'False
         Width           =   5895
         _Version        =   851968
         _ExtentX        =   10398
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtOtrosDatos 
         Height          =   315
         Index           =   3
         Left            =   -67600
         TabIndex        =   28
         Top             =   2160
         Visible         =   0   'False
         Width           =   5895
         _Version        =   851968
         _ExtentX        =   10398
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         MaxLength       =   50
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   8
         Left            =   -66640
         TabIndex        =   73
         Tag             =   "EstadoCliente"
         Top             =   2520
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtOtrosDatos 
         Height          =   315
         Index           =   4
         Left            =   -67600
         TabIndex        =   74
         Top             =   2520
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
         MaxLength       =   3
      End
      Begin XtremeSuiteControls.FlatEdit txtOtrosDatos 
         Height          =   315
         Index           =   5
         Left            =   -66160
         TabIndex        =   75
         Top             =   2520
         Visible         =   0   'False
         Width           =   4455
         _Version        =   851968
         _ExtentX        =   7858
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         MaxLength       =   50
      End
      Begin XtremeSuiteControls.FlatEdit vdesvendedor 
         Height          =   315
         Left            =   3900
         TabIndex        =   86
         Top             =   2160
         Width           =   4395
         _Version        =   851968
         _ExtentX        =   7752
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Porcentaje Vendedor:"
         Height          =   195
         Index           =   10
         Left            =   480
         TabIndex        =   83
         Top             =   2970
         Width           =   1755
      End
      Begin VB.Label lblOtrosDatos 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado del Proveedor :"
         Height          =   195
         Index           =   6
         Left            =   -69520
         TabIndex        =   76
         Top             =   2560
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblOtrosDatos 
         BackStyle       =   0  'Transparent
         Caption         =   "Mensaje Emergente :"
         Height          =   195
         Index           =   5
         Left            =   -69520
         TabIndex        =   72
         Top             =   2205
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblDatosComerciales 
         BackStyle       =   0  'Transparent
         Caption         =   "Credito Maximo :"
         Height          =   195
         Index           =   5
         Left            =   -69520
         TabIndex        =   71
         Top             =   2560
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Reparto:"
         Height          =   195
         Index           =   9
         Left            =   480
         TabIndex        =   70
         Top             =   2565
         Width           =   1755
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         Height          =   195
         Index           =   8
         Left            =   480
         TabIndex        =   64
         Top             =   2200
         Width           =   1755
      End
      Begin VB.Label lblDatosComerciales 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Precio :"
         Height          =   195
         Index           =   4
         Left            =   -69520
         TabIndex        =   63
         Top             =   2200
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblOtrosDatos 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Skype :"
         Height          =   195
         Index           =   4
         Left            =   -69500
         TabIndex        =   62
         Top             =   1840
         Visible         =   0   'False
         Width           =   1750
      End
      Begin VB.Label lblOtrosDatos 
         BackStyle       =   0  'Transparent
         Caption         =   "Sitio Web :"
         Height          =   195
         Index           =   3
         Left            =   -69500
         TabIndex        =   61
         Top             =   1480
         Visible         =   0   'False
         Width           =   1750
      End
      Begin VB.Label lblOtrosDatos 
         BackStyle       =   0  'Transparent
         Caption         =   "Correo Electronico :"
         Height          =   195
         Index           =   2
         Left            =   -69500
         TabIndex        =   60
         Top             =   1120
         Visible         =   0   'False
         Width           =   1750
      End
      Begin VB.Label lblOtrosDatos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Nacimiento:"
         Height          =   195
         Index           =   1
         Left            =   -65080
         TabIndex        =   58
         Top             =   765
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Label lblOtrosDatos 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Alta :"
         Height          =   195
         Index           =   0
         Left            =   -69500
         TabIndex        =   56
         Top             =   760
         Visible         =   0   'False
         Width           =   1750
      End
      Begin VB.Label lblDatosComerciales 
         BackStyle       =   0  'Transparent
         Caption         =   "Actividad :"
         Height          =   195
         Index           =   3
         Left            =   -69520
         TabIndex        =   54
         Top             =   1840
         Visible         =   0   'False
         Width           =   1750
      End
      Begin VB.Label lblDatosComerciales 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Proveedor :"
         Height          =   195
         Index           =   2
         Left            =   -69520
         TabIndex        =   52
         Top             =   1480
         Visible         =   0   'False
         Width           =   1750
      End
      Begin VB.Label lblDatosComerciales 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Cuit :"
         Height          =   195
         Index           =   1
         Left            =   -69520
         TabIndex        =   49
         Top             =   1120
         Visible         =   0   'False
         Width           =   1750
      End
      Begin VB.Label lblDatosComerciales 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Iva :"
         Height          =   195
         Index           =   0
         Left            =   -69520
         TabIndex        =   46
         Top             =   760
         Visible         =   0   'False
         Width           =   1750
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Celular:"
         Height          =   195
         Index           =   7
         Left            =   6120
         TabIndex        =   44
         Top             =   1840
         Width           =   540
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
         Height          =   195
         Index           =   6
         Left            =   4080
         TabIndex        =   43
         Top             =   1845
         Width           =   400
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono:"
         Height          =   195
         Index           =   5
         Left            =   480
         TabIndex        =   42
         Top             =   1840
         Width           =   1750
      End
      Begin VB.Label lblFicha 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pais:"
         Height          =   195
         Index           =   4
         Left            =   5160
         TabIndex        =   41
         Top             =   1485
         Width           =   555
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Provincia:"
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   40
         Top             =   1480
         Width           =   1750
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Localidad:"
         Height          =   195
         Index           =   2
         Left            =   4560
         TabIndex        =   39
         Top             =   1125
         Width           =   975
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo Postal:"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   38
         Top             =   1120
         Width           =   1750
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Domicilio:"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   37
         Top             =   760
         Width           =   1755
      End
   End
   Begin VB.PictureBox PicInferior 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      Picture         =   "frmProveedoresAlta.frx":0000
      ScaleHeight     =   615
      ScaleWidth      =   10965
      TabIndex        =   47
      Top             =   5250
      Width           =   10965
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   0
         Left            =   8520
         TabIndex        =   31
         Top             =   120
         Width           =   1275
         _Version        =   851968
         _ExtentX        =   2249
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Grabar<F2>"
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Picture         =   "frmProveedoresAlta.frx":50B3
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   1
         Left            =   9810
         TabIndex        =   48
         Top             =   120
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Cerrar"
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Picture         =   "frmProveedoresAlta.frx":54BA
      End
      Begin VB.Label lblWGestion 
         BackStyle       =   0  'Transparent
         Caption         =   "WGESTION 2010"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
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
         TabIndex        =   50
         Top             =   150
         Width           =   1770
      End
      Begin VB.Label lblWGestion 
         BackStyle       =   0  'Transparent
         Caption         =   "WGESTION 2010"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
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
         TabIndex        =   51
         Top             =   170
         Width           =   1770
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtAlta 
      Height          =   315
      Index           =   0
      Left            =   2580
      TabIndex        =   0
      Top             =   60
      Width           =   1905
      _Version        =   851968
      _ExtentX        =   3360
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Alignment       =   1
   End
   Begin VB.Label lblRol 
      BackStyle       =   0  'Transparent
      Caption         =   "Rol:"
      Height          =   195
      Index           =   4
      Left            =   4650
      TabIndex        =   81
      Top             =   120
      Width           =   360
   End
   Begin VB.Label lblAlta 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(*) Asociar con persona institucionales:"
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
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   77
      Top             =   1140
      Width           =   2130
   End
   Begin VB.Label lblAlta 
      BackStyle       =   0  'Transparent
      Caption         =   "Razon Social:"
      Height          =   195
      Index           =   2
      Left            =   1170
      TabIndex        =   35
      Top             =   840
      Width           =   990
   End
   Begin VB.Label lblAlta 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Proveedor :"
      Height          =   195
      Index           =   1
      Left            =   495
      TabIndex        =   34
      Top             =   480
      Width           =   2250
   End
   Begin VB.Label lblAlta 
      BackStyle       =   0  'Transparent
      Caption         =   "Código  del Proveedor :"
      Height          =   195
      Index           =   0
      Left            =   510
      TabIndex        =   33
      Top             =   120
      Width           =   1680
   End
   Begin VB.Shape shpSuperior 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   1605
      Left            =   0
      Top             =   0
      Width           =   10950
   End
End
Attribute VB_Name = "frmProveedoresAlta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vidVendedor As Long
Public vaccion, viente  As String
Private Sub Form_KeyPress(KeyAscii As Integer)
        On Error Resume Next
    
    If KeyAscii = 13 Then SendKeys "{TAB}"


    If Err Then GrabarLog "Form_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Load()
    On Error Resume Next

    With Me
        .Show
        .Top = 0
        .Left = 0
        .KeyPreview = True
    End With
    
    Call CentrarFormulario(Me)
    LimpiarCampos
    init

    If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub init()
Me.vtipoProveedor.Clear
Me.vtipoProveedor.AddItem ("Proveedor")
Me.vtipoProveedor.AddItem ("Eventuales")
Me.vtipoProveedor.AddItem ("Personal")
Me.vtipoProveedor.AddItem ("Funcionarios")
Me.vtipoProveedor.AddItem ("Externos")
Me.vtipoProveedor.AddItem ("Empresa")
Me.vtipoProveedor.AddItem ("Vendedor")
Me.vtipoProveedor.AddItem ("Rol1")
Me.vtipoProveedor.AddItem ("Rol2")
Me.vtipoProveedor.AddItem ("Rol3")
Me.vtipoProveedor.AddItem ("Rol4")
Me.vtipoProveedor.AddItem ("Rol5")
Me.vtipoProveedor.AddItem ("Rol6")
Me.vtipoProveedor.AddItem ("Rol7")
Me.vtipoProveedor.AddItem ("Rol8")
Me.vtipoProveedor.Text = "Proveedor"

vidVendedor = 0

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    If KeyCode = vbKeyF1 Then
        VerAyuda (Me.Name)
    End If
    
    If KeyCode = vbKeyF2 Then
        Grabar
    End If

If Err Then GrabarLog "Form_KeyUp", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub PbAcciones_Click(Index As Integer)
Dim vauxi As String

    On Error Resume Next

    Select Case Index
    
        Case 0
            
            If Me.viente = "frmcompras" Then
                frmCompras.txtProveedor(0).Text = Me.txtAlta(0)
                
            End If
            
            vauxi = Me.txtAlta(0).Text
            
            Grabar
            
            'Call frmCompras.txtProveedor_KeyPress(13, 1)
            
            If Not viente = "" Then
                Unload frmProveedores
                Unload Me
            End If
            
            
            If viente = "frmCompras" Then
                Call frmCompras.txtProveedor_KeyPress(0, 13)
                Unload Me
            End If
            
            If viente = "frmConsultas" Then
                frmConsultas.vbuscando.Text = vauxi
                frmConsultas.WindowState = 2
                frmConsultas.bandera = "enter"
                Unload Me
            End If
            
        Case 1
            Unload Me
    End Select

    If Err Then
        Exit Sub
       ' GrabarLog "PbAcciones_Click", Err.Number & " " & Err.Description, Me.Caption
    End If
End Sub
Private Sub pbCarga_Click(Index As Integer)
    On Error Resume Next

    vVuelveBusqueda = Me.Name
    vVieneBusqueda = pbCarga(Index).Tag

    Select Case Index
    
        Case 0

            'Foto
            With phtProveedor
                .PhotoFileName = ""
                .OpenPhotoFile

                If Not .PhotoFileName = "" Then
                    pbCierraFoto.Visible = True
                Else
                    pbCierraFoto.Visible = Not True
                End If

            End With
    
        Case 1, 2, 3, 4, 5, 6, 7, 8
            frmBusqueda.Show
    
    End Select

    
    If Err Then GrabarLog "pbCarga_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Grabar()
    On Error Resume Next

    If Not ValidarCampos() = True Then
        Exit Sub
    End If
    
    Dim rsProveedor As New ADODB.Recordset, sqlProveedor As String
    
    Select Case vaccion

        Case "Nuevo"
            sqlProveedor = "SELECT * FROM Proveedores WHERE 1=2"
        
        Case "Modificar"
            sqlProveedor = "SELECT * FROM Proveedores WHERE (Codigo = '" & Trim(txtAlta(0).Text) & "')"
        
        Case "Duplicar"
            
    End Select
        
    With rsProveedor
        Call .Open(sqlProveedor, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not .State = 0 Then
        
            Select Case vaccion
            
                Case "Nuevo"
                    .AddNew
                    .Fields("Codigo").Value = Trim(txtAlta(0).Text)
                    .Fields("codigoNum").Value = Val(txtAlta(0).Text)
                
                Case "Modificar"
                    'No hago nada
                    ' MsgBox sqlProveedor
                    
                Case "Duplicar"
                    .AddNew
                    .Fields("Codigo").Value = "" 'Tendria que traer el ultimo codigo
                    .Fields("codigo_num").Value = Val(txtAlta(0).Text)

            End Select
            
            'No Opcional
            .Fields("Nombre").Value = Left(txtAlta(1).Text, 255)
            .Fields("RazonSocial").Value = Left(txtAlta(2).Text, 255)
        
            'Call GuardarFoto(rsProveedor, phtProveedor.PhotoFileName)
        
            'Ficha
            .Fields("Direccion").Value = Left(txtFicha(0).Text, 150)
            .Fields("CodigoPostal").Value = txtFicha(1).Text
            .Fields("Localidad").Value = Left(txtFicha(2).Text, 150)
            .Fields("Provincia").Value = txtFicha(3).Text
            .Fields("Pais").Value = Left(txtFicha(4).Text, 50)
            .Fields("Telefono").Value = Left(txtFicha(5).Text, 20)
            .Fields("Fax").Value = Left(txtFicha(6).Text, 20)
            .Fields("Celular").Value = Left(txtFicha(7).Text, 20)
            
            '.Fields("idVendedor").Value = txtFicha(8).Text
            
            .Fields("idVendedor2").Value = vidVendedor
            
            .Fields("idReparto").Value = txtFicha(10).Text
                
            'Datos Comerciales
            .Fields("idTipoIva").Value = Left(txtDatosComerciales(0).Text, 3)
            .Fields("Cuit").Value = Left(txtDatosComerciales(2).Text, 15)
            .Fields("idTipoCliente").Value = txtDatosComerciales(3).Text
            .Fields("idActividad").Value = txtDatosComerciales(5).Text
            .Fields("idListas").Value = txtDatosComerciales(7).Text
            .Fields("CreditoMax").Value = Val(Format(txtDatosComerciales(9).Text, "########0.00"))
        
            'Otros Datos
            .Fields("Fecha_Alta").Value = strfechaMySQL(dtpAlta.Value)
            .Fields("Fecha_Nacimiento").Value = strfechaMySQL(dtpFechaNacimiento.Value)
            .Fields("E-Mail").Value = txtOtrosDatos(0).Text
            .Fields("Web").Value = txtOtrosDatos(1).Text
            .Fields("Skype").Value = txtOtrosDatos(2).Text
            .Fields("MensajeEmergente").Value = txtOtrosDatos(3).Text
            .Fields("idEstados").Value = txtOtrosDatos(4).Text
            
            '.Fields("Fecha_Baja").Value = strfechaMySQL(dtpAlta.Value)
            'Observaciones
            .Fields("Observaciones").Value = txtObservaciones.Text
            
            .Fields("idpersonas").Value = Me.vpersonas.Tag
            
            .Fields("tipocliente").Value = Me.vtipoProveedor.Text
            
            '.Fields("porcentaje_vendedor ").Value = vporcentajeVendedor.Text
            
           ' MsgBox sqlProveedor
            
       
            .Update
        Else
            MsgBox "Cuidado"
        
        End If
        
    End With

    sqlProveedor = ""
    
    rsProveedor.Close
    Set rsProveedor = Nothing


    LimpiarCampos


    If Err Then
        MsgBox Err.Description
        GrabarLog "Guardar", Err.Number & " " & Err.Description, Me.Name
    Else
        Unload Me
        frmProveedores.Buscar
    End If

End Sub
Private Function ValidarCampos() As Boolean
    On Error Resume Next

    Dim i As Integer
    
    ValidarCampos = True
    
    For i = 0 To Val(txtAlta.Count - 1)

        If Trim(txtAlta(i).Text) = "" Then
            MsgBox "Campos obligatorios vacios!", vbExclamation, "Mensaje ..."
            ValidarCampos = False
            Exit Function
        End If

    Next
    
    If (Trim(txtDatosComerciales(0).Text) = "001" Or Trim(txtDatosComerciales(0).Text) = "003") And (Trim(txtDatosComerciales(2).Text) = "") Then
        MsgBox "Debe ingresar el CUIT del Proveedor", vbExclamation, "Mensaje ..."
        ValidarCampos = False
        Exit Function
    End If
    
   ' anulado
    'If (Trim(txtDatosComerciales(0).Text) = "001" Or Trim(txtDatosComerciales(0).Text) = "003") And (ValidarCuit(txtDatosComerciales(2).Text) = False) Then
        If MsgBox("Este CUIT-CUIL no cumple con la validación. Continúa de todas manera ? ", vbYesNo, "Mensaje ...") = vbNo Then
            ValidarCampos = False
            Exit Function
        End If
    'End If
    
    If vaccion = "Nuevo" Then
        If Not Trim(TraerDato("Proveedores", "Codigo = '" & Trim(txtAlta(0).Text) & "'", "Codigo")) = "" Then
            MsgBox "Existe un registro con ese codigo!", vbExclamation, "Mensaje ..."
            ValidarCampos = False
            Exit Function
        End If
    End If
    
    If Err Then GrabarLog "ValidarCampos", Err.Number & " " & Err.Description, Me.Caption
End Function

Private Function GuardarFoto(rsFoto As ADODB.Recordset, _
                             vFilename As String) As Boolean
    On Error Resume Next
    
    If Not vFilename = "" Then
    
        Dim stmFoto As New ADODB.Stream
    
        With stmFoto

            If Not .State = adStateClosed Then .Close
            
            .Type = adTypeBinary
            .Open

            .LoadFromFile (vFilename)
        End With
    
    End If
    
    With rsFoto

        If Not (.State = 0) And Not (.EOF = True) Then
            If Not vFilename = "" Then
                .Fields("Foto").Value = stmFoto.Read
            Else
                .Fields("Foto").Value = Null
            End If
        End If
        
    End With
    
    If Not stmFoto.State = adStateClosed Then
        stmFoto.Close
        Set stmFoto = Nothing
    End If
    
    If Err Then
        GuardarFoto = Not True
        GrabarLog "GuardarFoto", Err.Number & " " & Err.Description, "Global"
    Else
        GuardarFoto = True
    End If

End Function

Private Sub LimpiarCampos()
    On Error Resume Next
    
    Dim i As Integer
    
    For i = 0 To txtAlta.Count - 1
        txtAlta(i).Text = ""
    Next
    
    phtProveedor.Reset
    pbCierraFoto.Visible = Not True
    
    For i = 0 To txtFicha.Count - 1
        txtFicha(i).Text = ""
    Next
    
    For i = 0 To txtDatosComerciales.Count - 1
        txtDatosComerciales(i).Text = ""
    Next
    
    For i = 0 To txtOtrosDatos.Count - 1
        txtOtrosDatos(i).Text = ""
    Next
    
    txtObservaciones.Text = ""
    
    vaccion = "Nuevo"
    txtAlta(0).Locked = False
    
    txtAlta(0).Text = Val(GenerarDato("SELECT MAX(Codigo) AS UltimoCodigo FROM Proveedores", "UltimoCodigo")) + 1
    txtAlta(0).Text = FormatoUltimoCodigo(4, txtAlta(0).Text)

    txtAlta(1).SetFocus
    
    If Err Then GrabarLog "Limpia", Err.Number & "-" & Err.Description, Me.Name
End Sub
Public Sub ModificarProveedor(vIDProveedor As Long)
    On Error Resume Next
    Dim v As String
    
    
    Dim rsProveedor As New ADODB.Recordset, sqlProveedor As String
    
    sqlProveedor = "SELECT * FROM Proveedores WHERE (idProveedores = " & vIDProveedor & ")"
    
    With rsProveedor
        Call .Open(sqlProveedor, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not (.EOF = True) And Not (.BOF = True) Then
        
            'No Opcionales
            txtAlta(0).Text = .Fields("codigo").Value
            txtAlta(0).Locked = True
        
            MsgBox "Usted está modificando clientes/proveedores." + Chr(13) + "Clic en aceptar para continuar", vbInformation, "Atención"

        
            txtAlta(1).Text = .Fields("Nombre").Value
            txtAlta(2).Text = .Fields("RazonSocial").Value
        
            If Not IsNull(.Fields("Foto").Value) = True And Not Trim(.Fields("Foto").Value) = "" Then
                BorrarArchivo (App.Path & "\" & .Fields("Codigo").Value & ".dat")
                phtProveedor.BlobToFile rsProveedor!Foto, App.Path & "\" & .Fields("Codigo").Value & ".dat"
                Call phtProveedor.AbrirFotoDesdeArchivo(App.Path & "\" & .Fields("Codigo").Value & ".dat")
                BorrarArchivo (App.Path & "\" & .Fields("Codigo").Value & ".dat")
                pbCierraFoto.Visible = True
            End If

            'Ficha
        
            Me.vpersonas.Tag = EsNulo(.Fields("idpersonas").Value)
            
            
            v = "select * from personas where id_personas=" + EsNulo(.Fields("idpersonas").Value)
            vpersonas.Text = traerDatos2(v, "apellido", pathDBMySQLComuna) + " " + traerDatos2(v, "nombre", pathDBMySQLComuna)
            
            
            Me.vporcentajeVendedor.Text = EsNulo(.Fields("porcentajeVendedor").Value)
            
            txtAlta(0).Text = .Fields("codigo").Value
            
            txtFicha(0).Text = EsNulo(.Fields("Direccion").Value)
            txtFicha(1).Text = EsNulo(.Fields("CodigoPostal").Value)
            txtFicha(2).Text = EsNulo(.Fields("Localidad").Value)
            txtFicha(3).Text = EsNulo(.Fields("Provincia").Value)
            txtFicha(4).Text = EsNulo(.Fields("Pais").Value)
            txtFicha(5).Text = EsNulo(.Fields("Telefono").Value)
            txtFicha(6).Text = EsNulo(.Fields("Fax").Value)
            txtFicha(7).Text = EsNulo(.Fields("Celular").Value)
            
            
            vcodvendedor.Text = EsNulo(TraerDato("proveedores", "idProveedores =  " & .Fields("idVendedor2").Value & "", "Codigo"))
            vdesvendedor.Text = EsNulo(TraerDato("proveedores", "idProveedores =  " & .Fields("idVendedor2").Value & "", "nombre"))
        
            txtFicha(10).Text = EsNulo(.Fields("idReparto").Value)
            txtFicha(11).Text = EsNulo(TraerDato("clireparto", "nreparto =  '" & .Fields("idReparto").Value & "'", "descrip"))
        
            'Datos Comerciales
            txtDatosComerciales(0).Text = EsNulo(.Fields("idTipoIva").Value)
            txtDatosComerciales(1).Text = EsNulo(TraerDato("TipoIva", "idTipoIva = '" & .Fields("idTipoIva").Value & "'", "TipoIva"))
            txtDatosComerciales(2).Text = EsNulo(.Fields("Cuit").Value)
            txtDatosComerciales(3).Text = EsNulo(.Fields("idTipoCliente").Value)
            txtDatosComerciales(4).Text = EsNulo(TraerDato("TipoClientes", "idTipoClientes = '" & .Fields("idTipoCliente").Value & "'", "Descripcion"))
            txtDatosComerciales(5).Text = .Fields("idActividad").Value
            txtDatosComerciales(6).Text = EsNulo(TraerDato("Actividades", "idActividades = '" & .Fields("idActividad").Value & "'", "Descripcion"))
            txtDatosComerciales(7).Text = EsNulo(.Fields("idListas").Value)
            txtDatosComerciales(8).Text = EsNulo(TraerDato("Listas", "idListas = '" & .Fields("idListas").Value & "'", "Lista"))
            txtDatosComerciales(9).Text = EsNulo(.Fields("CreditoMax").Value)
        
            'Otros datos
            If Not IsNull(.Fields("Fecha_Alta").Value) = True Then
                dtpAlta.Value = .Fields("Fecha_Alta").Value
            End If
            
            If Not IsNull(.Fields("Fecha_Nacimiento").Value) = True Then
                dtpFechaNacimiento.Value = .Fields("Fecha_Nacimiento").Value
            End If
    
            txtOtrosDatos(0).Text = EsNulo(.Fields("E-Mail").Value)
            txtOtrosDatos(1).Text = EsNulo(.Fields("Web").Value)
            txtOtrosDatos(2).Text = EsNulo(.Fields("Skype").Value)
            txtOtrosDatos(3).Text = EsNulo(.Fields("MensajeEmergente").Value)
            txtOtrosDatos(4).Text = EsNulo(.Fields("idEstados").Value)
            txtOtrosDatos(5).Text = EsNulo(TraerDato("Estados", "idEstados = '" & .Fields("idEstados").Value & "'", "Estado"))
        
            Me.vtipoProveedor.Text = EsNulo(.Fields("tipocliente").Value)
        
            'Observaciones
            txtObservaciones.Text = EsNulo(.Fields("Observaciones").Value)
        
        End If

    End With
    
    If Err Then GrabarLog "ModificarProveedor", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub pbCierraFoto_Click()
    On Error Resume Next

    phtProveedor.Reset
    pbCierraFoto.Visible = Not True

    If Err Then GrabarLog "pbCierraFoto_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub Pus_Click()
Dim vsql, vc1, vc2 As String

vsql = "(Select * from proveedores where tipocliente  = 'Vendedor') t"
vc1 = "Nombre"
vc2 = "Codigo"

Call fbuscarGrilla(vsql, vc1, vc2, Me.vdesvendedor.Name, Me)

End Sub

Private Sub PushButton1_Click()
Call fbuscarGrilla("personas", "nombre", "id_personas", Me.vpersonas.Name, Me, "apellido", True) ' ema:
End Sub

Private Sub txtAlta_Change(Index As Integer)
    On Error Resume Next

    If Index = 1 Then txtAlta(2).Text = txtAlta(1).Text

    If Err Then GrabarLog "txtAlta_Change", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtAlta_GotFocus(Index As Integer)
On Error Resume Next
If Index = 0 Then
    Dim vsql
    vsql = "select max(CONVERT(codigo,UNSIGNED INTEGER)) as c from proveedores"
    txtAlta(0) = traerDatos2(vsql, "c", pathDBMySQL) + 1
End If
If Err Then Exit Sub
End Sub

Private Sub txtFicha_Change(Index As Integer)
'If Index = 8 Then Me.vidVendedor = codigo2id(Me.vcodvendedor.TxT.Text)
End Sub

Private Sub vcodvendedor_Change()
vidVendedor = codigo2id(Me.vcodvendedor.Text)
End Sub

Private Sub vdesvendedor_Change()
    Me.vcodvendedor.Text = Me.vdesvendedor.Tag
End Sub
