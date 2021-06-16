VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{52463EDA-D668-43B6-8D47-4FA8035EF04A}#1.0#0"; "PhotoWSF.ocx"
Begin VB.Form frmEmpleadosAlta 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9390
   Icon            =   "frmEmpleadosAlta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   Begin PhotoWSF.Photo phtEmpleado 
      Height          =   1095
      Left            =   7200
      TabIndex        =   62
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1931
      BorderColor     =   -2147483640
      BackStyle       =   0
   End
   Begin XtremeSuiteControls.PushButton pbCarga 
      Height          =   315
      Index           =   0
      Left            =   9000
      TabIndex        =   23
      Top             =   120
      Width           =   315
      _Version        =   851968
      _ExtentX        =   556
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "..."
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton pbCierraFoto 
      Height          =   315
      Left            =   9000
      TabIndex        =   25
      Top             =   420
      Visible         =   0   'False
      Width           =   315
      _Version        =   851968
      _ExtentX        =   556
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "X"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtAlta 
      Height          =   315
      Index           =   1
      Left            =   2760
      TabIndex        =   1
      Top             =   480
      Width           =   4215
      _Version        =   851968
      _ExtentX        =   7435
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.TabControl TabAlta 
      Height          =   4935
      Left            =   120
      TabIndex        =   28
      Top             =   1560
      Width           =   9135
      _Version        =   851968
      _ExtentX        =   16113
      _ExtentY        =   8705
      _StockProps     =   68
      Color           =   8
      ItemCount       =   4
      Item(0).Caption =   "Ficha"
      Item(0).ControlCount=   17
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
      Item(1).Caption =   "Datos Comerciales"
      Item(1).ControlCount=   16
      Item(1).Control(0)=   "lblDatosComerciales(1)"
      Item(1).Control(1)=   "lblDatosComerciales(0)"
      Item(1).Control(2)=   "lblDatosComerciales(3)"
      Item(1).Control(3)=   "lblDatosComerciales(4)"
      Item(1).Control(4)=   "txtDatosComerciales(0)"
      Item(1).Control(5)=   "txtDatosComerciales(1)"
      Item(1).Control(6)=   "txtDatosComerciales(2)"
      Item(1).Control(7)=   "txtDatosComerciales(3)"
      Item(1).Control(8)=   "txtDatosComerciales(4)"
      Item(1).Control(9)=   "pbCarga(4)"
      Item(1).Control(10)=   "txtDatosComerciales(5)"
      Item(1).Control(11)=   "txtDatosComerciales(6)"
      Item(1).Control(12)=   "txtDatosComerciales(7)"
      Item(1).Control(13)=   "lblDatosComerciales(5)"
      Item(1).Control(14)=   "pbCarga(3)"
      Item(1).Control(15)=   "pbCarga(2)"
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
      Item(2).Control(12)=   "txtOtrosDatos(4)"
      Item(2).Control(13)=   "txtOtrosDatos(5)"
      Item(2).Control(14)=   "lblOtrosDatos(6)"
      Item(2).Control(15)=   "pbCarga(5)"
      Item(3).Caption =   "Observaciones"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "txtObservaciones"
      Begin XtremeSuiteControls.FlatEdit txtOtrosDatos 
         Height          =   315
         Index           =   0
         Left            =   -67600
         TabIndex        =   18
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
         TabIndex        =   12
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
         TabIndex        =   37
         Tag             =   "CodigoPostal"
         Top             =   1080
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   2
         Left            =   5400
         TabIndex        =   4
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
         TabIndex        =   2
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
         TabIndex        =   3
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
         TabIndex        =   5
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
         Index           =   5
         Left            =   2400
         TabIndex        =   7
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
         Index           =   7
         Left            =   6800
         TabIndex        =   9
         Top             =   1800
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtObservaciones 
         Height          =   3855
         Left            =   -69760
         TabIndex        =   22
         Top             =   600
         Visible         =   0   'False
         Width           =   8535
         _Version        =   851968
         _ExtentX        =   15055
         _ExtentY        =   6800
         _StockProps     =   77
         BackColor       =   -2147483643
         MultiLine       =   -1  'True
         ScrollBars      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosComerciales 
         Height          =   315
         Index           =   3
         Left            =   -67600
         TabIndex        =   13
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
         Index           =   3
         Left            =   -66640
         TabIndex        =   45
         Tag             =   "Actividad"
         Top             =   1440
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosComerciales 
         Height          =   315
         Index           =   4
         Left            =   -66160
         TabIndex        =   14
         Top             =   1440
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
         TabIndex        =   47
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
         TabIndex        =   49
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
         TabIndex        =   19
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
         TabIndex        =   20
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
      Begin XtremeSuiteControls.FlatEdit txtDatosComerciales 
         Height          =   315
         Index           =   0
         Left            =   -67600
         TabIndex        =   10
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
         Index           =   2
         Left            =   -66640
         TabIndex        =   54
         Tag             =   "TipoIva"
         Top             =   720
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosComerciales 
         Height          =   315
         Index           =   1
         Left            =   -66160
         TabIndex        =   11
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
         Index           =   5
         Left            =   -67600
         TabIndex        =   15
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
         Index           =   4
         Left            =   -66640
         TabIndex        =   55
         Tag             =   "Lista"
         Top             =   1800
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosComerciales 
         Height          =   315
         Index           =   6
         Left            =   -66160
         TabIndex        =   16
         Top             =   1800
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
         TabIndex        =   17
         Top             =   2160
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
         TabIndex        =   21
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
         Index           =   5
         Left            =   -66640
         TabIndex        =   58
         Tag             =   "EstadoCliente"
         Top             =   2520
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtOtrosDatos 
         Height          =   315
         Index           =   4
         Left            =   -67600
         TabIndex        =   59
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
         TabIndex        =   60
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
      Begin VB.Label lblOtrosDatos 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado del Empleado :"
         Height          =   195
         Index           =   6
         Left            =   -69520
         TabIndex        =   61
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
         TabIndex        =   57
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
         TabIndex        =   56
         Top             =   2205
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblDatosComerciales 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Precio :"
         Height          =   195
         Index           =   4
         Left            =   -69520
         TabIndex        =   53
         Top             =   1845
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblOtrosDatos 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Skype :"
         Height          =   195
         Index           =   4
         Left            =   -69500
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
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
         TabIndex        =   48
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
         TabIndex        =   46
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
         TabIndex        =   44
         Top             =   1485
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblDatosComerciales 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Cuit :"
         Height          =   195
         Index           =   1
         Left            =   -69520
         TabIndex        =   41
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
         TabIndex        =   38
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
         TabIndex        =   36
         Top             =   1840
         Width           =   540
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
         Height          =   195
         Index           =   6
         Left            =   4080
         TabIndex        =   35
         Top             =   1845
         Width           =   400
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono:"
         Height          =   195
         Index           =   5
         Left            =   480
         TabIndex        =   34
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
         TabIndex        =   33
         Top             =   1485
         Width           =   555
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Provincia:"
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   32
         Top             =   1480
         Width           =   1750
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Localidad:"
         Height          =   195
         Index           =   2
         Left            =   4560
         TabIndex        =   31
         Top             =   1125
         Width           =   975
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo Postal:"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   30
         Top             =   1120
         Width           =   1750
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Domicilio:"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   29
         Top             =   760
         Width           =   1755
      End
   End
   Begin VB.PictureBox PicInferior 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      Picture         =   "frmEmpleadosAlta.frx":6852
      ScaleHeight     =   555
      ScaleWidth      =   9405
      TabIndex        =   39
      Top             =   6600
      Width           =   9400
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   0
         Left            =   6960
         TabIndex        =   24
         Top             =   105
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Grabar"
         Appearance      =   6
         Picture         =   "frmEmpleadosAlta.frx":B905
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   1
         Left            =   8040
         TabIndex        =   40
         Top             =   105
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Cerrar"
         Appearance      =   6
         Picture         =   "frmEmpleadosAlta.frx":BD0C
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
         TabIndex        =   42
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
         TabIndex        =   43
         Top             =   170
         Width           =   1770
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtAlta 
      Height          =   315
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      _Version        =   851968
      _ExtentX        =   3201
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Alignment       =   1
   End
   Begin VB.Label lblAlta 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Empleado :"
      Height          =   195
      Index           =   1
      Left            =   495
      TabIndex        =   27
      Top             =   600
      Width           =   2250
   End
   Begin VB.Label lblAlta 
      BackStyle       =   0  'Transparent
      Caption         =   "Código  del Empleado :"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   26
      Top             =   240
      Width           =   2250
   End
   Begin VB.Shape shpSuperior 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      Height          =   1395
      Left            =   0
      Top             =   0
      Width           =   9420
   End
End
Attribute VB_Name = "frmEmpleadosAlta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vAccion As String
Public vVieneEmpleadosAlta As String
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    
    If KeyAscii = 13 Then SendKeys "{TAB}"

    If Err Then GrabarLog "Form_KeyPress", Err.Number & " " & Err.Description, Me.Caption
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
Private Sub Form_Load()
    On Error Resume Next

    With Me
        .Show
        .Top = 0
        .Left = 0
    End With
    
    LimpiarCampos

    If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub PbAcciones_Click(Index As Integer)
    On Error Resume Next

    Select Case Index
    
        Case 0
            Grabar
        
        Case 1
            Unload Me

    End Select

    If Err Then GrabarLog "PbAcciones_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub pbCarga_Click(Index As Integer)
    On Error Resume Next

    vVuelveBusqueda = Me.Name
    vVieneBusqueda = pbCarga(Index).Tag

    Select Case Index
    
        Case 0

            'Foto
            With phtEmpleado
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
    
    Dim rsempleados As New ADODB.Recordset, sqlEmpleados As String
    
    Select Case vAccion

        Case "Nuevo"
            sqlEmpleados = "SELECT * FROM Empleados WHERE 1=2"
        
        Case "Modificar"
            sqlEmpleados = "SELECT * FROM Empleados WHERE (Codigo = '" & Trim(txtAlta(0).Text) & "')"
        
        Case "Duplicar"
            
    End Select
        
    With rsempleados
        Call .Open(sqlEmpleados, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not .State = 0 Then
        
            Select Case vAccion
            
                Case "Nuevo"
                    .AddNew
                    .Fields("Codigo").Value = Trim(txtAlta(0).Text)
                    .Fields("codigoNum").Value = Val(txtAlta(0).Text)
                
                Case "Modificar"
                    'No hago nada
                    
                Case "Duplicar"
                    .AddNew
                    .Fields("Codigo").Value = "" 'Tendria que traer el ultimo codigo
                    .Fields("codigo_Num").Value = Val(txtAlta(0).Text)

            End Select
            
            'No Opcional
            .Fields("Nombre").Value = Left(txtAlta(1).Text, 255)
            '.Fields("RazonSocial").Value = Left(txtAlta(2).Text, 255)
        
            Call GuardarFoto(rsempleados, phtEmpleado.PhotoFileName)
        
            'Ficha
            .Fields("Direccion").Value = Left(txtFicha(0).Text, 150)
            .Fields("CodigoPostal").Value = txtFicha(1).Text
            .Fields("Localidad").Value = Left(txtFicha(2).Text, 150)
            .Fields("Provincia").Value = txtFicha(3).Text
            .Fields("Pais").Value = Left(txtFicha(4).Text, 50)
            .Fields("Telefono").Value = Left(txtFicha(5).Text, 20)
            .Fields("Fax").Value = Left(txtFicha(6).Text, 20)
            .Fields("Celular").Value = Left(txtFicha(7).Text, 20)
                
            'Datos Comerciales
            .Fields("idTipoIva").Value = Left(txtDatosComerciales(0).Text, 3)
            .Fields("Cuit").Value = Left(txtDatosComerciales(2).Text, 15)
            .Fields("idActividad").Value = txtDatosComerciales(3).Text
            .Fields("idListas").Value = txtDatosComerciales(5).Text
            .Fields("Quebranto").Value = Val(Format(txtDatosComerciales(7).Text, "########0.00"))
        
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
            .Fields("Observaciones").Value = EsNulo(txtObservaciones.Text)
            
            .Update
        
        End If
        
    End With

    sqlEmpleados = ""
    
    rsempleados.Close
    Set rsempleados = Nothing

    If Err Then
        GrabarLog "Guardar", Err.Number & " " & Err.Description, Me.Name
    Else

        Select Case vVieneEmpleadosAlta
        
            Case "frmEmpleados"
                LimpiarCampos
                frmEmpleados.Buscar

            Case "frmBusqueda"
                LimpiarCampos
                frmBusqueda.txtBusqueda_Change
                
        End Select
        
        Unload Me
        
    End If

End Sub
Private Function ValidarCampos() As Boolean
    On Error Resume Next

    Dim i As Integer
    
    ValidarCampos = True
    
    For i = 0 To Val(txtAlta.Count - 1)

        If Trim(txtAlta(i).Text) = "" Then
            MsgBox "Campos obligatorios vacios!", vbExclamation, "Mensaje ..."
            ValidarCampos = Not True
            Exit Function
        End If

    Next
    
    If (Trim(txtDatosComerciales(0).Text) = "001" Or Trim(txtDatosComerciales(0).Text) = "003") And (Trim(txtDatosComerciales(2).Text) = "") Then
        MsgBox "Debe ingresar el CUIT del Empleado", vbExclamation, "Mensaje ..."
        ValidarCampos = False
        Exit Function
    End If
        
    If (Trim(txtDatosComerciales(0).Text) = "001" Or Trim(txtDatosComerciales(0).Text) = "003") And (ValidarCuit(txtDatosComerciales(2).Text) = False) Then
        MsgBox "Debe ingresar el CUIT valido del Empleado", vbExclamation, "Mensaje ..."
        ValidarCampos = False
        Exit Function
    End If

    If vAccion = "Nuevo" Then
        If Not Trim(TraerDato("Empleados", "Codigo = '" & Trim(txtAlta(0).Text) & "'", "Codigo")) = "" Then
            MsgBox "Existe un registro con ese codigo!", vbExclamation, "Mensaje ..."
            ValidarCampos = Not True
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
    
    phtEmpleado.Reset
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
    
    vAccion = "Nuevo"
    vVieneEmpleadosAlta = ""
       
    txtAlta(0).Locked = Not True
    KeyPreview = True
    
    txtAlta(0).Text = Val(GenerarDato("SELECT MAX(Codigo) AS UltimoCodigo FROM Empleados", "UltimoCodigo")) + 1
    txtAlta(0).Text = FormatoUltimoCodigo(4, txtAlta(0).Text)

    txtAlta(1).SetFocus
    
    If Err Then GrabarLog "Limpia", Err.Number & "-" & Err.Description, Me.Name
End Sub
Public Sub ModificarEmpleado(vIDEmpleado As Long)
    On Error Resume Next
    
    Dim rsEmpleado As New ADODB.Recordset, sqlEmpleado As String
    
    sqlEmpleado = "SELECT * FROM Empleados WHERE (idEmpleados = " & vIDEmpleado & ")"
    
    With rsEmpleado
        Call .Open(sqlEmpleado, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not (.EOF = True) And Not (.BOF = True) Then
        
            'No Opcionales
            txtAlta(0).Text = EsNulo(.Fields("codigo").Value)
            txtAlta(0).Locked = True
        
            txtAlta(1).Text = .Fields("Nombre").Value
            txtAlta(2).Text = .Fields("RazonSocial").Value
        
            If Not IsNull(.Fields("Foto").Value) = True And Not Trim(.Fields("Foto").Value) = "" Then
                BorrarArchivo (App.Path & "\" & .Fields("Codigo").Value & ".dat")
                phtEmpleado.BlobToFile rsEmpleado!Foto, App.Path & "\" & .Fields("Codigo").Value & ".dat"
                Call phtEmpleado.AbrirFotoDesdeArchivo(App.Path & "\" & .Fields("Codigo").Value & ".dat")
                BorrarArchivo (App.Path & "\" & .Fields("Codigo").Value & ".dat")
                pbCierraFoto.Visible = True
            End If

            'Ficha
        
            txtFicha(0).Text = EsNulo(.Fields("Direccion").Value)
            txtFicha(1).Text = EsNulo(.Fields("CodigoPostal").Value)
            txtFicha(2).Text = EsNulo(.Fields("Localidad").Value)
            txtFicha(3).Text = EsNulo(.Fields("Provincia").Value)
            txtFicha(4).Text = EsNulo(.Fields("Pais").Value)
            txtFicha(5).Text = EsNulo(.Fields("Telefono").Value)
            txtFicha(6).Text = EsNulo(.Fields("Fax").Value)
            txtFicha(7).Text = EsNulo(.Fields("Celular").Value)
        
            'Datos Comerciales
            txtDatosComerciales(0).Text = EsNulo(.Fields("idTipoIva").Value)
            txtDatosComerciales(1).Text = EsNulo(TraerDato("TipoIva", "idTipoIva = '" & .Fields("idTipoIva").Value & "'", "TipoIva"))
            txtDatosComerciales(2).Text = EsNulo(.Fields("Cuit").Value)
            txtDatosComerciales(3).Text = .Fields("idActividad").Value
            txtDatosComerciales(4).Text = EsNulo(TraerDato("Actividades", "idActividades = '" & .Fields("idActividad").Value & "'", "Descripcion"))
            txtDatosComerciales(5).Text = EsNulo(.Fields("idListas").Value)
            txtDatosComerciales(6).Text = EsNulo(TraerDato("Listas", "idListas = '" & .Fields("idListas").Value & "'", "Lista"))
            txtDatosComerciales(7).Text = EsNulo(.Fields("Quebranto").Value)
            

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
        
            'Observaciones
            txtObservaciones.Text = EsNulo(.Fields("Observaciones").Value)
        
        End If

    End With
    
    If Err Then GrabarLog "ModificarCliente", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub pbCierraFoto_Click()
    On Error Resume Next

    phtEmpleado.Reset
    pbCierraFoto.Visible = Not True

    If Err Then GrabarLog "pbCierraFoto_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub txtAlta_Change(Index As Integer)
    On Error Resume Next

    If Index = 1 Then txtAlta(2).Text = txtAlta(1).Text

    If Err Then GrabarLog "txtAlta_Change", Err.Number & " " & Err.Description, Me.Caption
End Sub
