VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{52463EDA-D668-43B6-8D47-4FA8035EF04A}#1.0#0"; "PhotoWSF.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "KlexGrid.ocx"
Begin VB.Form frmClientesAlta 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Alta de Clientes"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9375
   Icon            =   "frmClientesAlta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   Begin PhotoWSF.Photo phtCliente 
      Height          =   1095
      Left            =   7200
      TabIndex        =   81
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
      TabIndex        =   39
      Top             =   120
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
      Left            =   9000
      TabIndex        =   41
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
      Left            =   2730
      TabIndex        =   1
      Top             =   420
      Width           =   4215
      _Version        =   851968
      _ExtentX        =   7435
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.FlatEdit txtAlta 
      Height          =   315
      Index           =   2
      Left            =   2730
      TabIndex        =   2
      Top             =   810
      Width           =   4215
      _Version        =   851968
      _ExtentX        =   7435
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.TabControl TabAlta 
      Height          =   4185
      Left            =   -60
      TabIndex        =   45
      Top             =   1800
      Width           =   9405
      _Version        =   851968
      _ExtentX        =   16589
      _ExtentY        =   7382
      _StockProps     =   68
      Color           =   8
      ItemCount       =   5
      Item(0).Caption =   "Ficha"
      Item(0).ControlCount=   34
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
      Item(0).Control(25)=   "txtFicha(12)"
      Item(0).Control(26)=   "pbCarga(10)"
      Item(0).Control(27)=   "txtFicha(13)"
      Item(0).Control(28)=   "lblFicha(10)"
      Item(0).Control(29)=   "lblFicha(11)"
      Item(0).Control(30)=   "txtFicha(14)"
      Item(0).Control(31)=   "vcodvendedor"
      Item(0).Control(32)=   "Pus"
      Item(0).Control(33)=   "vdesvendedor"
      Item(1).Caption =   "Datos Comerciales"
      Item(1).ControlCount=   21
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
      Item(1).Control(20)=   "vcboLista"
      Item(2).Caption =   "Otros Datos"
      Item(2).ControlCount=   18
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
      Item(2).Control(16)=   "lblOtrosDatos(7)"
      Item(2).Control(17)=   "txtOtrosDatos(6)"
      Item(3).Caption =   "Factura Automatica"
      Item(3).ControlCount=   17
      Item(3).Control(0)=   "cmdArticulo(0)"
      Item(3).Control(1)=   "txtArticulos(0)"
      Item(3).Control(2)=   "txtArticulos(1)"
      Item(3).Control(3)=   "pbCarga(9)"
      Item(3).Control(4)=   "txtArticulos(2)"
      Item(3).Control(5)=   "lblArticulos(1)"
      Item(3).Control(6)=   "lblArticulos(0)"
      Item(3).Control(7)=   "cmdArticulo(1)"
      Item(3).Control(8)=   "cmdArticulo(2)"
      Item(3).Control(9)=   "cmdArticulo(3)"
      Item(3).Control(10)=   "dtpFecha(0)"
      Item(3).Control(11)=   "dtpFecha(1)"
      Item(3).Control(12)=   "lblArticulos(4)"
      Item(3).Control(13)=   "lblArticulos(3)"
      Item(3).Control(14)=   "KlexArticulos"
      Item(3).Control(15)=   "cboIntervalo"
      Item(3).Control(16)=   "lblArticulos(2)"
      Item(4).Caption =   "Observaciones"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "txtObservaciones"
      Begin VB.ComboBox vcboLista 
         Height          =   315
         ItemData        =   "frmClientesAlta.frx":6852
         Left            =   -67600
         List            =   "frmClientesAlta.frx":686B
         TabIndex        =   102
         Text            =   "1"
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin XtremeSuiteControls.ComboBox cboIntervalo 
         Height          =   315
         Left            =   -67810
         TabIndex        =   37
         Top             =   450
         Visible         =   0   'False
         Width           =   4575
         _Version        =   851968
         _ExtentX        =   8070
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
      End
      Begin XtremeSuiteControls.PushButton cmdArticulo 
         Height          =   315
         Index           =   2
         Left            =   -63640
         TabIndex        =   84
         Top             =   3650
         Visible         =   0   'False
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "&Borrar"
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
      End
      Begin Grid.KlexGrid KlexArticulos 
         Height          =   1935
         Left            =   -69880
         TabIndex        =   97
         Top             =   1200
         Visible         =   0   'False
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3413
         EnterKeyBehaviour=   0
         BackColorAlternate=   0
         GridLinesFixed  =   2
         BackColorFixed  =   -2147483626
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColorFixed  =   8421504
         MouseIcon       =   "frmClientesAlta.frx":6884
         ScrollBars      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtOtrosDatos 
         Height          =   315
         Index           =   0
         Left            =   -67600
         TabIndex        =   30
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
         TabIndex        =   20
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
         TabIndex        =   54
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
         Width           =   1605
         _Version        =   851968
         _ExtentX        =   2831
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
         Height          =   3255
         Left            =   -69760
         TabIndex        =   38
         Top             =   600
         Visible         =   0   'False
         Width           =   8535
         _Version        =   851968
         _ExtentX        =   15055
         _ExtentY        =   5741
         _StockProps     =   77
         BackColor       =   -2147483643
         MultiLine       =   -1  'True
         ScrollBars      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosComerciales 
         Height          =   315
         Index           =   3
         Left            =   -67600
         TabIndex        =   21
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
         TabIndex        =   62
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
         TabIndex        =   22
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
         TabIndex        =   23
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
         TabIndex        =   64
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
         Left            =   -66130
         TabIndex        =   24
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
         TabIndex        =   28
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
         TabIndex        =   29
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
         TabIndex        =   31
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
         TabIndex        =   32
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
         Left            =   8250
         TabIndex        =   11
         Top             =   2160
         Width           =   165
         _Version        =   851968
         _ExtentX        =   291
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   2
         Left            =   8490
         TabIndex        =   72
         Tag             =   "Vendedor"
         Top             =   2160
         Width           =   255
         _Version        =   851968
         _ExtentX        =   450
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   9
         Left            =   8850
         TabIndex        =   12
         Top             =   2190
         Width           =   345
         _Version        =   851968
         _ExtentX        =   609
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosComerciales 
         Height          =   315
         Index           =   0
         Left            =   -67600
         TabIndex        =   18
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
         TabIndex        =   73
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
         TabIndex        =   19
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
         Left            =   -63880
         TabIndex        =   25
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
         Left            =   -63040
         TabIndex        =   74
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
         Left            =   -62650
         TabIndex        =   26
         Top             =   2160
         Visible         =   0   'False
         Width           =   975
         _Version        =   851968
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   10
         Left            =   2370
         TabIndex        =   13
         Top             =   2520
         Width           =   885
         _Version        =   851968
         _ExtentX        =   1561
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   3
         Left            =   3360
         TabIndex        =   75
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
         Left            =   3840
         TabIndex        =   14
         Top             =   2520
         Width           =   4455
         _Version        =   851968
         _ExtentX        =   7858
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosComerciales 
         Height          =   315
         Index           =   9
         Left            =   -67600
         TabIndex        =   27
         Top             =   2520
         Visible         =   0   'False
         Width           =   5955
         _Version        =   851968
         _ExtentX        =   10504
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtOtrosDatos 
         Height          =   315
         Index           =   3
         Left            =   -67570
         TabIndex        =   33
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
         TabIndex        =   79
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
         Left            =   -67570
         TabIndex        =   34
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
         TabIndex        =   35
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
      Begin XtremeSuiteControls.PushButton cmdArticulo 
         Height          =   315
         Index           =   0
         Left            =   -66280
         TabIndex        =   82
         Top             =   3650
         Visible         =   0   'False
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "&Nuevo"
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdArticulo 
         Height          =   315
         Index           =   1
         Left            =   -64960
         TabIndex        =   83
         Top             =   3650
         Visible         =   0   'False
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "&Modificar"
         BackColor       =   -2147483633
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtArticulos 
         Height          =   315
         Index           =   0
         Left            =   -68560
         TabIndex        =   85
         Top             =   3240
         Visible         =   0   'False
         Width           =   975
         _Version        =   851968
         _ExtentX        =   1720
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtArticulos 
         Height          =   315
         Index           =   1
         Left            =   -67000
         TabIndex        =   86
         Top             =   3240
         Visible         =   0   'False
         Width           =   3615
         _Version        =   851968
         _ExtentX        =   6376
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   9
         Left            =   -67480
         TabIndex        =   87
         Tag             =   "CodigoArticulo"
         Top             =   3240
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         BackColor       =   -2147483633
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtArticulos 
         Height          =   315
         Index           =   2
         Left            =   -62560
         TabIndex        =   88
         Top             =   3240
         Visible         =   0   'False
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2566
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdArticulo 
         Height          =   315
         Index           =   3
         Left            =   -62320
         TabIndex        =   89
         Top             =   3650
         Visible         =   0   'False
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "&Agregar"
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFecha 
         Height          =   315
         Index           =   0
         Left            =   -67840
         TabIndex        =   94
         Top             =   840
         Visible         =   0   'False
         Width           =   1575
         _Version        =   851968
         _ExtentX        =   2778
         _ExtentY        =   556
         _StockProps     =   68
         Enabled         =   0   'False
         CheckBox        =   -1  'True
         Format          =   1
         CurrentDate     =   40452
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFecha 
         Height          =   315
         Index           =   1
         Left            =   -64840
         TabIndex        =   95
         Top             =   840
         Visible         =   0   'False
         Width           =   1575
         _Version        =   851968
         _ExtentX        =   2778
         _ExtentY        =   556
         _StockProps     =   68
         Enabled         =   0   'False
         CheckBox        =   -1  'True
         Format          =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtOtrosDatos 
         Height          =   315
         Index           =   6
         Left            =   -67600
         TabIndex        =   36
         Top             =   2880
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
         Index           =   12
         Left            =   2370
         TabIndex        =   15
         Top             =   2880
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
         Index           =   10
         Left            =   3360
         TabIndex        =   99
         Tag             =   "TipoDocumento"
         Top             =   2880
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
         Index           =   13
         Left            =   3840
         TabIndex        =   16
         Top             =   2880
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   14
         Left            =   6720
         TabIndex        =   17
         Top             =   2880
         Width           =   1575
         _Version        =   851968
         _ExtentX        =   2778
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vcodvendedor 
         Height          =   315
         Left            =   2370
         TabIndex        =   105
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
         TabIndex        =   106
         Top             =   2160
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         BackColor       =   -2147483644
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit vdesvendedor 
         Height          =   315
         Left            =   3870
         TabIndex        =   107
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
         Caption         =   "Nro:"
         Height          =   195
         Index           =   11
         Left            =   5880
         TabIndex        =   101
         Top             =   2925
         Width           =   795
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Documento:"
         Height          =   195
         Index           =   10
         Left            =   480
         TabIndex        =   100
         Top             =   2925
         Width           =   1755
      End
      Begin VB.Label lblOtrosDatos 
         BackStyle       =   0  'Transparent
         Caption         =   "Password Web"
         Height          =   195
         Index           =   7
         Left            =   -69520
         TabIndex        =   98
         Top             =   2920
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblArticulos 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Fin :"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   -65920
         TabIndex        =   96
         Top             =   885
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label lblArticulos 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio :"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   -69760
         TabIndex        =   93
         Top             =   885
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblArticulos 
         BackStyle       =   0  'Transparent
         Caption         =   "Intervalo de Ejecucion :"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   -69760
         TabIndex        =   92
         Top             =   520
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblArticulos 
         BackStyle       =   0  'Transparent
         Caption         =   "Articulo :"
         Enabled         =   0   'False
         Height          =   195
         Index           =   3
         Left            =   -69760
         TabIndex        =   91
         Top             =   3285
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblArticulos 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio:"
         Enabled         =   0   'False
         Height          =   195
         Index           =   4
         Left            =   -63160
         TabIndex        =   90
         Top             =   3285
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblOtrosDatos 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado del Cliente :"
         Height          =   195
         Index           =   6
         Left            =   -69520
         TabIndex        =   80
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
         TabIndex        =   78
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
         TabIndex        =   77
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
         TabIndex        =   76
         Top             =   2565
         Width           =   1755
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
         Height          =   195
         Index           =   8
         Left            =   480
         TabIndex        =   71
         Top             =   2200
         Width           =   1755
      End
      Begin VB.Label lblDatosComerciales 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Precio :"
         Height          =   195
         Index           =   4
         Left            =   -69520
         TabIndex        =   70
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
         TabIndex        =   69
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
         TabIndex        =   68
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
         TabIndex        =   67
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
         TabIndex        =   66
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
         TabIndex        =   65
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
         TabIndex        =   63
         Top             =   1840
         Visible         =   0   'False
         Width           =   1750
      End
      Begin VB.Label lblDatosComerciales 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Cliente :"
         Height          =   195
         Index           =   2
         Left            =   -69520
         TabIndex        =   61
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
         TabIndex        =   58
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
         TabIndex        =   55
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
         TabIndex        =   53
         Top             =   1840
         Width           =   540
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
         Height          =   195
         Index           =   6
         Left            =   4080
         TabIndex        =   52
         Top             =   1845
         Width           =   400
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono:"
         Height          =   195
         Index           =   5
         Left            =   480
         TabIndex        =   51
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
         TabIndex        =   50
         Top             =   1485
         Width           =   555
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Provincia:"
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   49
         Top             =   1480
         Width           =   1750
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Localidad:"
         Height          =   195
         Index           =   2
         Left            =   4560
         TabIndex        =   48
         Top             =   1125
         Width           =   975
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo Postal:"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   47
         Top             =   1120
         Width           =   1750
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Domicilio:"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   46
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
      Left            =   -60
      Picture         =   "frmClientesAlta.frx":68A0
      ScaleHeight     =   555
      ScaleWidth      =   9405
      TabIndex        =   56
      Top             =   6000
      Width           =   9400
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   0
         Left            =   6720
         TabIndex        =   40
         Top             =   105
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Grabar <F2>"
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Picture         =   "frmClientesAlta.frx":B953
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   1
         Left            =   8040
         TabIndex        =   57
         Top             =   105
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Cerrar"
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Picture         =   "frmClientesAlta.frx":BD5A
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
         TabIndex        =   59
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
         TabIndex        =   60
         Top             =   170
         Width           =   1770
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtAlta 
      Height          =   315
      Index           =   0
      Left            =   2730
      TabIndex        =   0
      Top             =   30
      Width           =   1815
      _Version        =   851968
      _ExtentX        =   3201
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Alignment       =   1
   End
   Begin XtremeSuiteControls.ComboBox vtipoProveedor 
      Height          =   315
      Left            =   2730
      TabIndex        =   103
      Top             =   1230
      Width           =   4215
      _Version        =   851968
      _ExtentX        =   7435
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "Proveedor"
   End
   Begin VB.Label lblRol 
      BackStyle       =   0  'Transparent
      Caption         =   "Rol:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   1500
      TabIndex        =   104
      Top             =   1260
      Width           =   330
   End
   Begin VB.Label lblAlta 
      BackStyle       =   0  'Transparent
      Caption         =   "Razon Social:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   900
      TabIndex        =   44
      Top             =   810
      Width           =   1530
   End
   Begin VB.Label lblAlta 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Cliente :"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   495
      TabIndex        =   43
      Top             =   450
      Width           =   2250
   End
   Begin VB.Label lblAlta 
      BackStyle       =   0  'Transparent
      Caption         =   "Código  del Cliente :"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   540
      TabIndex        =   42
      Top             =   90
      Width           =   2250
   End
End
Attribute VB_Name = "frmClientesAlta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vaccion As String
Public vVieneClientesAlta, viente As String
Dim md5PasswordCliente As New MD5
Dim idVendedor As Long

Private Sub cboIntervalo_Click()
On Error Resume Next

    cboIntervalo.Tag = TraerDato("Intervalos", "Intervalo = '" & Trim(cboIntervalo.Text) & "'", "idIntervalos")

If Err Then GrabarLog "cboIntervalo_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cboIntervalo_GotFocus()
On Error Resume Next

    Call CargarComboNew("Intervalos", "Intervalo", cboIntervalo, True)

If Err Then GrabarLog "cboIntervalo_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cmdArticulo_Click(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 0
            NuevoArticulo
            HabilitarArticulos (True)

        Case 1
            'ModificarArticulo
        
        Case 2
            BorrarArticulo
            NuevoArticulo
            'CargarGrillaFacturaAutomatica (txtAlta(0).Text)
        
        Case 3
            If ValidarArticulo() = True Then
                AgregarArticulo
                NuevoArticulo
            End If
            
            'CargarGrillaArticulosProveedores (txtAlta(0).Text)
            
    End Select

If Err Then GrabarLog "cmdArticulo_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Function ValidarArticulo() As Boolean
On Error Resume Next
    
    ValidarArticulo = False
    
    If Trim(txtArticulos(0).Text) = "" Then
        MsgBox "Debe seleccionar un Articulo ", vbInformation, "Mensaje ..."
        Exit Function
    End If
    
    If Val(txtArticulos(2).Text) = 0 Then
        MsgBox "Debe ingresar un Precio !!!", vbInformation, "Mensaje ..."
        Exit Function
    End If
    
    ValidarArticulo = True

If Err Then GrabarLog "ValidarArticulo", Err.Number & " " & Err.Description, Me.Caption
End Function
Private Sub CargarGrillaFacturaAutomatica(vCodigoCliente As String)
On Error Resume Next
    
    Dim rsFacturaAutomatica As New ADODB.Recordset, sqlFacturaAutomatica As String, i As Integer

    sqlFacturaAutomatica = "SELECT * FROM FacturaAutomatica FA INNER JOIN FacturaAutomaticaDetalle FAD ON FA.idFacturaAutomatica=FAD.idFacturaAutomatica WHERE (CodigoCliente = '" & vCodigoCliente & "')"

    With rsFacturaAutomatica
        Call .Open(sqlFacturaAutomatica, ConnDDBB, adOpenStatic, adLockPessimistic)
                
        HabilitarArticulos (True)
                
        If Not .EOF = True Then
            .MoveFirst
            
            dtpFecha(0).Value = EsNulo(.Fields("FechaInicio").Value)
            dtpFecha(1).Value = EsNulo(.Fields("FechaFin").Value)
            
            cboIntervalo.Tag = EsNulo(.Fields("idIntervalos").Value)
            cboIntervalo.Text = TraerDato("Intervalos", "idIntervalos = " & Val(cboIntervalo.Tag) & "", "Intervalo")
            
            FormatoGrillaArticulos (.RecordCount)
        Else
            For i = 0 To KlexArticulos.Cols - 1
                KlexArticulos.TextMatrix(1, i) = ""
            Next
        End If
        
        Do Until .EOF = True
            KlexArticulos.TextMatrix(.AbsolutePosition, 1) = EsNulo(.Fields("idFacturaAutomaticaDetalle").Value)
            KlexArticulos.TextMatrix(.AbsolutePosition, 2) = "C" & EsNulo(.Fields("CodigoArticulo").Value)
            KlexArticulos.TextMatrix(.AbsolutePosition, 3) = EsNulo(TraerDato("Articulos", "Codigo = '" & EsNulo(.Fields("CodigoArticulo").Value) & "'", "Descrip"))
            KlexArticulos.TextMatrix(.AbsolutePosition, 4) = EsNulo(.Fields("PrecioArticulo").Value)
            
            .MoveNext
        Loop
    
    End With

    sqlFacturaAutomatica = ""
    
    If rsFacturaAutomatica.State = 1 Then
        rsFacturaAutomatica.Close
        Set rsFacturaAutomatica = Nothing
    End If
    
If Err Then GrabarLog "CargarGrillaFacturaAutomatica", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub HabilitarArticulos(vHabilitar As Boolean)
On Error Resume Next

    Dim i As Integer
    
    pbCarga(9).Enabled = vHabilitar

    For i = 0 To txtArticulos.Count - 1
        txtArticulos(i).Enabled = vHabilitar
        txtArticulos(i).Locked = Not vHabilitar
        txtArticulos(i).Text = ""
        txtArticulos(i).Tag = ""
    Next
    
    i = 0
    For i = 0 To lblArticulos.Count - 1
        lblArticulos(i).Enabled = vHabilitar
    Next
    
    i = 0
    For i = 0 To dtpFecha.Count - 1
        dtpFecha(i).Enabled = vHabilitar
    Next
    
    cboIntervalo.Enabled = vHabilitar
    cboIntervalo.Text = ""
    cboIntervalo.Tag = ""
    
   
If Err Then GrabarLog "HabilitarArticulos", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub NuevoArticulo()
On Error Resume Next
    
    Dim i As Integer

    i = 0
    For i = 0 To txtArticulos.Count - 1
        txtArticulos(i).Text = ""
        txtArticulos(i).Tag = ""
    Next

If Err Then GrabarLog "NuevoArticulo", Err.Number & "  " & Err.Description, Me.Caption
End Sub
Private Sub ModificarArticulo()
On Error Resume Next
    
    With KlexArticulos
        If Not Trim(.TextMatrix(.RowSel, 1)) = "" Then
            NuevoArticulo
            txtArticulos(0).Tag = .TextMatrix(.RowSel, 1)
            txtArticulos(0).Text = .TextMatrix(.RowSel, 2)
            txtArticulos(1).Text = .TextMatrix(.RowSel, 3)
            txtArticulos(2).Text = .TextMatrix(.RowSel, 6)
        Else
            MsgBox "Debe seleccionar un Registro para poder modificarlo!!", vbExclamation, "Mensaje ..."
        End If
    End With

    
If Err Then GrabarLog "ModificarArticulo", Err.Number & "  " & Err.Description, Me.Caption
End Sub
Private Sub BorrarArticulo()
On Error Resume Next
    
    Dim vFilaABorrar As Long
        
    With KlexArticulos
        
        'Controlo si selecciono una fila valida
        If Not Trim(.TextMatrix(.RowSel, 2)) = "" Then
            vFilaABorrar = .Row
            
            'Controlo que no se haya guardado previamente
            If Not Trim(.TextMatrix(.RowSel, 1)) = "" Then
                Call BorrarBase("FacturaAutomatica WHERE (CodigoProveedor = '" & (txtAlta(0).Text) & "') AND (CodigoArticulo = " & Trim(.TextMatrix(vFilaABorrar, 1)) & ")", pathDBMySQL)
            End If
            
            .RemoveItem (vFilaABorrar)
        Else
            MsgBox "Debe seleccionar un Registro para poder Borrarlo!!", vbExclamation, "Mensaje ..."
        End If
    End With
    
If Err Then GrabarLog "BorrarArticulo", Err.Number & "  " & Err.Description, Me.Caption
End Sub
Private Sub AgregarArticulo()
On Error Resume Next
    
    Dim j As Integer, i As Integer
    
    With KlexArticulos
        If .Rows <= 2 And .TextMatrix(.Rows - 1, 2) = "" Then
            FormatoGrillaArticulos (1)
        Else
            .Rows = .Rows + 1
        End If
        j = .Rows - 1
        
        
        .TextMatrix(j, 0) = ""
        .TextMatrix(j, 1) = ""
        
        .TextMatrix(j, 2) = "C" & (EsNulo(txtArticulos(0).Text))
        .TextMatrix(j, 3) = EsNulo(txtArticulos(1).Text)
        .TextMatrix(j, 4) = EsNulo(txtArticulos(2).Text)
    
    End With
    
If Err Then GrabarLog "AgregarArticulo", Err.Number & "  " & Err.Description, Me.Caption
End Sub
Private Sub FormatoGrillaArticulos(vCantidadRenglones As Integer)
On Error Resume Next

    Dim i As Integer

    With KlexArticulos
        .FixedRows = 1
        .FixedCols = 1
    
        .Cols = 10
        .Rows = vCantidadRenglones + 1
        
        If vCantidadRenglones = 1 Then
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
                .ColWidth(i) = 0
            Next
        End If
        
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 125
        
        'Aca Pego el IdFDetalle-Entonces se si modifico o NO
        .TextMatrix(0, 1) = "idFacturaAutomatica"
        .ColWidth(1) = 0
        
        .TextMatrix(0, 2) = "Codigo"
        .ColWidth(2) = 1250
        
        .TextMatrix(0, 3) = "Descripcion"
        .ColWidth(3) = 6000
        
        .TextMatrix(0, 4) = "Precio"
        .ColWidth(4) = 1250
        
        
        If .Rows = 2 Then
            .Row = 1
        Else
            .Row = .Rows
        End If
        
        .CellBackColor = &HFFFCCC
        
        .EnterKeyBehaviour = klexEKNone
        .BackColorAlternate = &HE0E0E0

    End With
    
If Err Then GrabarLog "FormatoGrillaArticulos", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub Form_KeyPress(Keyascii As Integer)
    On Error Resume Next
    
    If Keyascii = 13 Then SendKeys "{TAB}"

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
    
    FormatoGrillaArticulos (1)
    LimpiarCampos
 
    Call CentrarFormulario(Me)
    
    init
    
 
    If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub init()
Me.vtipoProveedor.Clear
Me.vtipoProveedor.AddItem ("Cliente")
Me.vtipoProveedor.AddItem ("Eventuales")
Me.vtipoProveedor.AddItem ("Personal")
Me.vtipoProveedor.AddItem ("Funcionarios")
Me.vtipoProveedor.AddItem ("Externos")
Me.vtipoProveedor.AddItem ("Empresa")
Me.vtipoProveedor.AddItem ("Rol1")
Me.vtipoProveedor.AddItem ("Rol2")
Me.vtipoProveedor.AddItem ("Rol3")
Me.vtipoProveedor.AddItem ("Rol4")
Me.vtipoProveedor.AddItem ("Rol5")
Me.vtipoProveedor.AddItem ("Rol6")
Me.vtipoProveedor.AddItem ("Rol7")
Me.vtipoProveedor.AddItem ("Rol8")
Me.vtipoProveedor.AddItem ("Creditos")
Me.vtipoProveedor.Text = "Cliente"

idVendedor = 0
End Sub

Private Sub PbAcciones_Click(Index As Integer)
    On Error Resume Next

    Select Case Index
    
        Case 0
        
        
           If Me.viente = "frmRemito" Then
                frmRemito.txtCliente(0).Text = Me.txtAlta(0)
               ' frmRemito.txtCliente(0).Tag = Me.txtAlta(0)
            End If
            
            Grabar
            
            'Call frmCompras.txtProveedor_KeyPress(13, 1)
            
            If Not viente = "" Then
                Call frmRemito.txtCliente_KeyPress(0, 13)
                'Unload frmClientes
                Unload Me
            End If
            
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
            With phtCliente
                .PhotoFileName = ""
                .OpenPhotoFile

                If Not .PhotoFileName = "" Then
                    pbCierraFoto.Visible = True
                Else
                    pbCierraFoto.Visible = Not True
                End If

            End With
        
        Case 1 To 10
            frmBusqueda.Show
    
    End Select
    
    If Err Then GrabarLog "pbCarga_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Grabar()
    On Error Resume Next

    If Not ValidarCampos() = True Then
        Exit Sub
    End If
    
    Dim rsClientes As New ADODB.Recordset, sqlClientes As String
    
    Select Case vaccion

        Case "Nuevo"
            sqlClientes = "SELECT * FROM Clientes WHERE 1=2"
        
        Case "Modificar"
            sqlClientes = "SELECT * FROM Clientes WHERE (Codigo = '" & Trim(txtAlta(0).Text) & "')"
        
        Case "Duplicar"
            
    End Select
        
    With rsClientes
        Call .Open(sqlClientes, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not .State = 0 Then
        
            Select Case vaccion
            
                Case "Nuevo"
                    .AddNew
                    .Fields("Codigo").Value = Trim(txtAlta(0).Text)
                    .Fields("codigoNum").Value = Val(txtAlta(0).Text)
                    .Fields("PasswordWeb").Value = md5PasswordCliente.DigestStrToHexStr(Trim(txtOtrosDatos(6).Text))
                    
                Case "Modificar"
                    'No hago nada
                    
                Case "Duplicar"
                    .AddNew
                    .Fields("Codigo").Value = "" 'Tendria que traer el ultimo codigo
                    .Fields("codigo_Num").Value = Val(txtAlta(0).Text)
                    .Fields("PasswordWeb").Value = md5PasswordCliente.DigestStrToHexStr(Trim(txtOtrosDatos(6).Text))

            End Select
            
            'No Opcional
            .Fields("Nombre").Value = Left(txtAlta(1).Text, 255)
            .Fields("RazonSocial").Value = Left(txtAlta(2).Text, 255)
        
            'Call GuardarFoto(rsClientes, phtCliente.PhotoFileName)
        
            'Ficha
            .Fields("Direccion").Value = Left(txtFicha(0).Text, 150)
            .Fields("CodigoPostal").Value = txtFicha(1).Text
            .Fields("Localidad").Value = Left(txtFicha(2).Text, 150)
            .Fields("Provincia").Value = txtFicha(3).Text
            .Fields("Pais").Value = Left(txtFicha(4).Text, 50)
            .Fields("Telefono").Value = Left(txtFicha(5).Text, 20)
            .Fields("Fax").Value = Left(txtFicha(6).Text, 20)
            .Fields("Celular").Value = Left(txtFicha(7).Text, 20)
            ''.Fields("idVendedor").Value = Left(EsNulo(txtFicha(8).Text), 20)
            
            .Fields("idVendedor").Value = Left(Me.vcodvendedor.Text, 3)
            
            .Fields("idVendedor2").Value = idVendedor
            
            .Fields("idReparto").Value = EsNulo(txtFicha(10).Text)
            .Fields("TipoDocumento").Value = txtFicha(12).Text
            .Fields("NroDocumento").Value = txtFicha(14).Text
            
            'Datos Comerciales
            .Fields("idTipoIva").Value = Left(txtDatosComerciales(0).Text, 3)
            .Fields("Cuit").Value = Left(txtDatosComerciales(2).Text, 15)
            .Fields("idTipoCliente").Value = txtDatosComerciales(3).Text
            .Fields("idActividad").Value = txtDatosComerciales(5).Text
            .Fields("idListas").Value = vcboLista.Text ' txtDatosComerciales(7).Text
            .Fields("CreditoMax").Value = Val(Format(txtDatosComerciales(9).Text, "########0.00"))
        
            'Otros Datos
            .Fields("Fecha_Alta").Value = strfechaMySQL(dtpAlta.Value)
            .Fields("Fecha_Nacimiento").Value = strfechaMySQL(dtpFechaNacimiento.Value)
            .Fields("E-Mail").Value = txtOtrosDatos(0).Text
            .Fields("Web").Value = txtOtrosDatos(1).Text
            .Fields("Skype").Value = txtOtrosDatos(2).Text
            .Fields("MensajeEmergente").Value = txtOtrosDatos(3).Text
            .Fields("idEstados").Value = txtOtrosDatos(4).Text
            
            .Fields("tipocliente").Value = Trim(Me.vtipoProveedor.Text)
            
            '.Fields("Fecha_Baja").Value = strfechaMySQL(dtpAlta.Value)
            'Observaciones
            .Fields("Observaciones").Value = txtObservaciones.Text
            
            .Update
        
            If Not Trim(KlexArticulos.TextMatrix(1, 2)) = "" Then
                Call GuardarFacturaAutomatica(txtAlta(0).Text)
            End If
            
        End If
        
    End With

    sqlClientes = ""
    
    rsClientes.Close
    Set rsClientes = Nothing

    If Err < 0 Then
        GrabarLog "Guardar", Err.Number & " " & Err.Description, Me.Name
    Else

        Select Case vVieneClientesAlta
        
            Case "frmClientes"
                LimpiarCampos
                frmClientes.Buscar

            Case "frmBusqueda"
                LimpiarCampos
                frmBusqueda.txtBusqueda_Change
                
        End Select
        
        Unload Me
        
    End If

End Sub
Private Sub GuardarFacturaAutomatica(vCodigoCliente As String)
On Error Resume Next

    Dim rsFacturaAutomatica As New ADODB.Recordset, sqlFacturaAutomatica As String, i As Integer, j As Integer
    
    sqlFacturaAutomatica = "SELECT * FROM FacturaAutomatica WHERE (CodigoCliente = '" & vCodigoCliente & "')"

    With rsFacturaAutomatica
        Call .Open(sqlFacturaAutomatica, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If .EOF = True Then
            
            .AddNew
            
            .Fields("CodigoCliente").Value = vCodigoCliente
            
            If IsNull(dtpFecha(0).Value) = True Then
                .Fields("FechaInicio").Value = Date
            Else
                .Fields("FechaInicio").Value = strfechaMySQL(dtpFecha(0).Value)
            End If
            
            .Fields("FechaFin").Value = dtpFecha(1).Value
            
            .Fields("idIntervalos").Value = Val(cboIntervalo.Tag)
            
            .Fields("FechaProximaEjecucion").Value = ControlarEjecuciones(dtpFecha(0).Value, .Fields("idIntervalos").Value)
            .Fields("FechaUltimaEjecucion").Value = Null
            
            .Update
        End If
        
        For i = 1 To Val(KlexArticulos.Rows - 1)
            If KlexArticulos.TextMatrix(i, 1) = "" Then
                Call EjecutarScript("INSERT INTO FacturaAutomaticaDetalle (idFacturaAutomatica, CodigoArticulo, PrecioArticulo) VALUES (" & .Fields("idFacturaAutomatica").Value & ",'" & Mid(KlexArticulos.TextMatrix(i, 2), 2, 4) & "'," & Val(KlexArticulos.TextMatrix(i, 4)) & ")")
            End If
        Next
            
        
    End With
    
    sqlFacturaAutomatica = ""
    
    If rsFacturaAutomatica.State = 1 Then
        rsFacturaAutomatica.Close
        Set rsFacturaAutomatica = Nothing
    End If

If Err Then GrabarLog "GuardarFacturaAutomatica", Err.Number & " " & Err.Description, Me.Caption
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
        MsgBox "Debe ingresar el CUIT del Cliente", vbExclamation, "Mensaje ..."
        ValidarCampos = False
        Exit Function
    End If
        
  '  If (Trim(txtDatosComerciales(0).Text) = "001" Or Trim(txtDatosComerciales(0).Text) = "003") And (ValidarCuit(txtDatosComerciales(2).Text) = False) Then
  '      MsgBox "Debe ingresar el CUIT valido del Cliente", vbExclamation, "Mensaje ..."
  '      ValidarCampos = False
  '      Exit Function
  '  End If

    If vaccion = "Nuevo" Then
        If Not Trim(TraerDato("Clientes", "Codigo = '" & Trim(txtAlta(0).Text) & "'", "Codigo")) = "" Then
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
    
    phtCliente.Reset
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
    vVieneClientesAlta = ""
           
    txtOtrosDatos(6).Text = GenerarPass
    
    txtAlta(0).Locked = Not True
    KeyPreview = True
    
    txtAlta(0).Text = Val(GenerarDato("SELECT MAX(Codigo) AS UltimoCodigo FROM Clientes", "UltimoCodigo")) + 1
    txtAlta(0).Text = FormatoUltimoCodigo(4, txtAlta(0).Text)

    txtAlta(1).SetFocus
    
    If Err Then GrabarLog "LimpiarCampos", Err.Number & "-" & Err.Description, Me.Name
End Sub
Public Sub ModificarCliente(vIDCliente As Long)
    On Error Resume Next
    
    'MsgBox vIDCliente
    
    Dim rsCliente As New ADODB.Recordset, sqlCliente As String
    
    sqlCliente = "SELECT * FROM Clientes WHERE (idClientes = " & Str(vIDCliente) & ")"
    
    With rsCliente
        Call .Open(sqlCliente, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not (.EOF = True) And Not (.BOF = True) Then
        
            'No Opcionales
            txtAlta(0).Text = .Fields("codigo").Value
            txtAlta(0).Locked = True
        
            MsgBox "Usted está modificando clientes/proveedores." + Chr(13) + "Clic en aceptar para continuar", vbInformation, "Atención"
        
        
            txtAlta(1).Text = .Fields("Nombre").Value
            txtAlta(2).Text = .Fields("RazonSocial").Value
            
            txtAlta(0).Text = .Fields("codigo").Value
        
            'If Not IsNull(.Fields("Foto").Value) = True And Not Trim(.Fields("Foto").Value) = "" Then
            '    BorrarArchivo (App.Path & "\" & .Fields("Codigo").Value & ".dat")
            '    phtCliente.BlobToFile rsCliente!Foto, App.Path & "\" & .Fields("Codigo").Value & ".dat"
            '    Call phtCliente.AbrirFotoDesdeArchivo(App.Path & "\" & .Fields("Codigo").Value & ".dat")
            '    BorrarArchivo (App.Path & "\" & .Fields("Codigo").Value & ".dat")
            '    pbCierraFoto.Visible = True
            'End If

            'Ficha
        
            'MsgBox .Fields("codigo").Value
            
            
            txtFicha(0).Text = EsNulo(.Fields("Direccion").Value)
            txtFicha(1).Text = EsNulo(.Fields("CodigoPostal").Value)
            txtFicha(2).Text = EsNulo(.Fields("Localidad").Value)
            txtFicha(3).Text = EsNulo(.Fields("Provincia").Value)
            txtFicha(4).Text = EsNulo(.Fields("Pais").Value)
            txtFicha(5).Text = EsNulo(.Fields("Telefono").Value)
            txtFicha(6).Text = EsNulo(.Fields("Fax").Value)
            txtFicha(7).Text = EsNulo(.Fields("Celular").Value)
            
            
            vdesvendedor.Text = EsNulo(TraerDato("proveedores", "idProveedores =  " & .Fields("idVendedor2").Value & "", "nombre"))
            vcodvendedor.Text = EsNulo(TraerDato("proveedores", "idProveedores =  " & .Fields("idVendedor2").Value & "", "Codigo"))
        
            idVendedor = .Fields("idVendedor2")
            
            'txtFicha(8).Text = EsNulo(.Fields("idVendedor").Value)
            'txtFicha(9).Text = EsNulo(TraerDato("Empleados", "Codigo =  '" & .Fields("idVendedor").Value & "'", "Nombre"))
        
            txtFicha(10).Text = EsNulo(.Fields("idReparto").Value)
            txtFicha(11).Text = EsNulo(TraerDato("clireparto", "nreparto =  '" & .Fields("idReparto").Value & "'", "descrip"))
        
            txtFicha(12).Text = .Fields("TipoDocumento").Value
            txtFicha(13).Text = EsNulo(TraerDato("TipoDocumentos", "idTipoDocumentos =  '" & .Fields("TipoDocumento").Value & "'", "Tipo"))
            txtFicha(14).Text = .Fields("NroDocumento").Value
            
            'Datos Comerciales
            txtDatosComerciales(0).Text = EsNulo(.Fields("idTipoIva").Value)
            txtDatosComerciales(1).Text = EsNulo(TraerDato("TipoIva", "idTipoIva = '" & .Fields("idTipoIva").Value & "'", "TipoIva"))
            txtDatosComerciales(2).Text = EsNulo(.Fields("Cuit").Value)
            txtDatosComerciales(3).Text = EsNulo(.Fields("idTipoCliente").Value)
            txtDatosComerciales(4).Text = EsNulo(TraerDato("TipoClientes", "idTipoClientes = '" & .Fields("idTipoCliente").Value & "'", "Descripcion"))
            txtDatosComerciales(5).Text = .Fields("idActividad").Value
            txtDatosComerciales(6).Text = EsNulo(TraerDato("Actividades", "idActividades = '" & .Fields("idActividad").Value & "'", "Descripcion"))
            Me.vcboLista.Text = EsNulo(.Fields("idListas").Value)
            'txtDatosComerciales(7).Text = EsNulo(.Fields("idListas").Value)
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
            txtOtrosDatos(6).Text = EsNulo(.Fields("PasswordWeb").Value)
            
            'Observaciones
            txtObservaciones.Text = EsNulo(.Fields("Observaciones").Value)
        
            Call CargarGrillaFacturaAutomatica(.Fields("codigo").Value)
        End If

    End With
    
    If Err Then GrabarLog "ModificarCliente", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub pbCierraFoto_Click()
    On Error Resume Next

    phtCliente.Reset
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

Private Sub txtAlta_Change(Index As Integer)
    On Error Resume Next

    If Index = 1 Then txtAlta(2).Text = txtAlta(1).Text

    If Err Then GrabarLog "txtAlta_Change", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtAlta_GotFocus(Index As Integer)
If Index = 0 Then txtAlta(0) = nuevoCodigo
End Sub

Private Sub txtArticulos_KeyPress(Index As Integer, Keyascii As Integer)
On Error Resume Next
     
    If Index = 2 Then If Keyascii = 13 Then cmdArticulo(3).SetFocus
    
If Err Then GrabarLog "txtArticulos_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub

Function nuevoCodigo() As String
On Error Resume Next
Dim vsql As String

vsql = "select max(CONVERT(codigo,UNSIGNED INTEGER)) as c from clientes"

nuevoCodigo = traerDatos2(vsql, "c", pathDBMySQL) + 1
If Err Then Exit Function
End Function

Private Sub vcodvendedor_Change()
    idVendedor = codigo2id(Me.vcodvendedor.Text)
End Sub

Private Sub vdesvendedor_Change()
   Me.vcodvendedor.Text = Me.vdesvendedor.Tag
End Sub
