VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{63BEADB1-20E1-478A-9B40-DDDAFBF3624F}#1.0#0"; "bsGradientLabel.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmConfiguracion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuracion General del Sistema"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   12075
   ShowInTaskbar   =   0   'False
   Begin Project1.bsGradientLabel bsTitulos 
      Height          =   315
      Left            =   0
      Top             =   6060
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   556
      Caption         =   "Bs titulos"
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Colour1         =   8421504
      Colour2         =   4210752
   End
   Begin XtremeSuiteControls.PushButton cmdAcciones 
      Height          =   420
      Index           =   1
      Left            =   10200
      TabIndex        =   18
      Top             =   5880
      Width           =   1815
      _Version        =   851968
      _ExtentX        =   3201
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Volver/Salir"
      BackColor       =   -2147483638
      Appearance      =   6
      Picture         =   "frmConfiguracion.frx":0000
   End
   Begin XtremeSuiteControls.PushButton cmdAcciones 
      Height          =   420
      Index           =   0
      Left            =   8160
      TabIndex        =   17
      Top             =   5880
      Width           =   1815
      _Version        =   851968
      _ExtentX        =   3201
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Aceptar"
      BackColor       =   -2147483638
      Appearance      =   6
      Picture         =   "frmConfiguracion.frx":6862
   End
   Begin XtremeSuiteControls.TabControl TabConfiguracion 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      _Version        =   851968
      _ExtentX        =   21193
      _ExtentY        =   10610
      _StockProps     =   68
      AllowReorder    =   -1  'True
      Appearance      =   7
      Color           =   32
      PaintManager.Position=   1
      PaintManager.BoldSelected=   -1  'True
      PaintManager.OneNoteColors=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      ItemCount       =   5
      SelectedItem    =   4
      Item(0).Caption =   "Servidores"
      Item(0).ControlCount=   12
      Item(0).Control(0)=   "txtDatosServidor(0)"
      Item(0).Control(1)=   "lblDatosServidor(0)"
      Item(0).Control(2)=   "lblDatosServidor(1)"
      Item(0).Control(3)=   "txtDatosServidor(1)"
      Item(0).Control(4)=   "lblDatosServidor(2)"
      Item(0).Control(5)=   "txtDatosServidor(2)"
      Item(0).Control(6)=   "lblDatosServidor(3)"
      Item(0).Control(7)=   "txtDatosServidor(3)"
      Item(0).Control(8)=   "lblDatosServidor(4)"
      Item(0).Control(9)=   "txtDatosServidor(4)"
      Item(0).Control(10)=   "cboServidorHabilitado"
      Item(0).Control(11)=   "lblDatosServidor(5)"
      Item(1).Caption =   "Empresas"
      Item(1).ControlCount=   20
      Item(1).Control(0)=   "txtDatosEmpresa(0)"
      Item(1).Control(1)=   "txtDatosEmpresa(1)"
      Item(1).Control(2)=   "txtDatosEmpresa(3)"
      Item(1).Control(3)=   "txtDatosEmpresa(4)"
      Item(1).Control(4)=   "txtDatosEmpresa(5)"
      Item(1).Control(5)=   "txtDatosEmpresa(6)"
      Item(1).Control(6)=   "txtDatosEmpresa(7)"
      Item(1).Control(7)=   "txtDatosEmpresa(2)"
      Item(1).Control(8)=   "lblDatosEmpresa(0)"
      Item(1).Control(9)=   "lblDatosEmpresa(1)"
      Item(1).Control(10)=   "lblDatosEmpresa(2)"
      Item(1).Control(11)=   "lblDatosEmpresa(4)"
      Item(1).Control(12)=   "lblDatosEmpresa(5)"
      Item(1).Control(13)=   "lblDatosEmpresa(6)"
      Item(1).Control(14)=   "lblDatosEmpresa(7)"
      Item(1).Control(15)=   "lblDatosEmpresa(3)"
      Item(1).Control(16)=   "txtDatosEmpresa(8)"
      Item(1).Control(17)=   "lblDatosEmpresa(8)"
      Item(1).Control(18)=   "cboEmpresaHabilitada"
      Item(1).Control(19)=   "lblDatosEmpresa(9)"
      Item(2).Caption =   "Usuarios"
      Item(2).ControlCount=   8
      Item(2).Control(0)=   "txtDatosUsuario(0)"
      Item(2).Control(1)=   "lblDatosUsuario(0)"
      Item(2).Control(2)=   "txtDatosUsuario(1)"
      Item(2).Control(3)=   "lblDatosUsuario(1)"
      Item(2).Control(4)=   "txtDatosUsuario(2)"
      Item(2).Control(5)=   "lblDatosUsuario(2)"
      Item(2).Control(6)=   "cboUsuarioHabilitado"
      Item(2).Control(7)=   "lblDatosUsuario(3)"
      Item(3).Caption =   "Grupos && Formularios"
      Item(3).ControlCount=   28
      Item(3).Control(0)=   "rbGrupo(0)"
      Item(3).Control(1)=   "rbGrupo(1)"
      Item(3).Control(2)=   "rbGrupo(2)"
      Item(3).Control(3)=   "rbGrupo(3)"
      Item(3).Control(4)=   "rbGrupo(4)"
      Item(3).Control(5)=   "rbGrupo(5)"
      Item(3).Control(6)=   "rbGrupo(6)"
      Item(3).Control(7)=   "rbGrupo(7)"
      Item(3).Control(8)=   "rbGrupo(8)"
      Item(3).Control(9)=   "rbGrupo(9)"
      Item(3).Control(10)=   "rbGrupo(10)"
      Item(3).Control(11)=   "rbGrupo(11)"
      Item(3).Control(12)=   "rbGrupo(12)"
      Item(3).Control(13)=   "rbGrupo(13)"
      Item(3).Control(14)=   "rbGrupo(14)"
      Item(3).Control(15)=   "rbGrupo(15)"
      Item(3).Control(16)=   "rbGrupo(16)"
      Item(3).Control(17)=   "rbGrupo(17)"
      Item(3).Control(18)=   "chkFormulario(0)"
      Item(3).Control(19)=   "chkFormulario(1)"
      Item(3).Control(20)=   "chkFormulario(2)"
      Item(3).Control(21)=   "chkFormulario(3)"
      Item(3).Control(22)=   "chkFormulario(4)"
      Item(3).Control(23)=   "chkFormulario(5)"
      Item(3).Control(24)=   "chkFormulario(6)"
      Item(3).Control(25)=   "chkFormulario(7)"
      Item(3).Control(26)=   "chkFormulario(8)"
      Item(3).Control(27)=   "chkFormulario(9)"
      Item(4).Caption =   "Asociaciones"
      Item(4).ControlCount=   9
      Item(4).Control(0)=   "lblDatosAsociacion(3)"
      Item(4).Control(1)=   "lblDatosAsociacion(5)"
      Item(4).Control(2)=   "lblDatosAsociacion(4)"
      Item(4).Control(3)=   "cboServidor"
      Item(4).Control(4)=   "cboEmpresa"
      Item(4).Control(5)=   "cboUsuario"
      Item(4).Control(6)=   "cmdAsociaciones(0)"
      Item(4).Control(7)=   "cmdAsociaciones(1)"
      Item(4).Control(8)=   "dgEmpresaAsociada"
      Begin VB.ComboBox cboUsuario 
         Height          =   315
         Left            =   2760
         TabIndex        =   79
         Top             =   1680
         Width           =   2655
      End
      Begin VB.ComboBox cboEmpresa 
         Height          =   315
         Left            =   2760
         TabIndex        =   78
         Top             =   1200
         Width           =   2655
      End
      Begin VB.ComboBox cboServidor 
         Height          =   315
         Left            =   2760
         TabIndex        =   77
         Top             =   720
         Width           =   2655
      End
      Begin XtremeSuiteControls.PushButton cmdAsociaciones 
         Height          =   345
         Index           =   0
         Left            =   3720
         TabIndex        =   75
         Top             =   2160
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Agregar Relacion"
         UseVisualStyle  =   -1  'True
      End
      Begin MSDataGridLib.DataGrid dgEmpresaAsociada 
         Height          =   2175
         Left            =   6120
         TabIndex        =   71
         Top             =   720
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   3836
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
      Begin XtremeSuiteControls.FlatEdit txtDatosEmpresa 
         Height          =   345
         Index           =   1
         Left            =   -67240
         TabIndex        =   2
         Top             =   1200
         Visible         =   0   'False
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
      Begin XtremeSuiteControls.FlatEdit txtDatosEmpresa 
         Height          =   345
         Index           =   3
         Left            =   -67240
         TabIndex        =   3
         Top             =   2160
         Visible         =   0   'False
         Width           =   3855
         _Version        =   851968
         _ExtentX        =   6800
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   4
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosEmpresa 
         Height          =   345
         Index           =   4
         Left            =   -67240
         TabIndex        =   4
         Top             =   2640
         Visible         =   0   'False
         Width           =   3855
         _Version        =   851968
         _ExtentX        =   6800
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   4
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosEmpresa 
         Height          =   345
         Index           =   5
         Left            =   -67240
         TabIndex        =   5
         Top             =   3120
         Visible         =   0   'False
         Width           =   3855
         _Version        =   851968
         _ExtentX        =   6800
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   4
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosEmpresa 
         Height          =   345
         Index           =   6
         Left            =   -67240
         TabIndex        =   6
         Top             =   3600
         Visible         =   0   'False
         Width           =   3855
         _Version        =   851968
         _ExtentX        =   6800
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   4
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosEmpresa 
         Height          =   345
         Index           =   7
         Left            =   -67240
         TabIndex        =   7
         Top             =   4080
         Visible         =   0   'False
         Width           =   3855
         _Version        =   851968
         _ExtentX        =   6800
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   4
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosEmpresa 
         Height          =   345
         Index           =   2
         Left            =   -67240
         TabIndex        =   8
         Top             =   1680
         Visible         =   0   'False
         Width           =   3855
         _Version        =   851968
         _ExtentX        =   6800
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   4
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosServidor 
         Height          =   345
         Index           =   1
         Left            =   -67240
         TabIndex        =   22
         Top             =   1200
         Visible         =   0   'False
         Width           =   3855
         _Version        =   851968
         _ExtentX        =   6800
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   4
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosServidor 
         Height          =   345
         Index           =   0
         Left            =   -67240
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   3855
         _Version        =   851968
         _ExtentX        =   6800
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   4
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosServidor 
         Height          =   345
         Index           =   2
         Left            =   -67240
         TabIndex        =   24
         Top             =   1680
         Visible         =   0   'False
         Width           =   3855
         _Version        =   851968
         _ExtentX        =   6800
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   4
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosServidor 
         Height          =   345
         Index           =   3
         Left            =   -67240
         TabIndex        =   26
         Top             =   2160
         Visible         =   0   'False
         Width           =   3855
         _Version        =   851968
         _ExtentX        =   6800
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   4
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosServidor 
         Height          =   345
         Index           =   4
         Left            =   -67240
         TabIndex        =   28
         Top             =   2640
         Visible         =   0   'False
         Width           =   3855
         _Version        =   851968
         _ExtentX        =   6800
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   4
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosUsuario 
         Height          =   345
         Index           =   1
         Left            =   -67240
         TabIndex        =   31
         Top             =   1200
         Visible         =   0   'False
         Width           =   3855
         _Version        =   851968
         _ExtentX        =   6800
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   4
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosUsuario 
         Height          =   345
         Index           =   2
         Left            =   -67240
         TabIndex        =   33
         Top             =   1680
         Visible         =   0   'False
         Width           =   3855
         _Version        =   851968
         _ExtentX        =   6800
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   4
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosEmpresa 
         Height          =   345
         Index           =   8
         Left            =   -67240
         TabIndex        =   35
         Top             =   4560
         Visible         =   0   'False
         Width           =   3855
         _Version        =   851968
         _ExtentX        =   6800
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   4
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosEmpresa 
         Height          =   345
         Index           =   0
         Left            =   -67240
         TabIndex        =   1
         Top             =   720
         Visible         =   0   'False
         Width           =   3855
         _Version        =   851968
         _ExtentX        =   6800
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   4
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosUsuario 
         Height          =   345
         Index           =   0
         Left            =   -67240
         TabIndex        =   29
         Top             =   720
         Visible         =   0   'False
         Width           =   3855
         _Version        =   851968
         _ExtentX        =   6800
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   4
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboUsuarioHabilitado 
         Height          =   315
         Left            =   -67240
         TabIndex        =   37
         Top             =   2160
         Visible         =   0   'False
         Width           =   1575
         _Version        =   851968
         _ExtentX        =   2778
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Sorted          =   -1  'True
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         EnableMarkup    =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboEmpresaHabilitada 
         Height          =   315
         Left            =   -67240
         TabIndex        =   38
         Top             =   5040
         Visible         =   0   'False
         Width           =   1575
         _Version        =   851968
         _ExtentX        =   2778
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Sorted          =   -1  'True
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         EnableMarkup    =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboServidorHabilitado 
         Height          =   315
         Left            =   -67240
         TabIndex        =   39
         Top             =   3120
         Visible         =   0   'False
         Width           =   1575
         _Version        =   851968
         _ExtentX        =   2778
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Sorted          =   -1  'True
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         EnableMarkup    =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton rbGrupo 
         Height          =   255
         Index           =   0
         Left            =   -69400
         TabIndex        =   43
         Top             =   1080
         Visible         =   0   'False
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         BackColor       =   14737632
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton rbGrupo 
         Height          =   255
         Index           =   1
         Left            =   -69400
         TabIndex        =   44
         Top             =   1320
         Visible         =   0   'False
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         BackColor       =   14737632
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton rbGrupo 
         Height          =   255
         Index           =   2
         Left            =   -69400
         TabIndex        =   45
         Top             =   1560
         Visible         =   0   'False
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         BackColor       =   12632256
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton rbGrupo 
         Height          =   255
         Index           =   3
         Left            =   -69400
         TabIndex        =   46
         Top             =   1800
         Visible         =   0   'False
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton rbGrupo 
         Height          =   255
         Index           =   4
         Left            =   -69400
         TabIndex        =   47
         Top             =   2040
         Visible         =   0   'False
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton rbGrupo 
         Height          =   255
         Index           =   5
         Left            =   -69400
         TabIndex        =   48
         Top             =   2280
         Visible         =   0   'False
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton rbGrupo 
         Height          =   255
         Index           =   6
         Left            =   -69400
         TabIndex        =   49
         Top             =   2520
         Visible         =   0   'False
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton rbGrupo 
         Height          =   255
         Index           =   7
         Left            =   -69400
         TabIndex        =   50
         Top             =   2760
         Visible         =   0   'False
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton rbGrupo 
         Height          =   255
         Index           =   8
         Left            =   -69400
         TabIndex        =   51
         Top             =   3000
         Visible         =   0   'False
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton rbGrupo 
         Height          =   255
         Index           =   9
         Left            =   -69400
         TabIndex        =   52
         Top             =   3240
         Visible         =   0   'False
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton rbGrupo 
         Height          =   255
         Index           =   10
         Left            =   -69400
         TabIndex        =   53
         Top             =   3480
         Visible         =   0   'False
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton rbGrupo 
         Height          =   255
         Index           =   11
         Left            =   -69400
         TabIndex        =   54
         Top             =   3720
         Visible         =   0   'False
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         BackColor       =   12632256
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton rbGrupo 
         Height          =   255
         Index           =   12
         Left            =   -69400
         TabIndex        =   55
         Top             =   3960
         Visible         =   0   'False
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton rbGrupo 
         Height          =   255
         Index           =   13
         Left            =   -69400
         TabIndex        =   56
         Top             =   4200
         Visible         =   0   'False
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton rbGrupo 
         Height          =   255
         Index           =   14
         Left            =   -69400
         TabIndex        =   57
         Top             =   4440
         Visible         =   0   'False
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton rbGrupo 
         Height          =   255
         Index           =   15
         Left            =   -69400
         TabIndex        =   58
         Top             =   4680
         Visible         =   0   'False
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton rbGrupo 
         Height          =   255
         Index           =   16
         Left            =   -69400
         TabIndex        =   59
         Top             =   4920
         Visible         =   0   'False
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton rbGrupo 
         Height          =   255
         Index           =   17
         Left            =   -69400
         TabIndex        =   60
         Top             =   5160
         Visible         =   0   'False
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox chkFormulario 
         Height          =   255
         Index           =   0
         Left            =   -64240
         TabIndex        =   61
         Top             =   1080
         Visible         =   0   'False
         Width           =   2535
         _Version        =   851968
         _ExtentX        =   4471
         _ExtentY        =   450
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox chkFormulario 
         Height          =   255
         Index           =   1
         Left            =   -64240
         TabIndex        =   62
         Top             =   1320
         Visible         =   0   'False
         Width           =   2535
         _Version        =   851968
         _ExtentX        =   4471
         _ExtentY        =   450
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox chkFormulario 
         Height          =   255
         Index           =   2
         Left            =   -64240
         TabIndex        =   63
         Top             =   1560
         Visible         =   0   'False
         Width           =   2535
         _Version        =   851968
         _ExtentX        =   4471
         _ExtentY        =   450
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox chkFormulario 
         Height          =   255
         Index           =   3
         Left            =   -64240
         TabIndex        =   64
         Top             =   1800
         Visible         =   0   'False
         Width           =   2535
         _Version        =   851968
         _ExtentX        =   4471
         _ExtentY        =   450
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox chkFormulario 
         Height          =   255
         Index           =   4
         Left            =   -64240
         TabIndex        =   65
         Top             =   2040
         Visible         =   0   'False
         Width           =   2535
         _Version        =   851968
         _ExtentX        =   4471
         _ExtentY        =   450
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox chkFormulario 
         Height          =   255
         Index           =   5
         Left            =   -64240
         TabIndex        =   66
         Top             =   2280
         Visible         =   0   'False
         Width           =   2535
         _Version        =   851968
         _ExtentX        =   4471
         _ExtentY        =   450
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox chkFormulario 
         Height          =   255
         Index           =   6
         Left            =   -64240
         TabIndex        =   67
         Top             =   2520
         Visible         =   0   'False
         Width           =   2535
         _Version        =   851968
         _ExtentX        =   4471
         _ExtentY        =   450
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox chkFormulario 
         Height          =   255
         Index           =   7
         Left            =   -64240
         TabIndex        =   68
         Top             =   2760
         Visible         =   0   'False
         Width           =   2535
         _Version        =   851968
         _ExtentX        =   4471
         _ExtentY        =   450
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox chkFormulario 
         Height          =   255
         Index           =   8
         Left            =   -64240
         TabIndex        =   69
         Top             =   3000
         Visible         =   0   'False
         Width           =   2535
         _Version        =   851968
         _ExtentX        =   4471
         _ExtentY        =   450
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox chkFormulario 
         Height          =   255
         Index           =   9
         Left            =   -64240
         TabIndex        =   70
         Top             =   3240
         Visible         =   0   'False
         Width           =   2535
         _Version        =   851968
         _ExtentX        =   4471
         _ExtentY        =   450
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton cmdAsociaciones 
         Height          =   345
         Index           =   1
         Left            =   10200
         TabIndex        =   76
         Top             =   3000
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Eliminar Relacion"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label lblDatosAsociacion 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> Seleccione Empresa :"
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   74
         Top             =   1245
         Width           =   2205
      End
      Begin VB.Label lblDatosAsociacion 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> Seleccione Usuario :"
         Height          =   195
         Index           =   5
         Left            =   360
         TabIndex        =   73
         Top             =   1725
         Width           =   2205
      End
      Begin VB.Label lblDatosAsociacion 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> Seleccione Servidor :"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   72
         Top             =   765
         Width           =   2205
      End
      Begin VB.Label lblDatosUsuario 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> Habilitado  :"
         Height          =   195
         Index           =   3
         Left            =   -69640
         TabIndex        =   42
         Top             =   2205
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Label lblDatosEmpresa 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> Habilitada :"
         Height          =   195
         Index           =   9
         Left            =   -69640
         TabIndex        =   41
         Top             =   5085
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Label lblDatosServidor 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   ">  Habilitado :"
         Height          =   195
         Index           =   5
         Left            =   -69640
         TabIndex        =   40
         Top             =   3165
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Label lblDatosEmpresa 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> Website :"
         Height          =   195
         Index           =   8
         Left            =   -69640
         TabIndex        =   36
         Top             =   4605
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Label lblDatosUsuario 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> Password de Usuario  :"
         Height          =   195
         Index           =   2
         Left            =   -69640
         TabIndex        =   34
         Top             =   1725
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Label lblDatosUsuario 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> Nombre de la Usuario  :"
         Height          =   195
         Index           =   1
         Left            =   -69640
         TabIndex        =   32
         Top             =   1245
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Label lblDatosUsuario 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> ID de Usuario  :"
         Height          =   195
         Index           =   0
         Left            =   -69640
         TabIndex        =   30
         Top             =   765
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Label lblDatosServidor 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   ">  Contraseña :"
         Height          =   195
         Index           =   4
         Left            =   -69640
         TabIndex        =   27
         Top             =   2685
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Label lblDatosServidor 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> Nombre de Usuario :"
         Height          =   195
         Index           =   3
         Left            =   -69640
         TabIndex        =   25
         Top             =   2205
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Label lblDatosServidor 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> Direccion del Servidor  :"
         Height          =   195
         Index           =   2
         Left            =   -69640
         TabIndex        =   23
         Top             =   1725
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Label lblDatosServidor 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> Nombre de la Servidor  :"
         Height          =   195
         Index           =   1
         Left            =   -69640
         TabIndex        =   21
         Top             =   1245
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Label lblDatosServidor 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> ID del Servidor  :"
         Height          =   195
         Index           =   0
         Left            =   -69640
         TabIndex        =   20
         Top             =   765
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Label lblDatosEmpresa 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> ID de la Empresa  :"
         Height          =   195
         Index           =   0
         Left            =   -69640
         TabIndex        =   16
         Top             =   765
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Label lblDatosEmpresa 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> Nombre de la Empresa  :"
         Height          =   195
         Index           =   1
         Left            =   -69640
         TabIndex        =   15
         Top             =   1260
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Label lblDatosEmpresa 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> Alias de la Empresa  :"
         Height          =   195
         Index           =   2
         Left            =   -69640
         TabIndex        =   14
         Top             =   1725
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Label lblDatosEmpresa 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> Direccion :"
         Height          =   195
         Index           =   4
         Left            =   -69640
         TabIndex        =   13
         Top             =   2685
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Label lblDatosEmpresa 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> Localidad :"
         Height          =   195
         Index           =   5
         Left            =   -69640
         TabIndex        =   12
         Top             =   3165
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Label lblDatosEmpresa 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> Telefono :"
         Height          =   195
         Index           =   6
         Left            =   -69640
         TabIndex        =   11
         Top             =   3645
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Label lblDatosEmpresa 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> E-Mail :"
         Height          =   195
         Index           =   7
         Left            =   -69640
         TabIndex        =   10
         Top             =   4125
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Label lblDatosEmpresa 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "> C.U.I.T :"
         Height          =   195
         Index           =   3
         Left            =   -69640
         TabIndex        =   9
         Top             =   2205
         Visible         =   0   'False
         Width           =   2205
      End
   End
End
Attribute VB_Name = "frmConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim rsEmpresaAsociada As ADODB.Recordset
Private Sub cboEmpresa_Click()
On Error Resume Next

    cboEmpresa.Tag = Trim(TraerDato("Empresas", "Empresa = '" & Trim(cboEmpresa.Text) & "'", "idEmpresas", PathDBConfig))

If Err Then GrabarLog "cboEmpresa_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cboEmpresa_GotFocus()
On Error Resume Next

    Call CargarCombo("Empresas", "Empresa", cboEmpresa, True, , PathDBConfig)

If Err Then GrabarLog "cboEmpresa_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cboUsuario_GotFocus()
On Error Resume Next

    Call CargarCombo("Usuarios", "Usuario", cboUsuario, True, , PathDBConfig)

If Err Then GrabarLog "cboUsuario_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cboUsuario_Click()
On Error Resume Next

    cboUsuario.Tag = Trim(TraerDato("Usuarios", "Usuario = '" & Trim(cboUsuario.Text) & "'", "idUsuarios", PathDBConfig))

If Err Then GrabarLog "cboUsuario_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cboServidor_GotFocus()
On Error Resume Next

    Call CargarCombo("Servidor", "Servidor", cboServidor, True, , PathDBConfig)

If Err Then GrabarLog "cboServidor_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cboServidor_Click()
On Error Resume Next

    cboServidor.Tag = Trim(TraerDato("Servidor", "Servidor = '" & Trim(cboServidor.Text) & "'", "idServidor", PathDBConfig))

If Err Then GrabarLog "cboServidor_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cboServidorHabilitado_GotFocus()
On Error Resume Next

    cboServidorHabilitad.AddItem ("SI")
    cboServidorHabilitad.AddItem ("No")
    
If Err Then GrabarLog "cboServidorHabilitad_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cboEmpresaHabilitada_GotFocus()
On Error Resume Next

    cboEmpresaHabilitada.AddItem ("SI")
    cboEmpresaHabilitada.AddItem ("No")
    
If Err Then GrabarLog "cboEmpresaHabilitada_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cboUsuarioHabilitado_GotFocus()
On Error Resume Next

    cboUsuarioHabilitado.AddItem ("SI")
    cboUsuarioHabilitado.AddItem ("No")
    
If Err Then GrabarLog "cboUsuarioHabilitado_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdAcciones_Click(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 0
        
        Case 1
            SalirVolver
        
    
    End Select

If Err Then GrabarLog "", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub SalirVolver()
On Error Resume Next

If Err Then GrabarLog "SalirVolver", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Load()
On Error Resume Next

    CargarGrupos
    
    CargarGrillaRelacion
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CargarGrillaRelacion()
On Error Resume Next

    Set rsEmpresaAsociada = New ADODB.Recordset
    Dim sqlEmpresaAsociada As String
    
    sqlEmpresaAsociada = "SELECT EmpresasAsociadas.idEmpresasAsociadas, Servidor.Servidor, Empresas.Empresa, Usuarios.Usuario FROM ((Empresas INNER JOIN EmpresasAsociadas ON Empresas.idEmpresas = EmpresasAsociadas.idEmpresas) INNER JOIN Servidor ON EmpresasAsociadas.idServidor = Servidor.idServidor) INNER JOIN Usuarios ON EmpresasAsociadas.idUsuarios = Usuarios.idUsuarios;"
    
    With rsEmpresaAsociada
        If .State = 1 Then .Close
        
        .CursorLocation = adUseClient
        
        Call .Open(sqlEmpresaAsociada, PathDBConfig, adOpenStatic, adLockReadOnly)
        
    End With
    
    With dgEmpresaAsociada
        Set .DataSource = rsEmpresaAsociada

        .HeadLines = 2
        
        .Columns(0).Width = 500
        .Columns(0).Caption = "ID"
        
        .Columns(1).Width = 1500
        .Columns(1).Caption = "Servidor"
        
        .Columns(2).Width = 1500
        .Columns(2).Caption = "Empresa"
        
        .Columns(3).Width = 1500
        .Columns(3).Caption = "Usuario"

    End With
    
If Err Then GrabarLog "CargarGrillaRelacion", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub LimpiarCampos(Index As Integer)
On Error Resume Next

    i = 0
            
    Select Case Index
    
        Case 0

        Case 1
            For i = 0 To txtDatosEmpresa.Count - 1
                txtDatosEmpresa(i).Text = ""
            Next
        Case 2
        
        Case 3
        
        Case 4
    
    End Select


If Err Then GrabarLog "LimpiarCampos", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CargarGrupos()
On Error Resume Next

    i = 0
    
    Dim rsFormularioGrupo As New ADODB.Recordset, sqlFormularioGrupo As String
    
    sqlFormularioGrupo = "SELECT * FROM FormularioGrupo ORDER BY idFormularioGrupo"
    
    With rsFormularioGrupo
        .CursorLocation = adUseClient
        
        Call .Open(sqlFormularioGrupo, PathDBConfig, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then .MoveFirst
    
        Do Until .EOF = True
            If .Fields("Habilitado").Value = "N" Then
                rbGrupo(i).ForeColor = vbRed
            End If
            
            rbGrupo(i).Caption = .Fields("FormularioGrupo").Value
            rbGrupo(i).Tag = .Fields("idFormularioGrupo").Value
            rbGrupo(i).Visible = True
            
            i = i + 1
            .MoveNext
        Loop
        
    End With
    
    sqlFormularioGrupo = ""
    
    If rsFormularioGrupo.State = 1 Then
        rsFormularioGrupo.Close
        Set rsFormularioGrupo = Nothing
    End If

If Err Then GrabarLog "CargarGrupos", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub GuardarEmpresa()
On Error Resume Next

    If Trim(txtDatosEmpresa(0).Text) = "" Or Trim(txtDatosEmpresa(1).Text) = "" Then
        MsgBox "Debe cargar la Empresa y/o Alias!!!", vbExclamation, "Mensaje ..."
        Exit Sub
    End If
    
    Dim vIdEmpresas As Long
    
    i = 0
    
    vIdEmpresas = Trim(TraerDato("Empresas", "Empresa = '" & Trim(txtDatosEmpresa(0).Text) & "' OR alias = '" & Trim(txtDatosEmpresa(1).Text) & "'", "idEmpresas", PathDBConfig))
    
    If Not Val(vIdEmpresas) = 0 Then
        MsgBox "La Empresa ya ha sido cargada previamente!!!", vbExclamation, "Mensaje ..."
        Exit Sub
    Else
        Call EjecutarScript("INSERT INTO Empresas (Empresa, Alias, Habilitada) VALUES ('" & Trim(txtDatosEmpresa(0).Text) & "','" & Trim(txtDatosEmpresa(1).Text) & "','S')", PathDBConfig)
        vIdEmpresas = Trim(TraerDato("Empresas", "(Empresa = '" & Trim(txtDatosEmpresa(0).Text) & "') AND (alias = '" & Trim(txtDatosEmpresa(1).Text) & "')", "idEmpresas", PathDBConfig))
    
    
        Call CopiarArchivo(vConfigGral.vDireccionDB & "db.new", vConfigGral.vDireccionDB & (txtDatosEmpresa(1).Text) & ".mdb", True)
    End If
    
    Dim rsDatosEmpresas As New ADODB.Recordset, sqlDatosEmpresas As String
    
    sqlDatosEmpresas = "SELECT * FROM DatosEmpresas WHERE (idEmpresas  = " & vIdEmpresas & ")"
    
    With rsDatosEmpresas
        .CursorLocation = adUseClient
        Call .Open(sqlDatosEmpresas, PathDBConfig, adOpenStatic, adLockPessimistic)
        
        .AddNew
        
        .Fields("idEmpresas").Value = vIdEmpresas
        
        For i = 2 To txtDatosEmpresa.Count - 1
            If Not Trim(txtDatosEmpresa(i).Text) = "" Then .Fields(i).Value = txtDatosEmpresa(i).Text
        Next
        
        .Update
    
    End With
    
    LimpiarCampos (1)
   
   sqlDatosEmpresas = ""
   
    If rsDatosEmpresas.State = 1 Then
        rsDatosEmpresas.Close
        Set rsDatosEmpresas = Nothing
    End If
    
If Err Then GrabarLog "GuardarEmpresa", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdAsociaciones_Click(Index As Integer)
On Error Resume Next
    
    Select Case Index
    
        Case 0
            AgregarRelacion
            CargarGrillaRelacion
            'LimpiarRelacion
        Case 1
            With rsEmpresaAsociada
                If .EOF = False And .BOF = False Then
                    Call BorrarBase("EmpresasAsociadas WHERE (idEmpresasAsociadas = " & .Fields(0).Value & ")", PathDBConfig)
                    MsgBox "Registro Borrado Correctamente", vbInformation, "Mensaje ..."
                    CargarGrillaRelacion
                End If
            End With
        Case 2
    
    End Select

If Err Then GrabarLog "cmdAsociaciones_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub AgregarRelacion()
On Error Resume Next

    If cboServidor.Tag = "" Or cboEmpresa.Tag = "" Or cboUsuario.Tag = "" Then
        MsgBox "Debe completar correctamente los 3 Parametros requeridos 'Servidor-Empresa-Usuario'", vbExclamation, "Mensaje ..."
        Exit Sub
    End If
    
    Dim vRelacionExiste As Long
    
    vRelacionExiste = Val(TraerDato("EmpresasAsociadas", "idServidor = " & Val(cboServidor.Tag) & " AND (idUsuarios = " & Val(cboUsuario.Tag) & ") AND (idEmpresas = " & Val(cboEmpresa.Tag) & ")", "idEmpresasAsociadas", PathDBConfig))
    
    If vRelacionExiste = 0 Then
        Call EjecutarScript("INSERT INTO EmpresasAsociadas (idServidor, idUsuarios, idEmpresas) VALUES (" & Val(cboServidor.Tag) & "," & Val(cboUsuario.Tag) & "," & Val(cboEmpresa.Tag) & ")", PathDBConfig)
    Else
        MsgBox "Existe esta Relacion previamente cargada : " & vbCrLf & "Servidor: " & cboServidor.Text & vbCrLf & "Empresa: " & cboEmpresa.Text & vbCrLf & "Usuario: " & cboUsuario.Text & "", vbExclamation, "Mensaje ..."
        Exit Sub
    End If
    
If Err Then GrabarLog "AgregarRelacion", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub rbGrupo_Click(Index As Integer)
On Error Resume Next

If Err Then GrabarLog "rbGrupo_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub TabConfiguracion_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error Resume Next

    With bsTitulos
    
        Select Case Item.Index
    
            Case 0
                .Caption = "Agregar Nuevo Servidor"
            Case 1
                .Caption = "Agregar Nueva Empresa"
            Case 2
                .Caption = "Agregar Nuevo Usuario"
            Case 3
                .Caption = "Seleccionar Formularios Activos"
            Case 4
                .Caption = "Asociar Servidor - Empresa - Usuario"
            Case 5
                'Por las dudas
    
        End Select
    
        .BorderStyle = Etched
        .Top = 120
        .Left = 310
        .ZOrder 0
    
    End With
    
If Err Then GrabarLog "TabConfiguracion_SelectedChanged", Err.Number & " " & Err.Description, Me.Name
End Sub
