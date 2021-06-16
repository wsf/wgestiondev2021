VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmActualizacionPrecio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizacion de Precio"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   12855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   12855
   Begin XtremeSuiteControls.TabControl tab 
      Height          =   6555
      Left            =   30
      TabIndex        =   13
      Top             =   60
      Width           =   12855
      _Version        =   851968
      _ExtentX        =   22675
      _ExtentY        =   11562
      _StockProps     =   68
      PaintManager.BoldSelected=   -1  'True
      ItemCount       =   4
      SelectedItem    =   2
      Item(0).Caption =   "Paso 1: Filtrar artículos"
      Item(0).ControlCount=   30
      Item(0).Control(0)=   "txtFiltro(0)"
      Item(0).Control(1)=   "txtFiltro(2)"
      Item(0).Control(2)=   "pbCarga(0)"
      Item(0).Control(3)=   "txtFiltro(1)"
      Item(0).Control(4)=   "txtFiltro(4)"
      Item(0).Control(5)=   "pbCarga(1)"
      Item(0).Control(6)=   "txtFiltro(3)"
      Item(0).Control(7)=   "lblActualizar(5)"
      Item(0).Control(8)=   "lblActualizar(4)"
      Item(0).Control(9)=   "lblActualizar(3)"
      Item(0).Control(10)=   "chkAbroArticulo"
      Item(0).Control(11)=   "txtFiltro(6)"
      Item(0).Control(12)=   "pbCarga(2)"
      Item(0).Control(13)=   "txtFiltro(10)"
      Item(0).Control(14)=   "pbCarga(4)"
      Item(0).Control(15)=   "txtFiltro(5)"
      Item(0).Control(16)=   "txtFiltro(9)"
      Item(0).Control(17)=   "txtFiltro(8)"
      Item(0).Control(18)=   "pbCarga(3)"
      Item(0).Control(19)=   "txtFiltro(7)"
      Item(0).Control(20)=   "txtFiltro(12)"
      Item(0).Control(21)=   "pbCarga(5)"
      Item(0).Control(22)=   "txtFiltro(11)"
      Item(0).Control(23)=   "lblActualizar(9)"
      Item(0).Control(24)=   "lblActualizar(7)"
      Item(0).Control(25)=   "lblActualizar(8)"
      Item(0).Control(26)=   "lblActualizar(6)"
      Item(0).Control(27)=   "PBAccionesActualizar(1)"
      Item(0).Control(28)=   "chkNot"
      Item(0).Control(29)=   "Label7"
      Item(1).Caption =   "Paso 2: Ingresar valores a cambiar"
      Item(1).ControlCount=   55
      Item(1).Control(0)=   "vPorcentajeCosto"
      Item(1).Control(1)=   "p(5)"
      Item(1).Control(2)=   "p(4)"
      Item(1).Control(3)=   "p(3)"
      Item(1).Control(4)=   "p(2)"
      Item(1).Control(5)=   "p(1)"
      Item(1).Control(6)=   "chkTlista"
      Item(1).Control(7)=   "txtPorcentaje"
      Item(1).Control(8)=   "cboOperacion"
      Item(1).Control(9)=   "cboLista"
      Item(1).Control(10)=   "lblActualizar(10)"
      Item(1).Control(11)=   "lblListaDe"
      Item(1).Control(12)=   "lblIngLos"
      Item(1).Control(13)=   "lblActualizar(0)"
      Item(1).Control(14)=   "lblActualizar(2)"
      Item(1).Control(15)=   "Label1"
      Item(1).Control(16)=   "Label2"
      Item(1).Control(17)=   "Label3"
      Item(1).Control(18)=   "Label4"
      Item(1).Control(19)=   "Label5"
      Item(1).Control(20)=   "pbCarga(6)"
      Item(1).Control(21)=   "lblActualizar(11)"
      Item(1).Control(22)=   "pbCarga(7)"
      Item(1).Control(23)=   "pbCarga(8)"
      Item(1).Control(24)=   "lblActualizar(12)"
      Item(1).Control(25)=   "lblActualizar(13)"
      Item(1).Control(26)=   "Label6"
      Item(1).Control(27)=   "Label8"
      Item(1).Control(28)=   "pbCarga(9)"
      Item(1).Control(29)=   "lblActualizar(1)"
      Item(1).Control(30)=   "vSubRubroDescipcion"
      Item(1).Control(31)=   "vSubRubrocodigo"
      Item(1).Control(32)=   "vproveedorDescripcion"
      Item(1).Control(33)=   "vRubrocodigo"
      Item(1).Control(34)=   "vproveedorCodigo"
      Item(1).Control(35)=   "vivaDescripcion"
      Item(1).Control(36)=   "vivaValor"
      Item(1).Control(37)=   "vRubroDescripcion"
      Item(1).Control(38)=   "Label10"
      Item(1).Control(39)=   "Label9"
      Item(1).Control(40)=   "PBAccionesActualizar(0)"
      Item(1).Control(41)=   "vppcosto"
      Item(1).Control(42)=   "GroupBox1"
      Item(1).Control(43)=   "GroupBox2"
      Item(1).Control(44)=   "GroupBox3"
      Item(1).Control(45)=   "GroupBox4"
      Item(1).Control(46)=   "GroupBox5"
      Item(1).Control(47)=   "cbp(1)"
      Item(1).Control(48)=   "cbp(2)"
      Item(1).Control(49)=   "cbp(3)"
      Item(1).Control(50)=   "cbp(4)"
      Item(1).Control(51)=   "cbp(5)"
      Item(1).Control(52)=   "Line1"
      Item(1).Control(53)=   "chkfijo"
      Item(1).Control(54)=   "Line2"
      Item(2).Caption =   "Paso 3: Ejecutar los cambios"
      Item(2).ControlCount=   4
      Item(2).Control(0)=   "ListBox1"
      Item(2).Control(1)=   "Barra"
      Item(2).Control(2)=   "g"
      Item(2).Control(3)=   "PushButton1"
      Item(3).Caption =   "Paso 4: Ver los cambios y resultados"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "log"
      Begin XtremeSuiteControls.CheckBox cbp 
         Height          =   375
         Index           =   1
         Left            =   -66670
         TabIndex        =   87
         Top             =   2880
         Visible         =   0   'False
         Width           =   2205
         _Version        =   851968
         _ExtentX        =   3889
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Ponerlo igual que el costo"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkNot 
         Height          =   255
         Left            =   -69820
         TabIndex        =   86
         Top             =   480
         Visible         =   0   'False
         Width           =   9945
         _Version        =   851968
         _ExtentX        =   17542
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Filtrar los artículos que  NO cumplen con estas condiciones."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   225
         Left            =   -69850
         TabIndex        =   81
         Top             =   1950
         Visible         =   0   'False
         Width           =   6795
         _Version        =   851968
         _ExtentX        =   11986
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "2.2. Listas: "
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   4
         BorderStyle     =   1
      End
      Begin XtremeSuiteControls.FlatEdit vppcosto 
         Height          =   345
         Left            =   -65620
         TabIndex        =   1
         Top             =   1230
         Visible         =   0   'False
         Width           =   2505
         _Version        =   851968
         _ExtentX        =   4419
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   345
         Left            =   120
         TabIndex        =   80
         Top             =   5760
         Width           =   12585
         _Version        =   851968
         _ExtentX        =   22199
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Aplicar cambios"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmActualizacionPrecio.frx":0000
      End
      Begin XtremeSuiteControls.ListBox ListBox1 
         Height          =   105
         Left            =   60
         TabIndex        =   55
         Top             =   5580
         Width           =   12645
         _Version        =   851968
         _ExtentX        =   22304
         _ExtentY        =   185
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin VB.ListBox log 
         Height          =   5715
         Left            =   -69910
         TabIndex        =   50
         Top             =   450
         Visible         =   0   'False
         Width           =   12495
      End
      Begin VB.CheckBox chkTlista 
         Caption         =   "Todas las listas de precios"
         Height          =   345
         Left            =   -69550
         TabIndex        =   42
         Top             =   1590
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.TextBox p 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   -68200
         TabIndex        =   2
         Top             =   2940
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox p 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   2
         Left            =   -68200
         TabIndex        =   3
         Top             =   3300
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox p 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   3
         Left            =   -68200
         TabIndex        =   4
         Top             =   3660
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox p 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   4
         Left            =   -68200
         TabIndex        =   5
         Top             =   4020
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox p 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   5
         Left            =   -68200
         TabIndex        =   6
         Top             =   4350
         Visible         =   0   'False
         Width           =   1335
      End
      Begin XtremeSuiteControls.FlatEdit txtFiltro 
         Height          =   315
         Index           =   0
         Left            =   -67570
         TabIndex        =   14
         Top             =   870
         Visible         =   0   'False
         Width           =   7755
         _Version        =   851968
         _ExtentX        =   13679
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFiltro 
         Height          =   315
         Index           =   2
         Left            =   -66250
         TabIndex        =   15
         Top             =   1230
         Visible         =   0   'False
         Width           =   2175
         _Version        =   851968
         _ExtentX        =   3836
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   0
         Left            =   -66730
         TabIndex        =   16
         Tag             =   "SubRubroD"
         Top             =   1230
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFiltro 
         Height          =   315
         Index           =   1
         Left            =   -67570
         TabIndex        =   17
         Top             =   1230
         Visible         =   0   'False
         Width           =   735
         _Version        =   851968
         _ExtentX        =   1296
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFiltro 
         Height          =   315
         Index           =   4
         Left            =   -62005
         TabIndex        =   18
         Top             =   1230
         Visible         =   0   'False
         Width           =   2175
         _Version        =   851968
         _ExtentX        =   3836
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   1
         Left            =   -62410
         TabIndex        =   19
         Tag             =   "SubRubroH"
         Top             =   1230
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFiltro 
         Height          =   315
         Index           =   3
         Left            =   -63250
         TabIndex        =   20
         Top             =   1230
         Visible         =   0   'False
         Width           =   735
         _Version        =   851968
         _ExtentX        =   1296
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkAbroArticulo 
         Height          =   375
         Left            =   -69760
         TabIndex        =   24
         Top             =   2910
         Visible         =   0   'False
         Width           =   12405
         _Version        =   851968
         _ExtentX        =   21881
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Abrir Formulario de Articulo y Mostrar los resultados del Filtro"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFiltro 
         Height          =   315
         Index           =   6
         Left            =   -66250
         TabIndex        =   25
         Top             =   1620
         Visible         =   0   'False
         Width           =   2175
         _Version        =   851968
         _ExtentX        =   3836
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   2
         Left            =   -66730
         TabIndex        =   26
         Tag             =   "RubroD"
         Top             =   1620
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFiltro 
         Height          =   315
         Index           =   10
         Left            =   -66250
         TabIndex        =   27
         Top             =   1980
         Visible         =   0   'False
         Width           =   2175
         _Version        =   851968
         _ExtentX        =   3836
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   4
         Left            =   -66730
         TabIndex        =   28
         Tag             =   "ProveedorD"
         Top             =   1980
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFiltro 
         Height          =   315
         Index           =   5
         Left            =   -67570
         TabIndex        =   29
         Top             =   1620
         Visible         =   0   'False
         Width           =   735
         _Version        =   851968
         _ExtentX        =   1296
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFiltro 
         Height          =   315
         Index           =   9
         Left            =   -67570
         TabIndex        =   30
         Top             =   1980
         Visible         =   0   'False
         Width           =   735
         _Version        =   851968
         _ExtentX        =   1296
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFiltro 
         Height          =   315
         Index           =   8
         Left            =   -62005
         TabIndex        =   31
         Top             =   1620
         Visible         =   0   'False
         Width           =   2175
         _Version        =   851968
         _ExtentX        =   3836
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   3
         Left            =   -62410
         TabIndex        =   32
         Tag             =   "RubroH"
         Top             =   1620
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFiltro 
         Height          =   315
         Index           =   7
         Left            =   -63250
         TabIndex        =   33
         Top             =   1620
         Visible         =   0   'False
         Width           =   735
         _Version        =   851968
         _ExtentX        =   1296
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFiltro 
         Height          =   315
         Index           =   12
         Left            =   -62005
         TabIndex        =   34
         Top             =   1980
         Visible         =   0   'False
         Width           =   2175
         _Version        =   851968
         _ExtentX        =   3836
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   5
         Left            =   -62410
         TabIndex        =   35
         Tag             =   "ProveedorH"
         Top             =   1980
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFiltro 
         Height          =   315
         Index           =   11
         Left            =   -63250
         TabIndex        =   36
         Top             =   1980
         Visible         =   0   'False
         Width           =   735
         _Version        =   851968
         _ExtentX        =   1296
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit vPorcentajeCosto 
         Height          =   315
         Left            =   -65620
         TabIndex        =   41
         Top             =   1590
         Visible         =   0   'False
         Width           =   1575
         _Version        =   851968
         _ExtentX        =   2778
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtPorcentaje 
         Height          =   315
         Left            =   -60070
         TabIndex        =   43
         Top             =   720
         Visible         =   0   'False
         Width           =   3045
         _Version        =   851968
         _ExtentX        =   5371
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboOperacion 
         Height          =   315
         Left            =   -66670
         TabIndex        =   0
         Top             =   570
         Visible         =   0   'False
         Width           =   3525
         _Version        =   851968
         _ExtentX        =   6218
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.ComboBox cboLista 
         Height          =   315
         Left            =   -58300
         TabIndex        =   44
         Top             =   1590
         Visible         =   0   'False
         Width           =   705
         _Version        =   851968
         _ExtentX        =   1244
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton PBAccionesActualizar 
         Height          =   435
         Index           =   1
         Left            =   -61510
         TabIndex        =   56
         Top             =   2400
         Visible         =   0   'False
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Limpiar campos"
         Appearance      =   3
         Picture         =   "frmActualizacionPrecio.frx":059A
      End
      Begin XtremeSuiteControls.ProgressBar Barra 
         Height          =   285
         Left            =   120
         TabIndex        =   57
         Top             =   6180
         Width           =   12645
         _Version        =   851968
         _ExtentX        =   22304
         _ExtentY        =   503
         _StockProps     =   93
      End
      Begin XtremeSuiteControls.FlatEdit vSubRubroDescipcion 
         Height          =   315
         Left            =   -67090
         TabIndex        =   59
         Top             =   5670
         Visible         =   0   'False
         Width           =   2175
         _Version        =   851968
         _ExtentX        =   3836
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   6
         Left            =   -67570
         TabIndex        =   60
         Tag             =   "SubRubroD"
         Top             =   5670
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit vSubRubrocodigo 
         Height          =   315
         Left            =   -68410
         TabIndex        =   61
         Top             =   5670
         Visible         =   0   'False
         Width           =   735
         _Version        =   851968
         _ExtentX        =   1296
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit vRubroDescripcion 
         Height          =   315
         Left            =   -67060
         TabIndex        =   63
         Top             =   6060
         Visible         =   0   'False
         Width           =   2175
         _Version        =   851968
         _ExtentX        =   3836
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   7
         Left            =   -67570
         TabIndex        =   64
         Tag             =   "RubroD"
         Top             =   6060
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit vproveedorDescripcion 
         Height          =   315
         Left            =   -60250
         TabIndex        =   65
         Top             =   2640
         Visible         =   0   'False
         Width           =   3045
         _Version        =   851968
         _ExtentX        =   5371
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   8
         Left            =   -60640
         TabIndex        =   66
         Tag             =   "ProveedorD"
         Top             =   2640
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit vRubrocodigo 
         Height          =   315
         Left            =   -68410
         TabIndex        =   67
         Top             =   6060
         Visible         =   0   'False
         Width           =   735
         _Version        =   851968
         _ExtentX        =   1296
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit vproveedorCodigo 
         Height          =   315
         Left            =   -61420
         TabIndex        =   68
         Top             =   2640
         Visible         =   0   'False
         Width           =   735
         _Version        =   851968
         _ExtentX        =   1296
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit vivaDescripcion 
         Height          =   315
         Left            =   -60280
         TabIndex        =   73
         Top             =   3690
         Visible         =   0   'False
         Width           =   3015
         _Version        =   851968
         _ExtentX        =   5318
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vivaValor 
         Height          =   315
         Left            =   -61390
         TabIndex        =   74
         Top             =   3690
         Visible         =   0   'False
         Width           =   735
         _Version        =   851968
         _ExtentX        =   1296
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   9
         Left            =   -60610
         TabIndex        =   75
         Tag             =   "PorcentajeIva"
         Top             =   3690
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid g 
         Height          =   5145
         Left            =   90
         TabIndex        =   79
         Top             =   390
         Width           =   12675
         _ExtentX        =   22357
         _ExtentY        =   9075
         _Version        =   393216
         Rows            =   5
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin XtremeSuiteControls.PushButton PBAccionesActualizar 
         Height          =   375
         Index           =   0
         Left            =   -60190
         TabIndex        =   8
         Top             =   6060
         Visible         =   0   'False
         Width           =   2955
         _Version        =   851968
         _ExtentX        =   5212
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Ejecutar la actualización de precios"
         BackColor       =   -2147483644
         Appearance      =   3
         Picture         =   "frmActualizacionPrecio.frx":0B34
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   195
         Left            =   -69850
         TabIndex        =   82
         Top             =   5010
         Visible         =   0   'False
         Width           =   6795
         _Version        =   851968
         _ExtentX        =   11986
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "2.3. Cambiar categoría de artículos: "
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   4
         BorderStyle     =   1
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   195
         Left            =   -62410
         TabIndex        =   83
         Top             =   2040
         Visible         =   0   'False
         Width           =   5085
         _Version        =   851968
         _ExtentX        =   8969
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "2.4. Proveedor: "
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   4
         BorderStyle     =   1
      End
      Begin XtremeSuiteControls.GroupBox GroupBox4 
         Height          =   225
         Left            =   -62260
         TabIndex        =   84
         Top             =   3210
         Visible         =   0   'False
         Width           =   5055
         _Version        =   851968
         _ExtentX        =   8916
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "2.5. IVA:"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   4
         BorderStyle     =   1
      End
      Begin XtremeSuiteControls.GroupBox GroupBox5 
         Height          =   315
         Left            =   -69880
         TabIndex        =   85
         Top             =   330
         Visible         =   0   'False
         Width           =   6855
         _Version        =   851968
         _ExtentX        =   12091
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "2.1. Operación:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   4
         BorderStyle     =   1
      End
      Begin XtremeSuiteControls.CheckBox cbp 
         Height          =   375
         Index           =   2
         Left            =   -66670
         TabIndex        =   88
         Top             =   3240
         Visible         =   0   'False
         Width           =   2205
         _Version        =   851968
         _ExtentX        =   3889
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Ponerlo igual que el costo"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox cbp 
         Height          =   375
         Index           =   3
         Left            =   -66670
         TabIndex        =   89
         Top             =   3600
         Visible         =   0   'False
         Width           =   2205
         _Version        =   851968
         _ExtentX        =   3889
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Ponerlo igual que el costo"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox cbp 
         Height          =   375
         Index           =   4
         Left            =   -66670
         TabIndex        =   90
         Top             =   3990
         Visible         =   0   'False
         Width           =   2205
         _Version        =   851968
         _ExtentX        =   3889
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Ponerlo igual que el costo"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox cbp 
         Height          =   375
         Index           =   5
         Left            =   -66670
         TabIndex        =   91
         Top             =   4350
         Visible         =   0   'False
         Width           =   2205
         _Version        =   851968
         _ExtentX        =   3889
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Ponerlo igual que el costo"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkfijo 
         Height          =   330
         Left            =   -62800
         TabIndex        =   93
         Top             =   900
         Visible         =   0   'False
         Width           =   2175
         _Version        =   851968
         _ExtentX        =   3836
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Fijar Monto"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0000FF00&
         X1              =   7245
         X2              =   9045
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   9810
         X2              =   12735
         Y1              =   6030
         Y2              =   6030
      End
      Begin VB.Label Label7 
         Caption         =   $"frmActualizacionPrecio.frx":10CE
         ForeColor       =   &H000000FF&
         Height          =   2415
         Left            =   -69340
         TabIndex        =   92
         Top             =   3840
         Visible         =   0   'False
         Width           =   11025
      End
      Begin XtremeSuiteControls.Label Label9 
         Height          =   195
         Left            =   -68170
         TabIndex        =   78
         Top             =   2700
         Visible         =   0   'False
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "% Fijo s/ Costo"
      End
      Begin VB.Label Label10 
         Caption         =   "Los campos que quedan en cero o sin datos no se toman en cuenta para el cambio de precio. "
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   -69700
         TabIndex        =   77
         Top             =   930
         Visible         =   0   'False
         Width           =   6795
      End
      Begin XtremeSuiteControls.Label lblActualizar 
         Height          =   255
         Index           =   1
         Left            =   -61810
         TabIndex        =   76
         Top             =   3720
         Visible         =   0   'False
         Width           =   345
         _Version        =   851968
         _ExtentX        =   609
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "IVA:"
         Transparent     =   -1  'True
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000B&
         Caption         =   "Cambiarle el tipo de IVA:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -59350
         TabIndex        =   72
         Top             =   3420
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000B&
         Caption         =   "Cambiarle el proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -59380
         TabIndex        =   71
         Top             =   2220
         Visible         =   0   'False
         Width           =   2025
      End
      Begin XtremeSuiteControls.Label lblActualizar 
         Height          =   255
         Index           =   13
         Left            =   -69850
         TabIndex        =   70
         Top             =   6060
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Rubro:"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblActualizar 
         Height          =   255
         Index           =   12
         Left            =   -62260
         TabIndex        =   69
         Top             =   2670
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Proveedor:"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblActualizar 
         Height          =   255
         Index           =   11
         Left            =   -69880
         TabIndex        =   62
         Top             =   5670
         Visible         =   0   'False
         Width           =   945
         _Version        =   851968
         _ExtentX        =   1667
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "SubRubro:"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000B&
         Caption         =   "Cambios en el tipo de rubro y subro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -66160
         TabIndex        =   58
         Top             =   5190
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label Label4 
         Caption         =   "Lista de precio 5:"
         Height          =   195
         Left            =   -69640
         TabIndex        =   54
         Top             =   4440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Lista de precio 4:"
         Height          =   195
         Left            =   -69640
         TabIndex        =   53
         Top             =   4110
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Lista de precio 3:"
         Height          =   195
         Left            =   -69640
         TabIndex        =   52
         Top             =   3720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Lista de precio 2:"
         Height          =   195
         Left            =   -69640
         TabIndex        =   51
         Top             =   3390
         Visible         =   0   'False
         Width           =   1215
      End
      Begin XtremeSuiteControls.Label lblActualizar 
         Height          =   255
         Index           =   2
         Left            =   -59710
         TabIndex        =   49
         Top             =   1590
         Visible         =   0   'False
         Width           =   1245
         _Version        =   851968
         _ExtentX        =   2196
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Lista de Precio :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblActualizar 
         Height          =   255
         Index           =   0
         Left            =   -69910
         TabIndex        =   48
         Top             =   630
         Visible         =   0   'False
         Width           =   2955
         _Version        =   851968
         _ExtentX        =   5212
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Seleccionar el tipo de operación :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin VB.Label lblIngLos 
         BackColor       =   &H8000000B&
         Caption         =   "Ingrese  los  porcentajes para cada una de las listas de precio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -68470
         TabIndex        =   47
         Top             =   2160
         Visible         =   0   'False
         Width           =   5475
      End
      Begin VB.Label lblListaDe 
         Caption         =   "Lista de precio 1:"
         Height          =   195
         Left            =   -69610
         TabIndex        =   46
         Top             =   3000
         Visible         =   0   'False
         Width           =   1215
      End
      Begin XtremeSuiteControls.Label lblActualizar 
         Height          =   255
         Index           =   10
         Left            =   -69610
         TabIndex        =   45
         Top             =   1290
         Visible         =   0   'False
         Width           =   3885
         _Version        =   851968
         _ExtentX        =   6853
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Porcentaje que quiere aplicarle al precios de Costo: "
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblActualizar 
         Height          =   255
         Index           =   6
         Left            =   -69820
         TabIndex        =   40
         Top             =   1620
         Visible         =   0   'False
         Width           =   1125
         _Version        =   851968
         _ExtentX        =   1984
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Rubro Desde:"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblActualizar 
         Height          =   255
         Index           =   8
         Left            =   -69820
         TabIndex        =   39
         Top             =   1980
         Visible         =   0   'False
         Width           =   1935
         _Version        =   851968
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Proveedor Desde:"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblActualizar 
         Height          =   255
         Index           =   7
         Left            =   -63970
         TabIndex        =   38
         Top             =   1620
         Visible         =   0   'False
         Width           =   1935
         _Version        =   851968
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Hasta:"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblActualizar 
         Height          =   255
         Index           =   9
         Left            =   -63970
         TabIndex        =   37
         Top             =   1980
         Visible         =   0   'False
         Width           =   1935
         _Version        =   851968
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Hasta:"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblActualizar 
         Height          =   255
         Index           =   3
         Left            =   -69820
         TabIndex        =   23
         Top             =   870
         Visible         =   0   'False
         Width           =   2355
         _Version        =   851968
         _ExtentX        =   4154
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Parte de descripción / código :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblActualizar 
         Height          =   255
         Index           =   4
         Left            =   -69820
         TabIndex        =   22
         Top             =   1230
         Visible         =   0   'False
         Width           =   1935
         _Version        =   851968
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "SubRubro Desde:"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblActualizar 
         Height          =   255
         Index           =   5
         Left            =   -63970
         TabIndex        =   21
         Top             =   1230
         Visible         =   0   'False
         Width           =   1935
         _Version        =   851968
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Hasta:"
         Transparent     =   -1  'True
      End
   End
   Begin VB.PictureBox PicInferior 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   1170
      Picture         =   "frmActualizacionPrecio.frx":1286
      ScaleHeight     =   555
      ScaleWidth      =   9735
      TabIndex        =   9
      Top             =   2070
      Width           =   9735
      Begin XtremeSuiteControls.PushButton PBAccionesActualizar 
         Height          =   375
         Index           =   2
         Left            =   8520
         TabIndex        =   7
         Top             =   90
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cerrar"
         Appearance      =   6
      End
      Begin VB.Label lblWGestion 
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
         TabIndex        =   10
         Top             =   150
         Width           =   1770
      End
      Begin VB.Label lblWGestion 
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
         TabIndex        =   11
         Top             =   170
         Width           =   1770
      End
   End
   Begin XtremeSuiteControls.Label lblDisplay 
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   4040
      Width           =   9735
      _Version        =   851968
      _ExtentX        =   17171
      _ExtentY        =   450
      _StockProps     =   79
      BackColor       =   14737632
      Transparent     =   -1  'True
   End
End
Attribute VB_Name = "frmActualizacionPrecio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sqlActualizacion As String
Dim sqlActualizacion2 As String

Dim vguardar As Integer
Dim rsActualizacion As ADODB.Recordset
Private Sub chkAbroArticulo_Click()
On Error Resume Next

    If chkAbroArticulo.Value = xtpChecked Then
        With frmArticulos
            .Show
            .Buscar (sqlActualizacion)
            .ZOrder (1)
        End With
    Else
        Unload frmArticulos
    End If
    
If Err Then GrabarLog "chkAbroArticulo_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If

If Err Then GrabarLog "Form_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Load()
On Error Resume Next
vguardar = 0
Me.tab.Selected = 0
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub pbCarga_Click(Index As Integer)
On Error Resume Next

    vVuelveBusqueda = Me.Name
    vVieneBusqueda = pbCarga(Index).Tag
    
    Select Case Index
    
        Case 0 To 5
            frmBusqueda.Show
        
        Case 6
            Call fbuscarGrilla("SubRubros", "SubRubro", "idSubRubros", Me.vSubRubroDescipcion.Name, Me)
        Case 7
            Call fbuscarGrilla("Rubros", "Rubro", "idRubros", Me.vRubroDescripcion.Name, Me)
        Case 8
            Call fbuscarGrilla("proveedores", "Nombre", "Codigo", Me.vproveedorDescripcion.Name, Me)
        Case 9
            Call fbuscarGrilla("PorcentajeIva", "Descripcion", "idPorcentajeIva", Me.vivaDescripcion.Name, Me)
              
    
    End Select
        
If Err Then GrabarLog "pbCarga_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub PBAccionesActualizar_Click(Index As Integer)
On Error Resume Next
    
    Select Case Index
    
        Case 0
            If ValidarCampos() = True Then
                g.Clear
                lbldisplay.Visible = False
                barra.Visible = True
                GenerarFiltro
                ActualizarPrecios
                Me.tab.SelectedItem = 2
                
            End If
        Case 1
            Limpiar
            
        Case 2
            Unload Me
            
    End Select

If Err Then GrabarLog "PBAccionesActualizar_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Function ValidarCampos() As Boolean
On Error Resume Next

    Dim i As Integer
    
    ValidarCampos = True
    
    If Trim(cboOperacion.Text) = "" Then
        MsgBox "Debe seleccionar una Operacion", vbExclamation, "Mensaje ..."
        ValidarCampos = Not True
        Exit Function
    End If

   ' If Val(txtPorcentaje.Text) = 0 Then
   '     MsgBox "El campo Porcentaje es ingreso obligatorio.", vbExclamation, "Mensaje ..."
   '     ValidarCampos = Not True
   '     Exit Function
   ' End If
    
  '  If Val(cboLista.Text) = 0 Then
  '      MsgBox "Debe seleccionar una Lista de Precios para Poder Actualizar", vbExclamation, "Mensaje ..."
  '      ValidarCampos = Not True
  '      Exit Function
  '  End If

If Err Then GrabarLog "ValidarCampos", Err.Number & " " & Err.Description, Me.Caption
End Function
Private Sub Limpiar()
On Error Resume Next

    Dim i As Integer
    
    chkAbroArticulo.Value = xtpUnchecked
    cboOperacion.Clear
    cboOperacion.Text = ""
    txtPorcentaje.Text = ""
    cbolista.Clear
    cbolista.Text = ""
    
    For i = 0 To Me.txtFiltro.Count - 1
        txtFiltro(i).Text = ""
    Next
            
    lbldisplay.Caption = ""
    barra.Value = 0
    barra.Visible = False
    
If Err Then GrabarLog "Limpiar", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub GenerarFiltro()
On Error Resume Next
    
    sqlActualizacion = ""
    lbldisplay.Caption = ""
    
    If Not Trim(txtFiltro(0).Text) = "" Then
        sqlActualizacion = sqlActualizacion & " AND (Descrip LIKE '%" & Trim(txtFiltro(0).Text) & "%' or codigo like '%" & Trim(txtFiltro(0).Text) & "%')"
        lbldisplay.Caption = lbldisplay.Caption & " Descripcion : " & Trim(txtFiltro(0).Text) & " /"
    End If

    If Not Trim(txtFiltro(1).Text) = "" Then
        sqlActualizacion = sqlActualizacion & " AND (idSubRubros >= '" & Trim(txtFiltro(1).Text) & "')"
        lbldisplay.Caption = lbldisplay.Caption & " SubRubro Desde : " & Trim(txtFiltro(2).Text) & " /"

        If Not Trim(txtFiltro(3).Text) = "" Then
            sqlActualizacion = sqlActualizacion & " AND (idSubRubros <= '" & Trim(txtFiltro(3).Text) & "')"
            lbldisplay.Caption = lbldisplay.Caption & " SubRubro Hasta : " & Trim(txtFiltro(4).Text) & " /"
        End If
    End If
    
    
    If Not Trim(txtFiltro(5).Text) = "" Then
        sqlActualizacion = sqlActualizacion & " AND (idRubros >= '" & Trim(txtFiltro(5).Text) & "')"
        lbldisplay.Caption = lbldisplay.Caption & " Rubro Desde : " & Trim(txtFiltro(6).Text) & " /"
    
        If Not Trim(txtFiltro(7).Text) = "" Then
            sqlActualizacion = sqlActualizacion & " AND (idRubros <= '" & Trim(txtFiltro(7).Text) & "')"
            lbldisplay.Caption = lbldisplay.Caption & " Rubro Desde : " & Trim(txtFiltro(8).Text) & " /"
        End If
    
    End If
    
    If Not Trim(txtFiltro(9).Text) = "" Then
        sqlActualizacion = sqlActualizacion & " AND (idProveedor >= '" & Trim(txtFiltro(9).Text) & "')"
        lbldisplay.Caption = lbldisplay.Caption & " Proveedor Desde : " & Trim(txtFiltro(10).Text)

        If Not Trim(txtFiltro(11).Text) = "" Then
            sqlActualizacion = sqlActualizacion & " AND (idProveedor <= '" & Trim(txtFiltro(11).Text) & "')"
            lbldisplay.Caption = lbldisplay.Caption & " Proveedor Desde : " & Trim(txtFiltro(12).Text)
        End If
    
    End If

    sqlActualizacion2 = "SELECT * FROM Articulos WHERE 1=1" & sqlActualizacion & " ORDER BY 1"
    
If Err Then GrabarLog "GenerarFiltro", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub ActualizarPrecios2()
Dim vauxi, vovalores, vpvalores As String
Dim i As Integer
Dim vvPrecioCosto As Double
Dim signo As Integer

signo = 1
If Not Trim(cboOperacion.Text) = "Aumentar Precio" Then signo = -1

vauxi = ""
    On Error Resume Next
    
    Dim vPrecioArticulo As Double, vPrecio2 As Double, vPrecioCosto As Double, vPrecioCosto2 As Double
    
    If MsgBox("¿Esta seguro que desea realizar esta accion ? ", vbInformation + vbYesNo) = vbYes Then
    
        Set rsActualizacion = New ADODB.Recordset
        
        With rsActualizacion
            .CursorLocation = adUseClient
        
            Call .Open("SELECT * FROM Articulos WHERE 1=1" & sqlActualizacion, ConnDDBB, adOpenStatic, adLockPessimistic)
            
            .Fields.Refresh
            
            If Not .EOF = True And Not .BOF = True Then
                barra.Value = 0
                barra.Max = .RecordCount
                
                .MoveFirst
    
           
    
    For i = 1 To 5
                
                cbolista.Text = i
                .MoveFirst
                
                barra.Value = 0
                                
                Do Until .EOF = True
                    
                    vPrecioArticulo = 0
                    vPrecioCosto = .Fields("pcosto")
                    
                    If Not IsNull(.Fields("PVenta" & Trim(cbolista.Text)).Value) = True Then vPrecioArticulo = .Fields("PVenta" & Trim(cbolista.Text)).Value
                    
                    'vPrecio2 = (vPrecioArticulo * Val(txtPorcentaje.Text) / 100)
                    
                    'vvPrecioCosto = Fields("PCosto").Value * (1 + Me.vPorcentajeCosto / 100)
                    If Val(vPorcentajeCosto.Text) > 0 Then vPrecioCosto = Val(EsNulo(.Fields("pcosto"))) * (1 + signo * Val(vPorcentajeCosto.Text) / 100)
                    
    
                    If Val(p(i)) > 0 Then
                        ' acá tengo que agregarle al costo el aumento
                       ' vPrecio2 = .Fields("PCosto").Value * p(i) / 100
                        vPrecio2 = vPrecioCosto * p(i) / 100
                    Else
                         If p(1).Text = "" And p(2).Text = "" And p(3).Text = "" And p(4).Text = "" Then
                                Call actualizarPreciosArticulo(.Fields("idarticulos"), vPrecioCosto)
                         End If
                    End If
    
    
                   ' If Not Format(.Fields("PCosto").Value, "###########0.00") = "" Then vPrecioCosto = vvPrecioCosto
                   ' vPrecioCosto2 = (vPrecioCosto * Val(txtPorcentaje.Text) / 100)
    
                    vPrecio2 = vPrecio2 * signo
                    'If Not Trim(cboOperacion.Text) = "Aumentar Precio" Then vPrecioCosto2 = vPrecioCosto2 * -1
                    ' doing
                    
                    
                    
                   ' If Val(vPorcentajeCosto.Text) > 0 Then vPrecioCosto = .Fields("pcosto") * (1 + Val(vPorcentajeCosto.Text) / 100)
                    
                    
                    log.AddItem ("> " + .Fields("codigo") + " - P.Anterior: " + Str(.Fields("Pventa" & Trim(cbolista.Text)).Value) + "  - P.Actual: ")
                    
                    
                    '----- actualizaciones de campos diferentes a los precios -------
                    vovalores = ""
                    If Not Me.vSubRubroDescipcion.Tag = "" Then vovalores = vovalores + ",idSubRubros='" + Me.vSubRubroDescipcion.Tag + "'"
                    If Not Me.vRubroDescripcion.Tag = "" Then vovalores = vovalores + ",idRubros='" + Me.vRubroDescripcion.Tag + "'"
                    If Not Me.vproveedorDescripcion.Tag = "" Then vovalores = vovalores + ",idProveedor='" + Me.vproveedorDescripcion.Tag + "'"
                    If Not Me.vivaDescripcion.Tag = "" Then vovalores = vovalores + ",idPorcentajeIva='" + Me.vivaDescripcion.Tag + "'"
                    '-----------------------------------------------------------------
                    
                    
                    
                    '--------------- actualizaciones de precios de costos y las otras listas ---------------
                    vpvalores = ""
                    If Val(vPrecioCosto) > 0 Then vpvalores = vpvalores + ",pcosto=" + Str(vPrecioCosto)
                    If Val(vPrecio2) > 0 Then vpvalores = vpvalores + ", pventa" + Trim(cbolista.Text) + " = " + Str(vPrecioCosto + vPrecio2)
                    '-----------------------------------------------------------------
                    
                    
                    ' acá tengo que cambiar el precio de costo tambien
                    'vauxi = "update articulos set pcosto=" + Str(vPrecioCosto) + ", pventa" + Trim(cboLista.Text) + " = " + Str(vPrecioCosto + vPrecio2) + vovalores + " where idArticulos=" + Str(.Fields("idArticulos"))
                    
                    
                    vauxi = "update articulos set codigo=codigo" + vpvalores + vovalores + " where idArticulos=" + Str(.Fields("idArticulos"))
                   
                    Call EjecutarScript(vauxi, pathDBMySQL)
    
    
                    
                    'vauxi = "update articulos set pventa" + Trim(cboLista.Text) + " = " + Str(vPrecioArticulo + vPrecio2) + " where idArticulos=" + Str(.Fields("idArticulos"))
                    
                    'Call EjecutarScript(vauxi, pathDBMySQL)
    
    
    
    
    
                    '.Fields("Pventa" & Trim(cboLista.Text)).Value = vPrecioArticulo + vPrecio2
                    '.Fields("PCosto").Value = vPrecioCosto + vPrecioCosto2
                    '.Update
                    .MoveNext
                                    
                    barra.Value = barra.Value + 1
                Loop
    Next
            Else
                MsgBox "No hay Articulos encontrados para aplicar los cambios que desea", vbInformation, "Mensaje ..."
            End If
        
        End With
        
        MsgBox "Los precios fueron actualizados", vbInformation, "Actualización de precios"
    
    End If

    If Err Then GrabarLog "ActualizarPrecios", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cboLista_GotFocus()
On Error Resume Next

    Call CargarComboNew("Listas", "Lista", cbolista, True)
       
If Err Then GrabarLog "cboLista_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cboOperacion_GotFocus()
On Error Resume Next
    
    Call cboOperacion.Clear
    Call cboOperacion.AddItem("Aumentar Precio", 0)
    Call cboOperacion.AddItem("Disminuir Precio", 1)

If Err Then GrabarLog "cboOperacion_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub PushButton1_Click()
vguardar = 1
Call PBAccionesActualizar_Click(0)
Call PBAccionesActualizar_Click(1)
End Sub

Private Sub txtFiltro_Change(Index As Integer)
On Error Resume Next

    Call GenerarFiltro
    
    If chkAbroArticulo.Value = xtpChecked Then
        
        With frmArticulos
            .Show
            .Buscar (sqlActualizacion)
            .ZOrder (1)
        End With
    
    Else
        
        Unload frmArticulos
        
    End If
    
If Err Then GrabarLog "txtFiltro_Change", Err.Number & " " & Err.Description, Me.Caption
End Sub


Private Sub ActualizarPrecios()
Dim vauxi, vlog As String
Dim i, signo  As Integer
Dim vpcosto As Double


g.Cols = 7
g.Rows = 1
g.ColWidth(1) = 3000
g.Clear
g.AddItem ("Código" + vbTab + "Descrip" + vbTab + "CostoViejo" + vbTab + "CostoNuevo" + vbTab + "Lista" + vbTab + "PrecioViejo" + vbTab + "PrecioNuevo")

vauxi = ""
    On Error Resume Next
    
    Dim vPrecioArticulo As Double, vPrecio2 As Double, vPrecioCosto As Double, vPrecioCosto2 As Double
    
    If MsgBox("¿Esta seguro que desea realizar esta acción ? ", vbInformation + vbYesNo) = vbYes Then
    
        Set rsActualizacion = New ADODB.Recordset
        
        With rsActualizacion
            .CursorLocation = adUseClient
        
            
            If Me.chkNot = xtpChecked And Not sqlActualizacion = "" Then
            
                Call .Open("SELECT * FROM Articulos WHERE not (1=1" & sqlActualizacion + ")", ConnDDBB, adOpenStatic, adLockPessimistic)
            
            Else
            
                Call .Open("SELECT * FROM Articulos WHERE 1=1" & sqlActualizacion, ConnDDBB, adOpenStatic, adLockPessimistic)
            
            End If
            
            .Fields.Refresh
            
            If Not .EOF = True And Not .BOF = True Then
                barra.Value = 0
                barra.Max = .RecordCount
                
                .MoveFirst
                
                
                
  Do Until .EOF = True
  
                    signo = 1
                    If Not Trim(cboOperacion.Text) = "Aumentar Precio" Then signo = -1
  
                    If Val(Me.vppcosto) >= 0 Then
                    
                        
                       If chkfijo.Value Then
                            vpcosto = Val(vppcosto.Text)
                       Else
                            vpcosto = (signo * (.Fields("pcosto").Value * Val(vppcosto.Text) / 100)) + .Fields("pcosto").Value
                       End If
 
                        
                    
                    
                    
                    End If

 
    For i = 1 To 2
                
                cbolista.Text = i
                '.MoveFirst
                                  
                    vPrecioArticulo = 0
                    
                    If Not IsNull(.Fields("PVenta" & Trim(cbolista.Text)).Value) = True Then
                        vPrecioArticulo = .Fields("PVenta" & Trim(cbolista.Text)).Value
                    Else
                        vPrecioArticulo = .Fields("PCosto").Value
                    End If
                    'vPrecio2 = (vPrecioArticulo * Val(txtPorcentaje.Text) / 100)
    
    
                    If Val(p(i)) > 0 And cbp(i).Value = xtpUnchecked Then
                    
                        If .Fields("PVenta" + Trim(Str(i))).Value = 0 Then
                        
                            If chkfijo.Value Then
                                vPrecio2 = p(i)
                            Else
                                vPrecio2 = .Fields("PCosto").Value * p(i) / 100
                            End If
                        Else
                            If Not chkfijo.Value = 1 Then
                                vPrecio2 = .Fields("PVenta" + Trim(Str(i))).Value * p(i) / 100
                            
                            Else
                                vPrecio2 = p(i)
                            
                            End If

                            Debug.Print Str(.Fields("PVenta" + Trim(Str(i))).Value)
                        End If
                    
                    
                    If chkfijo.Value Then
                        If vPrecio2 = 0 Then vPrecio2 = p(i)
                    Else
                        If vPrecio2 = 0 Then vPrecio2 = .Fields("pcosto").Value * p(i) / 100
                    End If
                                              
                        
                    Else
                        
                        vPrecio2 = 0
                    End If
                    
                    
                    
                  
                    
    
    
                    If Not Format(.Fields("PCosto").Value, "###########0.00") = "" Then vPrecioCosto = .Fields("Pcosto").Value ' este código está muerto
                   ' vPrecioCosto2 = (vPrecioCosto * Val(txtPorcentaje.Text) / 100)
    
                    If Not Trim(cboOperacion.Text) = "Aumentar Precio" Then vPrecio2 = vPrecio2 * -1
                    'If Not Trim(cboOperacion.Text) = "Aumentar Precio" Then vPrecioCosto2 = vPrecioCosto2 * -1
                    ' doing
                    
                    
                    If Not chkfijo.Value = xtpChecked Then
                        vlog = Trim((.Fields("codigo"))) + vbTab + Trim((.Fields("descrip"))) + vbTab + Trim(Str(.Fields("Pcosto"))) + vbTab + Trim(Str(vpcosto)) + vbTab + "Lista" + Trim(i) + vbTab + Trim(Str(.Fields("Pventa" + Trim(i)))) + vbTab + Trim(Str(.Fields("Pventa" + Trim(i)) + vPrecio2))
                    Else
                        vlog = Trim((.Fields("codigo"))) + vbTab + Trim((.Fields("descrip"))) + vbTab + Trim(Str(.Fields("Pcosto"))) + vbTab + Trim(Str(vpcosto)) + vbTab + "Lista" + Trim(i) + vbTab + Trim(Str(.Fields("Pventa" + Trim(i)))) + vbTab + (Str(vPrecio2))
                    End If
                   
                    
                    
                   ' If Not log = "" Then g.AddItem (vlog)
                    'log.AddItem ("> " + .Fields("codigo") + " - P.Anterior: " + Str(.Fields("Pventa" & Trim(cboLista.Text)).Value) + "  - P.Actual: ")
                    
                    If Not chkfijo.Value = 1 Then

                        vauxi = "update articulos set pventa" + Trim(i) + " = " + Str(.Fields("Pventa" + Trim(i)) + vPrecio2) + ", pcosto=" + Str(vpcosto) + " where idArticulos=" + Str(.Fields("idArticulos"))
                    Else
                    
                        vauxi = "update articulos set pventa" + Trim(i) + " = " + Str(vPrecio2) + ", pcosto=" + Str(vpcosto) + " where idArticulos=" + Str(.Fields("idArticulos"))
                    End If
                    
                    If Me.cbp(i).Value = 1 Then
                        
                        vauxi = "update articulos set pventa" + Trim(i) + " = " + Str(vpcosto) + ", pcosto=" + Str(vpcosto) + " where idArticulos=" + Str(.Fields("idArticulos"))
                        
                        g.AddItem Trim((.Fields("codigo"))) + vbTab + Trim((.Fields("descrip"))) + vbTab + Trim(Str(.Fields("Pcosto"))) + vbTab + Trim(Str(vpcosto)) + vbTab + "Lista" + Trim(i) + vbTab + Trim(Str(.Fields("Pventa" + Trim(i)))) + vbTab + Trim(Str(vpcosto))

                    Else
                                        
                    If Not chkfijo.Value = xtpChecked Then

                    
                        g.AddItem Trim((.Fields("codigo"))) + vbTab + Trim((.Fields("descrip"))) + vbTab + Trim(Str(.Fields("Pcosto"))) + vbTab + Trim(Str(vpcosto)) + vbTab + "Lista" + Trim(i) + vbTab + Trim(Str(.Fields("Pventa" + Trim(i)))) + vbTab + Trim(Str(.Fields("Pventa" + Trim(i)) + vPrecio2))

                    Else
                        g.AddItem Trim((.Fields("codigo"))) + vbTab + Trim((.Fields("descrip"))) + vbTab + Trim(Str(.Fields("Pcosto"))) + vbTab + Trim(Str(vpcosto)) + vbTab + "Lista" + Trim(i) + vbTab + Trim(Str(.Fields("Pventa" + Trim(i)))) + vbTab + Str(vPrecio2)
                    
                    End If
                    
                        Debug.Print Trim((.Fields("codigo"))) + vbTab + Trim((.Fields("descrip"))) + vbTab + Trim(Str(.Fields("Pcosto"))) + vbTab + Trim(Str(vpcosto)) + vbTab + "Lista" + Trim(i) + vbTab + Trim(Str(.Fields("Pventa" + Trim(i)))) + vbTab + Str(vPrecio2)
                
                        
                    End If
                    
                    
                    
                    If vguardar = 1 Then Call EjecutarScript(Trim(vauxi), pathDBMySQL)
    
    
                    
                    'vauxi = "update articulos set pventa" + Trim(cboLista.Text) + " = " + Str(vPrecioArticulo + vPrecio2) + " where idArticulos=" + Str(.Fields("idArticulos"))
                    
                    'Call EjecutarScript(vauxi, pathDBMySQL)
    
    
    
    
    
                    '.Fields("Pventa" & Trim(cboLista.Text)).Value = vPrecioArticulo + vPrecio2
                    '.Fields("PCosto").Value = vPrecioCosto + vPrecioCosto2
                    '.Update
               '     .MoveNext
    Next i
                    barra.Value = barra.Value + 1
                    .MoveNext
                Loop
    
            Else
                MsgBox "No hay Articulos encontrados para aplicar los cambios que desea", vbInformation, "Mensaje ..."
            End If
        
        End With
        
        MsgBox "Los precios fueron actualizados", vbInformation, "Actualización de precios"
    
    End If

    vguardar = 0
    
    If Err Then GrabarLog "ActualizarPrecios", Err.Number & " " & Err.Description, Me.Name
End Sub
