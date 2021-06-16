VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#13.0#0"; "Codejock.ReportControl.v13.0.0.Demo.ocx"
Begin VB.Form frmTablero 
   Caption         =   "Tablero de Control"
   ClientHeight    =   7560
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12675
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7560
   ScaleWidth      =   12675
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.GroupBox cartel 
      Height          =   1815
      Left            =   1590
      TabIndex        =   29
      Top             =   2520
      Width           =   17265
      _Version        =   851968
      _ExtentX        =   30454
      _ExtentY        =   3201
      _StockProps     =   79
      BackColor       =   4210752
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.ProgressBar barra 
         Height          =   195
         Left            =   150
         TabIndex        =   32
         Top             =   1290
         Width           =   16845
         _Version        =   851968
         _ExtentX        =   29713
         _ExtentY        =   344
         _StockProps     =   93
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label lblProcesandoLos 
         Height          =   1095
         Left            =   150
         TabIndex        =   37
         Top             =   180
         Width           =   16575
         _Version        =   851968
         _ExtentX        =   29236
         _ExtentY        =   1931
         _StockProps     =   79
         Caption         =   "Procesando los datos del tablero. Espere ..."
         ForeColor       =   65280
         BackColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Mincho"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin XtremeSuiteControls.TabControl tab1 
      Height          =   5805
      Left            =   150
      TabIndex        =   5
      Top             =   90
      Width           =   19845
      _Version        =   851968
      _ExtentX        =   35004
      _ExtentY        =   10239
      _StockProps     =   68
      ItemCount       =   5
      Item(0).Caption =   "Generales"
      Item(0).ControlCount=   10
      Item(0).Control(0)=   "PusAgregarConcepto"
      Item(0).Control(1)=   "GroDeDonde"
      Item(0).Control(2)=   "PusActualizar"
      Item(0).Control(3)=   "vfecha"
      Item(0).Control(4)=   "g1"
      Item(0).Control(5)=   "GroupBox1"
      Item(0).Control(6)=   "PusBorrarConcepto"
      Item(0).Control(7)=   "lblDatosPara"
      Item(0).Control(8)=   "vconcepto"
      Item(0).Control(9)=   "GroOptimismo"
      Item(1).Caption =   "Configuración"
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "vpromedios"
      Item(1).Control(1)=   "lblCantidadDe"
      Item(1).Control(2)=   "Label1"
      Item(1).Control(3)=   "vpromediosFactura"
      Item(2).Caption =   "Alarmas"
      Item(2).ControlCount=   4
      Item(2).Control(0)=   "grillaAlarma"
      Item(2).Control(1)=   "PushButton1"
      Item(2).Control(2)=   "Gastos"
      Item(2).Control(3)=   "personas"
      Item(3).Caption =   "Dimensionales"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "wndReportControl"
      Item(4).Caption =   "Estado de las personas"
      Item(4).ControlCount=   5
      Item(4).Control(0)=   "gpersona"
      Item(4).Control(1)=   "PusPersonas"
      Item(4).Control(2)=   "vcliprovee"
      Item(4).Control(3)=   "Picture3"
      Item(4).Control(4)=   "lblAsientos(11)"
      Begin XtremeReportControl.ReportControl wndReportControl 
         Height          =   4905
         Left            =   -69760
         TabIndex        =   38
         Top             =   690
         Visible         =   0   'False
         Width           =   19575
         _Version        =   851968
         _ExtentX        =   34528
         _ExtentY        =   8652
         _StockProps     =   64
         ShowGroupBox    =   -1  'True
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -68770
         Picture         =   "frmTablero.frx":0000
         ScaleHeight     =   315
         ScaleWidth      =   495
         TabIndex        =   1
         Top             =   660
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox vcliprovee 
         Height          =   315
         Left            =   -68260
         TabIndex        =   2
         Top             =   660
         Visible         =   0   'False
         Width           =   5925
      End
      Begin XtremeSuiteControls.GroupBox GroOptimismo 
         Height          =   675
         Left            =   3300
         TabIndex        =   33
         Top             =   570
         Width           =   3405
         _Version        =   851968
         _ExtentX        =   6006
         _ExtentY        =   1191
         _StockProps     =   79
         Caption         =   "Optimismo:"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton RadPesimista 
            Height          =   255
            Left            =   150
            TabIndex        =   34
            Top             =   300
            Width           =   1005
            _Version        =   851968
            _ExtentX        =   1773
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Pesimista"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RadModerado 
            Height          =   255
            Left            =   1230
            TabIndex        =   35
            Top             =   300
            Width           =   1005
            _Version        =   851968
            _ExtentX        =   1773
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Moderado"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RadOptimista 
            Height          =   255
            Left            =   2340
            TabIndex        =   36
            Top             =   300
            Width           =   1005
            _Version        =   851968
            _ExtentX        =   1773
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Optimista"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.FlatEdit vconcepto 
         Height          =   255
         Left            =   2250
         TabIndex        =   28
         Top             =   5430
         Width           =   4125
         _Version        =   851968
         _ExtentX        =   7276
         _ExtentY        =   450
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grillaAlarma 
         Height          =   4845
         Left            =   -69910
         TabIndex        =   21
         Top             =   810
         Visible         =   0   'False
         Width           =   19545
         _ExtentX        =   34475
         _ExtentY        =   8546
         _Version        =   393216
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin XtremeSuiteControls.FlatEdit vpromedios 
         Height          =   345
         Left            =   -65800
         TabIndex        =   17
         Top             =   810
         Visible         =   0   'False
         Width           =   1635
         _Version        =   851968
         _ExtentX        =   2884
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "2"
      End
      Begin XtremeSuiteControls.PushButton PusAgregarConcepto 
         Height          =   255
         Left            =   150
         TabIndex        =   6
         Top             =   5460
         Width           =   1425
         _Version        =   851968
         _ExtentX        =   2514
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Agregar concepto "
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroDeDonde 
         Height          =   4755
         Left            =   8340
         TabIndex        =   7
         Top             =   960
         Width           =   7635
         _Version        =   851968
         _ExtentX        =   13467
         _ExtentY        =   8387
         _StockProps     =   79
         Caption         =   "De donde salen los datos del resultado:"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.PushButton Imprimir 
            Height          =   285
            Left            =   180
            TabIndex        =   8
            Top             =   5460
            Width           =   1785
            _Version        =   851968
            _ExtentX        =   3149
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Imprimir"
            UseVisualStyle  =   -1  'True
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid g4 
            Height          =   3975
            Left            =   120
            TabIndex        =   9
            Top             =   660
            Width           =   7425
            _ExtentX        =   13097
            _ExtentY        =   7011
            _Version        =   393216
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin XtremeSuiteControls.FlatEdit vbuscar 
            Height          =   375
            Left            =   3000
            TabIndex        =   41
            Top             =   210
            Width           =   4515
            _Version        =   851968
            _ExtentX        =   7964
            _ExtentY        =   661
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.Label lblIngreseParte 
            Height          =   315
            Left            =   120
            TabIndex        =   42
            Top             =   270
            Width           =   2895
            _Version        =   851968
            _ExtentX        =   5106
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Ingrese parte del nombre de la persona:"
         End
      End
      Begin XtremeSuiteControls.PushButton PusActualizar 
         Height          =   345
         Left            =   6960
         TabIndex        =   10
         Top             =   810
         Width           =   1275
         _Version        =   851968
         _ExtentX        =   2249
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Actualizar"
         UseVisualStyle  =   -1  'True
      End
      Begin MSComCtl2.DTPicker vfecha 
         Height          =   285
         Left            =   1560
         TabIndex        =   11
         Top             =   780
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   503
         _Version        =   393216
         Format          =   59375617
         CurrentDate     =   41967
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid g1 
         Height          =   4005
         Left            =   150
         TabIndex        =   12
         Top             =   1350
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   7064
         _Version        =   393216
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   4755
         Left            =   15990
         TabIndex        =   13
         Top             =   960
         Width           =   3645
         _Version        =   851968
         _ExtentX        =   6429
         _ExtentY        =   8387
         _StockProps     =   79
         Caption         =   "Clasificador:"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   4365
            Left            =   90
            TabIndex        =   14
            Top             =   270
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   7699
            _Version        =   393216
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin XtremeSuiteControls.PushButton PusBorrarConcepto 
         Height          =   285
         Left            =   6960
         TabIndex        =   15
         Top             =   5430
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Borrar Concepto"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit vpromediosFactura 
         Height          =   345
         Left            =   -65800
         TabIndex        =   19
         Top             =   1290
         Visible         =   0   'False
         Width           =   1635
         _Version        =   851968
         _ExtentX        =   2884
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "2"
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   315
         Left            =   -69880
         TabIndex        =   30
         Top             =   420
         Visible         =   0   'False
         Width           =   2565
         _Version        =   851968
         _ExtentX        =   4524
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Mal balanceado las compras"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton personas 
         Height          =   315
         Left            =   -67240
         TabIndex        =   31
         Top             =   420
         Visible         =   0   'False
         Width           =   1365
         _Version        =   851968
         _ExtentX        =   2408
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Personas"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton Gastos 
         Height          =   315
         Left            =   -65830
         TabIndex        =   39
         Top             =   420
         Visible         =   0   'False
         Width           =   1365
         _Version        =   851968
         _ExtentX        =   2408
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Gastos"
         UseVisualStyle  =   -1  'True
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gpersona 
         Height          =   4485
         Left            =   -69910
         TabIndex        =   3
         Top             =   1170
         Visible         =   0   'False
         Width           =   19695
         _ExtentX        =   34740
         _ExtentY        =   7911
         _Version        =   393216
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin XtremeSuiteControls.PushButton PusPersonas 
         Height          =   345
         Left            =   -62170
         TabIndex        =   4
         Top             =   660
         Visible         =   0   'False
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Personas"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label lblAsientos 
         Alignment       =   1  'Right Justify
         Caption         =   "Personas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   11
         Left            =   -69910
         TabIndex        =   25
         Top             =   720
         Visible         =   0   'False
         Width           =   1065
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   315
         Left            =   -69760
         TabIndex        =   20
         Top             =   1290
         Visible         =   0   'False
         Width           =   3825
         _Version        =   851968
         _ExtentX        =   6747
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "> Cantidad de meses para vencimientos de facturas:"
      End
      Begin XtremeSuiteControls.Label lblCantidadDe 
         Height          =   315
         Left            =   -69730
         TabIndex        =   18
         Top             =   810
         Visible         =   0   'False
         Width           =   3705
         _Version        =   851968
         _ExtentX        =   6535
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "> Cantidad de meses necesario para los promedios:"
      End
      Begin VB.Label lblDatosPara 
         Caption         =   "Datos para el día: "
         Height          =   225
         Left            =   180
         TabIndex        =   16
         Top             =   810
         Width           =   1395
      End
   End
   Begin XtremeSuiteControls.TabControl tab2 
      Height          =   3555
      Left            =   150
      TabIndex        =   0
      Top             =   6000
      Width           =   19815
      _Version        =   851968
      _ExtentX        =   34951
      _ExtentY        =   6271
      _StockProps     =   68
      ItemCount       =   3
      Item(0).Caption =   "Indicadores Grales"
      Item(0).ControlCount=   6
      Item(0).Control(0)=   "grillaConceptos"
      Item(0).Control(1)=   "vbuscarConcepto"
      Item(0).Control(2)=   "lblBuscarUn"
      Item(0).Control(3)=   "lbli"
      Item(0).Control(4)=   "lble"
      Item(0).Control(5)=   "PushButton2"
      Item(1).Caption =   "Alarmas"
      Item(1).ControlCount=   0
      Item(2).Caption =   "Predicciones"
      Item(2).ControlCount=   0
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   315
         Left            =   18870
         TabIndex        =   40
         Top             =   360
         Width           =   555
         _Version        =   851968
         _ExtentX        =   979
         _ExtentY        =   556
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grillaConceptos 
         Height          =   2715
         Left            =   10620
         TabIndex        =   22
         Top             =   720
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   4789
         _Version        =   393216
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin XtremeSuiteControls.FlatEdit vbuscarConcepto 
         Height          =   285
         Left            =   13110
         TabIndex        =   23
         Top             =   390
         Width           =   5625
         _Version        =   851968
         _ExtentX        =   9922
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.Label lble 
         Height          =   375
         Left            =   8220
         TabIndex        =   27
         Top             =   2940
         Width           =   2265
         _Version        =   851968
         _ExtentX        =   3995
         _ExtentY        =   661
         _StockProps     =   79
         BackColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lbli 
         Height          =   345
         Left            =   8250
         TabIndex        =   26
         Top             =   2490
         Width           =   2265
         _Version        =   851968
         _ExtentX        =   3995
         _ExtentY        =   609
         _StockProps     =   79
         BackColor       =   65280
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblBuscarUn 
         Height          =   195
         Left            =   10650
         TabIndex        =   24
         Top             =   420
         Width           =   2355
         _Version        =   851968
         _ExtentX        =   4154
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Buscar un concepto particular: "
      End
   End
   Begin VB.Menu not 
      Caption         =   "Notificaciones"
      Begin VB.Menu personal 
         Caption         =   "Personal"
      End
      Begin VB.Menu funcionarios 
         Caption         =   "Funcionarios"
      End
      Begin VB.Menu asesores 
         Caption         =   "Asesores"
      End
   End
   Begin VB.Menu configuracion 
      Caption         =   "Configuraciones"
      Begin VB.Menu confconceptos 
         Caption         =   "Conceptos"
      End
      Begin VB.Menu confalarmas 
         Caption         =   "Alarmas"
      End
   End
   Begin VB.Menu informes 
      Caption         =   "Informes"
      Begin VB.Menu infopresidencia 
         Caption         =   "Presidencia"
      End
      Begin VB.Menu infosecretaria 
         Caption         =   "Secretaría"
      End
      Begin VB.Menu inforevisores 
         Caption         =   "Revisores de Ctas"
      End
      Begin VB.Menu infocontador 
         Caption         =   "Contaduría Exterterna"
      End
      Begin VB.Menu infoproveedores 
         Caption         =   "Proveedores"
      End
   End
End
Attribute VB_Name = "frmTablero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vTemp As String
Dim ai(1 To 20) ' importes de los distintos conceptos
Dim ae(1 To 20)
Dim atotal(1 To 2) ' total de ingresos y egresos
Dim asaldos(1 To 20) ' todos los saldo

Dim vgr As Double


Private Sub calendario_ExpandButtonClick()

End Sub

Public Sub Form_Initialize()
g1.Cols = 8
g1.Rows = 20

g1.ColWidth(0) = 100
g1.ColWidth(1) = 2000
g1.ColWidth(2) = 1200
g1.ColWidth(3) = 1200
g1.ColWidth(4) = 1200
g1.ColWidth(5) = 1200
g1.ColWidth(6) = 0
g1.ColWidth(7) = 1000

g1.TextMatrix(0, 1) = "Concepto"
g1.TextMatrix(0, 2) = "Valor"
g1.TextMatrix(0, 3) = "Previsto"

g1.TextMatrix(0, 4) = "Saldo"

g1.TextMatrix(0, 5) = "S.Previsto"

g1.TextMatrix(0, 7) = "%Cobro"



Me.grillaConceptos.ColWidth(0) = 100
Me.grillaConceptos.ColWidth(1) = 2000
Me.grillaConceptos.ColWidth(2) = 3000
Me.grillaConceptos.ColWidth(3) = 1000


tab1.SelectedItem = 0
tab2.SelectedItem = 0


Me.vfecha.Value = Date

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmBancoCajaDetalle.Show
End Sub

Private Sub g1_DblClick()
Dim i As Integer
Dim vsql As String


i = g1.Row

vsql = g1.TextMatrix(i, 6)

Call actualizarGrilla(g4, vsql)

End Sub


Private Sub actualizarGrilla(ByRef g As MSHFlexGrid, ByVal vsql As String)
On Error Resume Next
Dim r As New ADODB.Recordset

r.Open vsql, ConnDDBB, adOpenKeyset, adLockOptimistic

Set g.DataSource = r.DataSource

If Err Then Exit Sub
End Sub

Function getTotales() As Double

Dim i, jj As Integer
Dim vi, vip, ve, vep, vs, vsp

g1.TextMatrix(0, 4) = 0
g1.TextMatrix(0, 5) = 0

jj = 0

barra.Max = Me.g1.Rows - 1
barra.Value = 0

For i = 1 To Me.g1.Rows - 1

            If Not Trim(g1.TextMatrix(i, 1)) = "" Then
            
                                If Val(g1.TextMatrix(i, 2)) > 0 Or Val(g1.TextMatrix(i, 3)) > 0 Then
                                
                                    vi = vi + Val(g1.TextMatrix(i, 2))
                                    vip = vip + Val(g1.TextMatrix(i, 3))
                                    
                                
                                Else
                                
                                    ve = ve + Abs(Val(g1.TextMatrix(i, 2)))
                                    vep = vep + Abs(Val(g1.TextMatrix(i, 3)))
                                  
                                
                                End If
                                                       
                                ai(i) = Val(g1.TextMatrix(i, 2))
                                 
                                g1.TextMatrix(i, 4) = Val(g1.TextMatrix(i, 2)) + Val(g1.TextMatrix(i - 1, 4))
                                g1.TextMatrix(i, 5) = Val(g1.TextMatrix(i, 3)) + Val(g1.TextMatrix(i - 1, 5))
                                    
                                vs = vs + Val(g1.TextMatrix(i, 4))
                                vsp = vsp + Val(g1.TextMatrix(i, 5))
                                
                               If i > 1 Then
                                                If Not Val(g1.TextMatrix(i - 1, 4)) = Val(g1.TextMatrix(i, 4)) Then
                                                 jj = jj + 1
                                                 asaldos(jj) = Val(g1.TextMatrix(i, 4))
                                                End If
                               Else
                                asaldos(1) = Val(g1.TextMatrix(i, 4))
                               End If
            End If

barra.Value = barra.Value + 1

Next

Me.lbli.Caption = Format(vi, "###,###,##0.00")
Me.lble.Caption = Format(ve, "###,###,##0.00")

atotal(1) = vi
atotal(2) = ve

End Function

Private Sub Gastos_Click()
    Call mostrar(" concat(year(t.Fecha),month(t.Fecha)) ", " year(t.Fecha) as A, month(t.Fecha) as M, ")
End Sub


Private Sub mostrar(vgrupo As String, vcampos As String, Optional vorder As String)
Dim vsql As String
Dim r As New ADODB.Recordset
Dim vvgrupo As String

If Trim(vgrupo) = "" Then
    vvgrupo = " "
Else
    vvgrupo = "  group by " + vgrupo
End If


vsql = " SELECT " + _
vcampos + _
" format(sum(total),'###,###,##0.00') as total, " + _
" format(max(total),'###,###,##0.00') as Maximo, " + _
" format(min(total),'###,###,##0.00') as Minimo, " + _
" format(avg(total),'###,###,##0.00') as Promedio, " + _
" format(STD(total),'###,###,##0.00') as Balanceo " + _
" FROM " + _
"  `pfactura` t " + _
vvgrupo + vorder

Call r.Open(vsql, ConnDDBB, adOpenStatic, adLockPessimistic)

Set Me.grillaAlarma.DataSource = r.DataSource

End Sub

Private Sub gra2_DblClick()
'If Err Then gra2.ChartType = 14
End Sub

Private Sub personas_Click()
Call mostrar(" codigo ", " codigo, nombre , ", " order by total desc ")
End Sub

Public Sub PusActualizar_Click()
On Error Resume Next
Dim i As Integer
Dim vIreal, vIprevisto, vEreal, VEprevisto As Double

   Call actualizar
   
    Me.cartel.Top = 2490
    Me.cartel.Visible = True
   

    i = 1
    g1.TextMatrix(i, 1) = " > Disponibilidad :"
    Call pintar(g1, i, 1, vbGreen)
    g1.TextMatrix(i, 2) = getSaldosCaja(vfecha)
    g1.TextMatrix(i, 3) = g1.TextMatrix(i, 2)
    g1.TextMatrix(i, 6) = getSaldosCajaDetalle(vfecha)
    
    ai(i) = g1.TextMatrix(i, 2)
    
    vIreal = g1.TextMatrix(i, 2)
    vIprevisto = vIprevisto + Val(g1.TextMatrix(i, 3))
    
    i = i + 1
    g1.TextMatrix(i, 1) = " > Cobros Urbano :"
    Call pintar(g1, i, 1, vbGreen)
    g1.TextMatrix(i, 3) = getCobrosUrbano(vfecha - (30 * Val(Me.vpromedios)))
    g1.TextMatrix(i, 7) = getCobrosUrbanoDetalle(vfecha - (30 * Val(Me.vpromedios)))
    
    
    ai(i) = g1.TextMatrix(i, 3)
    
    vIprevisto = vIprevisto + Val(g1.TextMatrix(i, 3))
    
    
    
    i = i + 1
    g1.TextMatrix(i, 1) = " > Cobros Rural :"
    Call pintar(g1, i, 1, vbGreen)
    g1.TextMatrix(i, 3) = getCobrosRural(vfecha - (30 * Val(Me.vpromedios)))
    g1.TextMatrix(i, 7) = getCobrosRuralDetalle(vfecha - (30 * Val(Me.vpromedios)))
    
    vIprevisto = vIprevisto + Val(g1.TextMatrix(i, 3))
    ai(i) = g1.TextMatrix(i, 3)
    
    
    
    i = i + 1
    g1.TextMatrix(i, 1) = " > Coopart."
    Call pintar(g1, i, 1, vbGreen)
    g1.TextMatrix(i, 3) = getCooparticipacion(vfecha)
    vIprevisto = vIprevisto + Val(g1.TextMatrix(i, 3))
    
    ai(i) = g1.TextMatrix(i, 2)
     
     
' -------------------- Egresos

    vEreal = 0

    i = i + 1
    g1.TextMatrix(i, 1) = "> Deuda a proveedores"
    g1.TextMatrix(i, 2) = -1 * getSaldoProveedor(vfecha)
    g1.TextMatrix(i, 6) = getSaldoProveedorDetalle(vfecha, Me.vbuscar)
    
    g1.TextMatrix(i, 4) = vIreal + Val(g1.TextMatrix(i, 2))
    g1.TextMatrix(i, 5) = vIprevisto + Val(g1.TextMatrix(i, 2))
    
    Call pintar(g1, i, 1, vbRed)
    ai(i) = g1.TextMatrix(i, 4)
   
    
    
    i = i + 1
    g1.TextMatrix(i, 1) = " > > Deuda vencidas:"
    g1.TextMatrix(i, 2) = -1 * GetDeudasVencidas(vfecha - 30 * Me.vpromediosFactura)
    g1.TextMatrix(i, 4) = vIreal + Val(g1.TextMatrix(i, 2))
    g1.TextMatrix(i, 5) = vIprevisto + Val(g1.TextMatrix(i, 2))
    g1.TextMatrix(i, 6) = GetDeudasVencidasDetalles(vfecha - 30 * Me.vpromediosFactura, Me.vbuscar)
    
    Call pintar(g1, i, 1, vbRed)
    ai(i) = g1.TextMatrix(i, 4)
   

    vEreal = vEreal + g1.TextMatrix(i, 4)
    VEprevisto = VEprevisto + g1.TextMatrix(i, 5)
    

    i = i + 1
    g1.TextMatrix(i, 1) = " > F.A.E."
 
    g1.TextMatrix(i, 2) = -1 * getfae(vfecha)
    g1.TextMatrix(i, 4) = vIreal + Val(g1.TextMatrix(i, 2))
    g1.TextMatrix(i, 5) = vIprevisto + Val(g1.TextMatrix(i, 2))
    ai(i) = g1.TextMatrix(i, 4)
    
    Call pintar(g1, i, 1, vbRed)

    vEreal = vEreal + g1.TextMatrix(i, 4)
    VEprevisto = VEprevisto + g1.TextMatrix(i, 5)
    
    
    i = i + 1
    g1.TextMatrix(i, 1) = " > Eventuales:"
    g1.TextMatrix(i, 2) = -1 * getEventuales(vfecha)
    g1.TextMatrix(i, 4) = vIreal + Val(g1.TextMatrix(i, 2))
    g1.TextMatrix(i, 5) = vIprevisto + Val(g1.TextMatrix(i, 2))
    
    ai(i) = g1.TextMatrix(i, 4)
   
    Call pintar(g1, i, 1, vbRed)

    vEreal = vEreal + g1.TextMatrix(i, 4)
    VEprevisto = VEprevisto + g1.TextMatrix(i, 5)
    
    
    'i = i + 1
    'g1.TextMatrix(i, 1) = " > Ayudas:"
    'g1.TextMatrix(i, 2) = -1 * getCtas(vfecha, "social")
    'g1.TextMatrix(i, 4) = vIreal + Val(g1.TextMatrix(i, 2))
    'g1.TextMatrix(i, 5) = vIprevisto + Val(g1.TextMatrix(i, 2))
    
    'ai(i) = g1.TextMatrix(i, 4)
    
    
    'Call pintar(g1, i, 1, vbRed)

    'vEreal = vEreal + g1.TextMatrix(i, 4)
    'VEprevisto = VEprevisto + g1.TextMatrix(i, 5)
    
    
    atotal(1) = Abs(vIprevisto)
    atotal(2) = Abs(VEprevisto)
    
    Me.lbli.Caption = Format(vIprevisto, "###,###,##0.00")
    Me.lble.Caption = Format(VEprevisto, "###,###,##0.00")
       
       
    vgr = i
    
    Call actualizar
    
    
    Me.cartel.Top = 2490
    Me.cartel.Visible = False
    
If Err Then
    Me.cartel.Top = 2490
    Me.cartel.Visible = False
    Exit Sub
End If
End Sub

Private Sub actualizar()
    
    Call getTotales
    
    Call formatearGrilla
    
    Call actualizargraficas
    
   ' Call actualizarSaldos
    
End Sub

Private Sub actualizarSaldos()
Dim i, j  As Integer
Dim vsaldo As Double
j = 0
For i = 1 To g1.Rows - 1

       vsaldo = (Val(g1.TextMatrix(i, 4)))
       
       If Abs(vsaldo) > 0 Then
        j = j + 1
        asaldos(j) = vsaldo
       End If
       
Next

gra2.ChartData = asaldos

End Sub

Private Sub actualizargraficas()

gra1.ChartData = ai
gra2.ChartData = ai
gra3.ChartData = atotal


Set Me.grillaConceptos.DataSource = getDataSource(fvsqlTCtas(Me.vfecha)).DataSource
Set gra4.DataSource = getDataSource(fvsqlTCtasGraficaI(Me.vfecha))
Set gra5.DataSource = getDataSource(fvsqlTCtasGraficaE(Me.vfecha))



End Sub

Function fvsqlTCtas(vfechas As Date, Optional vbuscar As String) As String
fvsqlTCtas = " SELECT " + _
"  cuentas.CodigoCuenta, " + _
"  cuentas.Cuenta, " + _
"  case  " + _
"       when cuentas.CodigoCuenta like '01%' then format(avg(`asientosdetalle`.`debe`),'###,###,##0.00')  " + _
"       when cuentas.CodigoCuenta like '02%' then  format(-1* avg(`asientosdetalle`.`haber`),'###,###,##0.00') end Saldo " + _
"  FROM " + _
"  `asientos` " + _
"  INNER JOIN `asientosdetalle` ON (`asientos`.`Numero` = `asientosdetalle`.`Numero`) " + _
"  INNER JOIN `cuentas` ON (`asientosdetalle`.`CodigoCuenta` = `cuentas`.`CodigoCuenta`) " + _
" where cuentas.cuenta like '%" + vbuscar + "%' " + _
" group by cuentas.CodigoCuenta "
End Function


Function fvsqlTCtasGraficaI(vfechas As Date) As String
fvsqlTCtasGraficaI = " SELECT " + _
"  case  " + _
"        when cuentas.CodigoCuenta like '01%' then avg(`asientosdetalle`.`debe`) else 0 end I   " + _
"  FROM " + _
"  `asientos` " + _
"  INNER JOIN `asientosdetalle` ON (`asientos`.`Numero` = `asientosdetalle`.`Numero`) " + _
"  INNER JOIN `cuentas` ON (`asientosdetalle`.`CodigoCuenta` = `cuentas`.`CodigoCuenta`) " + _
" group by cuentas.CodigoCuenta having I > 0"
End Function


Function fvsqlTCtasGraficaE(vfechas As Date) As String
fvsqlTCtasGraficaE = " SELECT " + _
"  case  " + _
"        when cuentas.CodigoCuenta like '02%' then avg(`asientosdetalle`.`haber`) else 0 end E   " + _
"  FROM " + _
"  `asientos` " + _
"  INNER JOIN `asientosdetalle` ON (`asientos`.`Numero` = `asientosdetalle`.`Numero`) " + _
"  INNER JOIN `cuentas` ON (`asientosdetalle`.`CodigoCuenta` = `cuentas`.`CodigoCuenta`) " + _
" group by cuentas.CodigoCuenta having E > 0"
End Function

Private Sub formatearGrilla()
Dim i, j  As Integer

For i = 0 To g1.Rows - 1
    
    For j = 1 To g1.Cols - 2
        g1.TextMatrix(i, j) = Format(g1.TextMatrix(i, j), "###,###,##0.00")
    Next
    
Next


End Sub



Private Sub pintar(ByRef g As MSHFlexGrid, ByVal i As Integer, ByVal j As Integer, vColor)
On Error Resume Next

g.Row = i
g.Col = j
g.CellBackColor = vColor

If Err Then Exit Sub
End Sub

Private Sub PusAgregarConcepto_Click()
  Call fbuscarGrilla("(select * from cuentas where Imputable ='S') as t", "Cuenta", "CodigoCuenta", Me.vconcepto.Name, Me)
End Sub

Private Sub PushButton3_Click()

End Sub

Private Sub PushButton1_Click()
Call mostrar(" ", " ", "order by Balanceo desc")
End Sub

Private Sub PusPresupuestos_Click()

End Sub

Private Sub PushButton2_Click()
Set Me.grillaConceptos.DataSource = getDataSource(fvsqlTCtas(Me.vfecha, Me.vbuscarConcepto.Text)).DataSource
End Sub

Private Sub PusPersonas_Click()
Call fbuscarGrilla("proveedores", "Nombre", "Codigo", Me.vcliprovee.Name, Me, , False)
End Sub

Private Sub vbuscarConcepto_Change()
 Set Me.grillaConceptos.DataSource = getDataSource(fvsqlTCtas(Me.vfecha, Me.vbuscarConcepto.Text)).DataSource
End Sub

Private Sub vconcepto_Change()
On Error Resume Next
Dim v As Double

v = getCtas(vfecha, Me.vconcepto.Tag)

vgr = vgr + 1

 g1.TextMatrix(vgr, 1) = vconcepto

If Left(vconcepto.Tag, 2) = "01" Then
 
  g1.TextMatrix(vgr, 2) = v
  Call pintar(g1, vgr, 1, vbGreen)
 
Else

 g1.TextMatrix(vgr, 2) = -1 * v
 Call pintar(g1, vgr, 1, vbRed)
 
End If

 g1.TextMatrix(vgr, 6) = getCtasdetalle(vfecha, Me.vconcepto.Tag)
 
 Call getTotales

If Err Then Exit Sub
End Sub
