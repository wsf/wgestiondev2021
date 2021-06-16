VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmConceptos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión de relacioens entre CONCEPTOS <-> CAJA <-> CTAS. CONTABLES"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   15870
   Begin XtremeSuiteControls.TabControl tab 
      Height          =   7155
      Left            =   60
      TabIndex        =   8
      Top             =   0
      Width           =   15765
      _Version        =   851968
      _ExtentX        =   27808
      _ExtentY        =   12621
      _StockProps     =   68
      ItemCount       =   2
      SelectedItem    =   1
      Item(0).Caption =   "Altas"
      Item(0).ControlCount=   29
      Item(0).Control(0)=   "GroupBox3"
      Item(0).Control(1)=   "GroupBox2"
      Item(0).Control(2)=   "lblCuentaContable"
      Item(0).Control(3)=   "Label1"
      Item(0).Control(4)=   "vcodigoCta"
      Item(0).Control(5)=   "Label3"
      Item(0).Control(6)=   "vbanco"
      Item(0).Control(7)=   "vcodigobanco"
      Item(0).Control(8)=   "vcuenta"
      Item(0).Control(9)=   "lblReferencia"
      Item(0).Control(10)=   "GroupBox4"
      Item(0).Control(11)=   "lblComentarios"
      Item(0).Control(12)=   "GroupBox5"
      Item(0).Control(13)=   "vcomentario"
      Item(0).Control(14)=   "vref"
      Item(0).Control(15)=   "PusBuscarCliente"
      Item(0).Control(16)=   "vconcepto"
      Item(0).Control(17)=   "PushButton4"
      Item(0).Control(18)=   "GroupBox6"
      Item(0).Control(19)=   "GroupBox7"
      Item(0).Control(20)=   "Picture1"
      Item(0).Control(21)=   "Picture2"
      Item(0).Control(22)=   "Picture3"
      Item(0).Control(23)=   "PushButton1"
      Item(0).Control(24)=   "Label2"
      Item(0).Control(25)=   "vcodigoCta2"
      Item(0).Control(26)=   "vcuenta2"
      Item(0).Control(27)=   "vrendicion"
      Item(0).Control(28)=   "PushButton2"
      Item(1).Caption =   "Datos"
      Item(1).ControlCount=   7
      Item(1).Control(0)=   "grilla"
      Item(1).Control(1)=   "lblBuscar"
      Item(1).Control(2)=   "vbucar"
      Item(1).Control(3)=   "GroupBox1"
      Item(1).Control(4)=   "PbAcciones3"
      Item(1).Control(5)=   "PbAcciones2"
      Item(1).Control(6)=   "PushButton3"
      Begin VB.TextBox vrendicion 
         Height          =   315
         Left            =   -64900
         TabIndex        =   42
         Top             =   3900
         Visible         =   0   'False
         Width           =   6555
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   -69880
         Picture         =   "frmConceptos.frx":0000
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   38
         Top             =   3420
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -68620
         Picture         =   "frmConceptos.frx":058A
         ScaleHeight     =   285
         ScaleWidth      =   255
         TabIndex        =   36
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   -69880
         Picture         =   "frmConceptos.frx":0B14
         ScaleHeight     =   285
         ScaleWidth      =   285
         TabIndex        =   35
         Top             =   2910
         Visible         =   0   'False
         Width           =   285
      End
      Begin XtremeSuiteControls.GroupBox GroupBox6 
         Height          =   495
         Left            =   -58180
         TabIndex        =   29
         Top             =   2790
         Visible         =   0   'False
         Width           =   3255
         _Version        =   851968
         _ExtentX        =   5741
         _ExtentY        =   873
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton RadDebe 
            Height          =   315
            Left            =   270
            TabIndex        =   31
            Top             =   150
            Width           =   975
            _Version        =   851968
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Debe"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RadHaber 
            Height          =   315
            Left            =   1800
            TabIndex        =   32
            Top             =   150
            Width           =   975
            _Version        =   851968
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Haber"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.FlatEdit vcomentario 
         Height          =   315
         Left            =   -64900
         TabIndex        =   6
         Top             =   4320
         Visible         =   0   'False
         Width           =   10035
         _Version        =   851968
         _ExtentX        =   17701
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.GroupBox GroupBox4 
         Height          =   225
         Left            =   -68290
         TabIndex        =   25
         Top             =   2100
         Visible         =   0   'False
         Width           =   13455
         _Version        =   851968
         _ExtentX        =   23733
         _ExtentY        =   397
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
      End
      Begin XtremeSuiteControls.FlatEdit vref 
         Height          =   285
         Left            =   -64870
         TabIndex        =   0
         Top             =   1200
         Visible         =   0   'False
         Width           =   3945
         _Version        =   851968
         _ExtentX        =   6959
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   195
         Left            =   60
         TabIndex        =   23
         Top             =   780
         Width           =   15675
         _Version        =   851968
         _ExtentX        =   27649
         _ExtentY        =   344
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
      End
      Begin XtremeSuiteControls.FlatEdit vbucar 
         Height          =   315
         Left            =   960
         TabIndex        =   17
         Top             =   1170
         Width           =   6885
         _Version        =   851968
         _ExtentX        =   12144
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla 
         Height          =   5295
         Left            =   60
         TabIndex        =   15
         Top             =   1710
         Width           =   15585
         _ExtentX        =   27490
         _ExtentY        =   9340
         _Version        =   393216
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).BandIndent=   9
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   135
         Left            =   -69970
         TabIndex        =   9
         Top             =   780
         Visible         =   0   'False
         Width           =   15705
         _Version        =   851968
         _ExtentX        =   27702
         _ExtentY        =   238
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   405
         Left            =   -69940
         TabIndex        =   10
         Top             =   420
         Visible         =   0   'False
         Width           =   1725
         _Version        =   851968
         _ExtentX        =   3043
         _ExtentY        =   714
         _StockProps     =   79
         Appearance      =   2
         BorderStyle     =   2
         Begin XtremeSuiteControls.PushButton PbAcciones 
            Height          =   345
            Index           =   6
            Left            =   90
            TabIndex        =   7
            Top             =   0
            Width           =   1485
            _Version        =   851968
            _ExtentX        =   2619
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "Guardar <F10>"
            UseVisualStyle  =   -1  'True
            Picture         =   "frmConceptos.frx":109E
         End
      End
      Begin XtremeSuiteControls.PushButton PusBuscarCliente 
         Height          =   285
         Left            =   -65710
         TabIndex        =   2
         Top             =   2400
         Visible         =   0   'False
         Width           =   765
         _Version        =   851968
         _ExtentX        =   1349
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "<F2>"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmConceptos.frx":146B
      End
      Begin XtremeSuiteControls.FlatEdit vconcepto 
         Height          =   285
         Left            =   -64870
         TabIndex        =   1
         Top             =   1620
         Visible         =   0   'False
         Width           =   6465
         _Version        =   851968
         _ExtentX        =   11404
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vbanco 
         Height          =   315
         Left            =   -64900
         TabIndex        =   13
         Top             =   2370
         Visible         =   0   'False
         Width           =   3945
         _Version        =   851968
         _ExtentX        =   6959
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vcodigoCta 
         Height          =   315
         Left            =   -60850
         TabIndex        =   14
         Top             =   2940
         Visible         =   0   'False
         Width           =   2475
         _Version        =   851968
         _ExtentX        =   4366
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton PbAcciones3 
         Height          =   345
         Left            =   870
         TabIndex        =   18
         Top             =   420
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Borrar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmConceptos.frx":1A05
      End
      Begin XtremeSuiteControls.PushButton PbAcciones2 
         Height          =   345
         Left            =   1980
         TabIndex        =   19
         Top             =   420
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Modificar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmConceptos.frx":1F9F
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   285
         Left            =   -65680
         TabIndex        =   3
         Top             =   2940
         Visible         =   0   'False
         Width           =   765
         _Version        =   851968
         _ExtentX        =   1349
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "<F3>"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmConceptos.frx":2539
      End
      Begin XtremeSuiteControls.FlatEdit vcuenta 
         Height          =   315
         Left            =   -64900
         TabIndex        =   20
         Top             =   2940
         Visible         =   0   'False
         Width           =   3945
         _Version        =   851968
         _ExtentX        =   6959
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vcodigobanco 
         Height          =   315
         Left            =   -60880
         TabIndex        =   22
         Top             =   2370
         Visible         =   0   'False
         Width           =   2475
         _Version        =   851968
         _ExtentX        =   4366
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.GroupBox GroupBox5 
         Height          =   1995
         Left            =   -68230
         TabIndex        =   27
         Top             =   4920
         Visible         =   0   'False
         Width           =   13515
         _Version        =   851968
         _ExtentX        =   23839
         _ExtentY        =   3519
         _StockProps     =   79
         Caption         =   "Información de aplicación al módulo contable:"
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
         BorderStyle     =   1
         Begin VB.ListBox vcontexto 
            BackColor       =   &H80000004&
            Height          =   1425
            Left            =   3930
            TabIndex        =   37
            Top             =   270
            Width           =   9525
         End
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   345
         Left            =   3090
         TabIndex        =   28
         Top             =   420
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmConceptos.frx":2AD3
      End
      Begin XtremeSuiteControls.GroupBox GroupBox7 
         Height          =   495
         Left            =   -58180
         TabIndex        =   30
         Top             =   2250
         Visible         =   0   'False
         Width           =   3255
         _Version        =   851968
         _ExtentX        =   5741
         _ExtentY        =   873
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton RadDebito 
            Height          =   255
            Left            =   270
            TabIndex        =   33
            Top             =   180
            Width           =   975
            _Version        =   851968
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Debita"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RadCredito 
            Height          =   255
            Left            =   1800
            TabIndex        =   34
            Top             =   180
            Width           =   975
            _Version        =   851968
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Acredita"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.FlatEdit vcodigoCta2 
         Height          =   315
         Left            =   -60850
         TabIndex        =   39
         Top             =   3450
         Visible         =   0   'False
         Width           =   2475
         _Version        =   851968
         _ExtentX        =   4366
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   285
         Left            =   -65680
         TabIndex        =   4
         Top             =   3450
         Visible         =   0   'False
         Width           =   765
         _Version        =   851968
         _ExtentX        =   1349
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "<F4>"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmConceptos.frx":306D
      End
      Begin XtremeSuiteControls.FlatEdit vcuenta2 
         Height          =   315
         Left            =   -64900
         TabIndex        =   40
         Top             =   3450
         Visible         =   0   'False
         Width           =   3945
         _Version        =   851968
         _ExtentX        =   6959
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   315
         Left            =   -66520
         TabIndex        =   5
         Top             =   3900
         Visible         =   0   'False
         Width           =   1575
         _Version        =   851968
         _ExtentX        =   2778
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Rendición  <F9>"
         ForeColor       =   0
         BackColor       =   -2147483644
         UseVisualStyle  =   -1  'True
         Picture         =   "frmConceptos.frx":3607
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Left            =   -69700
         TabIndex        =   41
         Top             =   3450
         Visible         =   0   'False
         Width           =   3885
         _Version        =   851968
         _ExtentX        =   6853
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Seleccionar Cuenta  asociada Personas/Entidades: "
         ForeColor       =   255
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblComentarios 
         Height          =   255
         Left            =   -67000
         TabIndex        =   26
         Top             =   4320
         Visible         =   0   'False
         Width           =   1125
         _Version        =   851968
         _ExtentX        =   1984
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Comentarios: "
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblReferencia 
         Height          =   255
         Left            =   -69820
         TabIndex        =   24
         Top             =   1170
         Visible         =   0   'False
         Width           =   3675
         _Version        =   851968
         _ExtentX        =   6482
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Referencia:"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Left            =   -69700
         TabIndex        =   21
         Top             =   2940
         Visible         =   0   'False
         Width           =   3585
         _Version        =   851968
         _ExtentX        =   6324
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Seleccionar Cuenta Contable asaciada a la Caja:"
         Alignment       =   1
      End
      Begin VB.Label lblBuscar 
         Caption         =   "Buscar:"
         Height          =   285
         Left            =   150
         TabIndex        =   16
         Top             =   1200
         Width           =   705
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Left            =   -69760
         TabIndex        =   12
         Top             =   1620
         Visible         =   0   'False
         Width           =   3675
         _Version        =   851968
         _ExtentX        =   6482
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Conceptos para tansacciones: "
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblCuentaContable 
         Height          =   255
         Left            =   -69790
         TabIndex        =   11
         Top             =   2370
         Visible         =   0   'False
         Width           =   3675
         _Version        =   851968
         _ExtentX        =   6482
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Seleccionar Cajas o Bancos:"
         Alignment       =   1
      End
   End
End
Attribute VB_Name = "frmConceptos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vidArticulos As Long
Dim vulinea As Integer
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()
    init
End Sub

Private Sub init()

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2 - 1000

Me.vcodigoCta.Text = 0


Call configurarGrid

Me.tab.SelectedItem = 0

Me.vcontexto.AddItem "Cuenta auxiliar: " + traerDatos2("select * from ctaauxiliar", "codigocta", pathDBMySQL)

'Me.vnombrecta.Tag = ""
'Me.vnombrecta.Text = ""

'CentrarFormulario (Me)

'set Me.vconcepto.SetFocus

End Sub

Private Sub lblImporte_Click()

End Sub
Private Sub LimpiarCampos()

Me.vref.Text = ""
Me.vconcepto.Text = ""
Me.vcodigobanco.Text = ""
Me.vcomentario.Text = ""
Me.vcuenta.Tag = ""
Me.vcuenta.Text = ""
Me.vcodigoCta.Text = ""
Me.vbanco.Text = ""
Me.vcodigobanco.Text = ""

Me.vcuenta2.Text = ""
Me.vcodigoCta2.Text = ""

Me.vrendicion.Tag = ""
Me.vrendicion.Text = ""



End Sub


Private Sub pintar(i As Integer, g As MSHFlexGrid)
On Error Resume Next
Dim j, k, kk As Integer

k = g.Row
kk = g.Col

g.Row = i

For j = 1 To g.Cols - 1
    g.Col = j
    g.CellBackColor = vbGreen
Next

g.Row = k
g.Col = kk
If Err Then Exit Sub
End Sub


Private Sub grilla_Click()


vidArticulos = grilla.TextMatrix(grilla.Row, 13)

Call pintar(grilla.Row, Me.grilla)

Call despintar(vulinea, Me.grilla)

grilla.CellBackColor = vbRed

vulinea = grilla.Row


End Sub



Private Sub despintar(i As Integer, g As MSHFlexGrid)
On Error Resume Next

Dim j, k, kk As Integer
k = g.Row
kk = g.Col
If i = 0 Then Exit Sub
g.Row = i

For j = 1 To g.Cols - 1
    g.Col = j
    g.CellBackColor = vbWhite
Next

g.Row = k
g.Col = kk

If Err Then Exit Sub
End Sub


'Private Sub g_Click()
'vidArticulos = g.TextMatrix(g.Row, 13)
'
'Call pintar(g.Row, grilla)
'
'Call despintar(vulinea, grilla)
'
'grilla.CellBackColor = vbRed
'
'vulinea = g.Row
'
'End Sub

Private Sub PbAcciones_Click(Index As Integer)
On Error Resume Next
Dim vsql, vcampos, vvalores As String

If Not ValidarCampos Then Exit Sub

vcampos = "`conceptos2`.`ref`,  `conceptos2`.`descripcion`,  `conceptos2`.`idbancos`,`conceptos2`.`comentarios`," & _
"`conceptos2`.`idcuentas`,`conceptos2`.`idcuentas2`" & _
", debe, haber, debito, credito, idrendiciones"

vvalores = "'" + Me.vref.Text + "','" + Me.vconcepto.Text + "','" + Me.vcodigobanco.Text + "','" + Me.vcomentario.Text + "'" & _
"," + Str(Val(Me.vcuenta.Tag)) + "," + Str(Val(Me.vcuenta2.Tag)) & _
"," + strBool(Me.RadDebe.Value) + "," + strBool(Me.RadHaber) + "," + strBool(Me.RadDebito) + "," + strBool(Me.RadCredito) + "," + (Me.vrendicion.Tag)

vsql = "insert into conceptos2 (" + vcampos + ") values (" + vvalores + ")"
Call EjecutarScript(vsql, pathDBMySQL)

Call LimpiarCampos

Me.vconcepto.SetFocus

If Err Then Exit Sub
End Sub
Function ValidarCampos() As Boolean
ValidarCampos = True

If Me.vref.Text = "" Or Me.vconcepto.Text = "" Or Me.vcodigobanco.Text = "" Or Me.vcuenta.Tag = "" Then
    MsgBox "Hay campos mal cargados", vbInformation, "Validación"
    ValidarCampos = False
End If
End Function

Private Sub PbAcciones3_Click()
Dim vsql As String
Dim vidConcepto As Integer

If Not MsgBox("Confirma la opración de eliminación", vbYesNo, "Borrar") = vbYes Then
    Exit Sub
End If


vidConcepto = grilla.TextMatrix(grilla.Row, 8)

vsql = "delete from conceptos2 where idconceptos = " + Str(vidConcepto) + ""

Call EjecutarScript(vsql, pathDBMySQL)

Call vbucar_Change

End Sub

Private Sub PushBuscarCliente_Click()
Call fbuscarGrilla("bancos", "Descripcion", "idBancos", Me.vbanco.Name, Me)  ' ema:
End Sub



Private Sub PusBuscarCliente_Click()
Call fbuscarGrilla("bancos", "Descripcion", "idBancos", Me.vbanco.Name, Me)  ' ema:
End Sub

Private Sub PushButton1_Click()
    Call fbuscarGrilla("(select * from cuentas where Imputable ='S') as t", "Cuenta", "idCuentas", Me.vcuenta2.Name, Me)   ' ema:
End Sub
Private Sub PushButton2_Click()

Call fbuscarGrilla("rendiciones", "nombre", "idrendiciones", Me.vrendicion.Name, Me)    ' ema:

End Sub


Private Sub PushButton3_Click()
    'Call ImprimirGrid2(Me.grilla)
    
    Call imprimirGrilla(Me.grilla, 10)

End Sub

Private Sub PushButton4_Click()
    Call fbuscarGrilla("(select * from cuentas where Imputable ='S') as t", "Cuenta", "idCuentas", Me.vcuenta.Name, Me)  ' ema:
End Sub

Private Sub tab_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Call vbucar_Change
End Sub

Private Sub vbanco_Change()
Dim vsql As String

'vsql = "select idbancos from bancos t where t.idcuentas = " + Me.vcuenta.Tag
'Me.vCodigoCuenta = traerDatos2(vsql, "codigocuenta", pathDBMySQL)
If Me.vbanco.Text = "" Then Exit Sub
  
 Me.vcodigobanco.Text = Me.vbanco.Tag
  
End Sub

Private Sub vbucar_Change()
On Error Resume Next
Dim vsql, vwhere, vcampo  As String

vcampo = "`conceptos2`.`ref`,  `conceptos2`.`descripcion`,  `conceptos2`.`idbancos`,`bancos`.`Descripcion`,conceptos2.debito as D,conceptos2.credito as C,`cuentas`.`CodigoCuenta`,`cuentas`.`Cuenta`, conceptos2.debe as DD, conceptos2.haber as H,rendiciones.nombre,`conceptos2`.`comentarios`,`conceptos2`.`idconceptos`,`conceptos2`.`idcuentas`,c2.codigocuenta,c2.cuenta"

Dim rsPresupuesto As New ADODB.Recordset

vwhere = " where conceptos2.ref  like '%" + vbucar + "%' or conceptos2.descripcion like '%" + vbucar + "%' or rendiciones.nombre like '%" + vbucar + "%'"

vsql = "select " + vcampo + " From " + _
  " `conceptos2` " + _
  " left JOIN `cuentas` ON (`conceptos2`.`idcuentas` = `cuentas`.`idCuentas`) " + _
  " left JOIN `cuentas` as c2 ON (`conceptos2`.`idcuentas2` = `c2`.`idCuentas`) " + _
  " left JOIN `bancos` ON (`conceptos2`.`idbancos` = `bancos`.`idBancos`) " + _
  " left JOIN `rendiciones` ON (`conceptos2`.`idrendiciones` = `rendiciones`.`idrendiciones`) " + vwhere
    
    
    With rsPresupuesto
        Call .Open(vsql, ConnDDBB, adOpenStatic, adLockPessimistic)
        Set grilla.DataSource = .DataSource
    End With
        
End Sub

Private Sub vcodigoCta_Change()
'Me.vimporte.SetFocus
End Sub

Private Sub configurarGrid()

Me.grilla.Cols = 17

Me.grilla.ColWidth(0) = 100
Me.grilla.ColWidth(1) = 1000
Me.grilla.ColWidth(2) = 3000
Me.grilla.ColWidth(3) = 1000
Me.grilla.ColWidth(4) = 3000



Me.grilla.ColWidth(5) = 190
Me.grilla.ColWidth(6) = 190

Me.grilla.ColWidth(7) = 1000
Me.grilla.ColWidth(8) = 3000


Me.grilla.ColWidth(9) = 190
Me.grilla.ColWidth(10) = 190

Me.grilla.ColWidth(11) = 1000

Me.grilla.ColWidth(12) = 4000
Me.grilla.ColWidth(13) = 10
Me.grilla.ColWidth(14) = 10

Me.grilla.ColWidth(15) = 1000
Me.grilla.ColWidth(16) = 4000

End Sub

Private Sub vnombrecta_Change()
Dim vsql As String

'vsql = "select codigocuenta from cuentas t where t.idcuentas = " + Me.vnombrecta.Tag
'Me.vcodigoCta = traerDatos2(vsql, "codigocuenta", pathDBMySQL)
    
    
End Sub

Private Sub vcuenta_Change()
Dim vsql As String

If Me.vcuenta.Text = "" Then Exit Sub


vsql = "select codigocuenta from cuentas t where t.idcuentas = " + Me.vcuenta.Tag
Me.vcodigoCta.Text = traerDatos2(vsql, "codigocuenta", pathDBMySQL)
        

End Sub

Private Sub vcuenta2_Change()
Dim vsql As String

If vcuenta2.Text = "" Then Exit Sub

vsql = "select codigocuenta from cuentas t where t.idcuentas = " + Me.vcuenta.Tag
Me.vcodigoCta2.Text = traerDatos2(vsql, "codigocuenta", pathDBMySQL)
        

End Sub

