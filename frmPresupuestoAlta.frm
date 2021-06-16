VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmPresupuesto 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   14625
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.TabControl tab 
      Height          =   7155
      Left            =   60
      TabIndex        =   4
      Top             =   30
      Width           =   14595
      _Version        =   851968
      _ExtentX        =   25744
      _ExtentY        =   12621
      _StockProps     =   68
      ItemCount       =   2
      SelectedItem    =   1
      Item(0).Caption =   "Altas"
      Item(0).ControlCount=   15
      Item(0).Control(0)=   "GroupBox3"
      Item(0).Control(1)=   "GroupBox2"
      Item(0).Control(2)=   "lblCuentaContable"
      Item(0).Control(3)=   "PusBuscarCliente"
      Item(0).Control(4)=   "Label1"
      Item(0).Control(5)=   "PushButton1"
      Item(0).Control(6)=   "vperiodo"
      Item(0).Control(7)=   "vnombrecta"
      Item(0).Control(8)=   "vcodigoCta"
      Item(0).Control(9)=   "vprevisionesSemanales"
      Item(0).Control(10)=   "Label2"
      Item(0).Control(11)=   "GroupBox1"
      Item(0).Control(12)=   "vhistorial"
      Item(0).Control(13)=   "Label5"
      Item(0).Control(14)=   "PushButton3"
      Item(1).Caption =   "Datos"
      Item(1).ControlCount=   11
      Item(1).Control(0)=   "grilla"
      Item(1).Control(1)=   "lblBuscar"
      Item(1).Control(2)=   "vbuscar"
      Item(1).Control(3)=   "PbAcciones2"
      Item(1).Control(4)=   "PbAcciones3"
      Item(1).Control(5)=   "PusImprimir"
      Item(1).Control(6)=   "Exportar"
      Item(1).Control(7)=   "PushButton4"
      Item(1).Control(8)=   "vperiodobuscar"
      Item(1).Control(9)=   "Label7"
      Item(1).Control(10)=   "Shape1"
      Begin XtremeSuiteControls.FlatEdit vhistorial 
         Height          =   375
         Left            =   -68920
         TabIndex        =   26
         Top             =   4860
         Visible         =   0   'False
         Width           =   12195
         _Version        =   851968
         _ExtentX        =   21511
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1815
         Left            =   -69010
         TabIndex        =   20
         Top             =   1980
         Visible         =   0   'False
         Width           =   12315
         _Version        =   851968
         _ExtentX        =   21722
         _ExtentY        =   3201
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.FlatEdit vimporte 
            Height          =   315
            Left            =   1950
            TabIndex        =   21
            Top             =   210
            Width           =   2505
            _Version        =   851968
            _ExtentX        =   4419
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vnuevo_importe 
            Height          =   315
            Left            =   7530
            TabIndex        =   23
            Top             =   210
            Width           =   2505
            _Version        =   851968
            _ExtentX        =   4419
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.PushButton PushButton2 
            Height          =   405
            Left            =   150
            TabIndex        =   25
            Top             =   1320
            Width           =   12045
            _Version        =   851968
            _ExtentX        =   21246
            _ExtentY        =   714
            _StockProps     =   79
            Caption         =   "Actualizar Registro de  Historial"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit vampliar 
            Height          =   315
            Left            =   7530
            TabIndex        =   29
            Top             =   570
            Width           =   2505
            _Version        =   851968
            _ExtentX        =   4419
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   -2147483635
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vreducir 
            Height          =   315
            Left            =   7530
            TabIndex        =   31
            Top             =   930
            Width           =   2505
            _Version        =   851968
            _ExtentX        =   4419
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   255
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.Label Label6 
            Height          =   225
            Left            =   5250
            TabIndex        =   32
            Top             =   960
            Width           =   2085
            _Version        =   851968
            _ExtentX        =   3678
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "Importe de reducción:"
            Alignment       =   1
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   225
            Left            =   5250
            TabIndex        =   30
            Top             =   600
            Width           =   2085
            _Version        =   851968
            _ExtentX        =   3678
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "Importe de ampliación:"
            Alignment       =   1
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   225
            Left            =   5280
            TabIndex        =   24
            Top             =   240
            Width           =   2085
            _Version        =   851968
            _ExtentX        =   3678
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "Nuevo importe:"
            Alignment       =   1
         End
         Begin XtremeSuiteControls.Label lblImporte 
            Height          =   225
            Left            =   90
            TabIndex        =   22
            Top             =   240
            Width           =   1605
            _Version        =   851968
            _ExtentX        =   2831
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "Importe Presupuestado Actual:"
            Alignment       =   1
         End
      End
      Begin XtremeSuiteControls.PushButton Exportar 
         Height          =   345
         Left            =   4560
         TabIndex        =   17
         Top             =   930
         Width           =   1515
         _Version        =   851968
         _ExtentX        =   2672
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Exportar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit vbuscar 
         Height          =   315
         Left            =   1200
         TabIndex        =   13
         Top             =   450
         Width           =   13275
         _Version        =   851968
         _ExtentX        =   23416
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla 
         Height          =   5595
         Left            =   60
         TabIndex        =   11
         Top             =   1410
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   9869
         _Version        =   393216
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   135
         Left            =   -69970
         TabIndex        =   5
         Top             =   780
         Visible         =   0   'False
         Width           =   14535
         _Version        =   851968
         _ExtentX        =   25638
         _ExtentY        =   238
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   465
         Left            =   -69910
         TabIndex        =   6
         Top             =   390
         Visible         =   0   'False
         Width           =   12105
         _Version        =   851968
         _ExtentX        =   21352
         _ExtentY        =   820
         _StockProps     =   79
         Appearance      =   2
         BorderStyle     =   2
         Begin XtremeSuiteControls.PushButton PbAcciones 
            Height          =   345
            Index           =   0
            Left            =   30
            TabIndex        =   0
            Top             =   60
            Width           =   1455
            _Version        =   851968
            _ExtentX        =   2566
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "Guardar <F2>"
            UseVisualStyle  =   -1  'True
            Picture         =   "frmPresupuestoAlta.frx":0000
         End
      End
      Begin XtremeSuiteControls.PushButton PusBuscarCliente 
         Height          =   315
         Left            =   -67960
         TabIndex        =   2
         Top             =   1560
         Visible         =   0   'False
         Width           =   765
         _Version        =   851968
         _ExtentX        =   1349
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "<F3>"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmPresupuestoAlta.frx":03CD
         ImageGap        =   2
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   315
         Left            =   -67930
         TabIndex        =   3
         Top             =   1140
         Visible         =   0   'False
         Width           =   765
         _Version        =   851968
         _ExtentX        =   1349
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "<F1>"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmPresupuestoAlta.frx":0967
      End
      Begin XtremeSuiteControls.FlatEdit vperiodo 
         Height          =   285
         Left            =   -67000
         TabIndex        =   1
         Top             =   1140
         Visible         =   0   'False
         Width           =   5085
         _Version        =   851968
         _ExtentX        =   8969
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vnombrecta 
         Height          =   285
         Left            =   -67000
         TabIndex        =   9
         Top             =   1560
         Visible         =   0   'False
         Width           =   5085
         _Version        =   851968
         _ExtentX        =   8969
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vcodigoCta 
         Height          =   285
         Left            =   -61720
         TabIndex        =   10
         Top             =   1560
         Visible         =   0   'False
         Width           =   3315
         _Version        =   851968
         _ExtentX        =   5847
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483644
         BackColor       =   -2147483644
         Appearance      =   3
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton PbAcciones2 
         Height          =   345
         Left            =   1200
         TabIndex        =   14
         Top             =   930
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Borrar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmPresupuestoAlta.frx":0F01
      End
      Begin XtremeSuiteControls.PushButton PbAcciones3 
         Height          =   345
         Left            =   2310
         TabIndex        =   15
         Top             =   930
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Modificar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmPresupuestoAlta.frx":149B
      End
      Begin XtremeSuiteControls.PushButton PusImprimir 
         Height          =   345
         Left            =   3420
         TabIndex        =   16
         Top             =   930
         Width           =   1035
         _Version        =   851968
         _ExtentX        =   1826
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmPresupuestoAlta.frx":1A35
      End
      Begin XtremeSuiteControls.FlatEdit vprevisionesSemanales 
         Height          =   285
         Left            =   -67090
         TabIndex        =   18
         Top             =   3900
         Visible         =   0   'False
         Width           =   2505
         _Version        =   851968
         _ExtentX        =   4419
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   345
         Left            =   -58030
         TabIndex        =   28
         Top             =   4350
         Visible         =   0   'False
         Width           =   1245
         _Version        =   851968
         _ExtentX        =   2196
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Actualizar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   285
         Left            =   12060
         TabIndex        =   33
         Top             =   930
         Width           =   765
         _Version        =   851968
         _ExtentX        =   1349
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "<F1>"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit vperiodobuscar 
         Height          =   285
         Left            =   12900
         TabIndex        =   34
         Top             =   900
         Width           =   1485
         _Version        =   851968
         _ExtentX        =   2619
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.Label Label7 
         Height          =   255
         Left            =   10980
         TabIndex        =   35
         Top             =   930
         Width           =   915
         _Version        =   851968
         _ExtentX        =   1614
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Período:"
         ForeColor       =   4210752
         BackColor       =   -2147483644
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'Transparent
         Height          =   495
         Left            =   10830
         Top             =   810
         Width           =   3645
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   375
         Left            =   -68920
         TabIndex        =   27
         Top             =   4470
         Visible         =   0   'False
         Width           =   1635
         _Version        =   851968
         _ExtentX        =   2884
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Historial presupestario: "
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   225
         Left            =   -69310
         TabIndex        =   19
         Top             =   3930
         Visible         =   0   'False
         Width           =   1995
         _Version        =   851968
         _ExtentX        =   3519
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Previsiones Semanales:"
         Alignment       =   1
      End
      Begin VB.Label lblBuscar 
         Caption         =   "Buscar:"
         Height          =   285
         Left            =   180
         TabIndex        =   12
         Top             =   510
         Width           =   705
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Left            =   -69790
         TabIndex        =   8
         Top             =   1140
         Visible         =   0   'False
         Width           =   1575
         _Version        =   851968
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Período:"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblCuentaContable 
         Height          =   255
         Left            =   -69790
         TabIndex        =   7
         Top             =   1560
         Visible         =   0   'False
         Width           =   1605
         _Version        =   851968
         _ExtentX        =   2831
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Seleccionar cuenta:"
         Alignment       =   1
      End
   End
End
Attribute VB_Name = "frmPresupuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim camposTamano(8) As Integer
Dim campos(8) As String
Dim vMod As Boolean
Dim r As Integer


Private Sub Exportar_Click()
On Error Resume Next
  Call grillaToExcel(Me.grilla)
If Err Then Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF1 Then Call PushButton1_Click

If KeyCode = vbKeyF2 Then Call PbAcciones_Click(0)

If KeyCode = vbKeyF3 Then Call PusBuscarCliente_Click

End Sub

Private Sub Form_Load()
init
End Sub
Private Sub FormatoGrilla()
Dim i As Integer

With grilla

    For i = 0 To 8
        .ColWidth(i) = camposTamano(i)
        .TextMatrix(0, i) = campos(i)
    Next

End With

End Sub

Private Sub init()

Me.grilla.Cols = 9

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2 - 1000

 'cuentas.CodigoCuenta,
 'cuentas.Cuenta,
 'cuentas.Imputable,
 'presupuesto.periodo,
 'presupuesto.importe
 

Me.vcodigoCta.Text = 0
Me.vnombrecta.Tag = ""
Me.vnombrecta.Text = ""

camposTamano(0) = 300
camposTamano(1) = 2500
camposTamano(2) = 4000
camposTamano(3) = 600
camposTamano(4) = 1200
camposTamano(5) = 1200
camposTamano(6) = 0
camposTamano(7) = 0
camposTamano(8) = 4000

campos(0) = " "
campos(1) = "CodigoCuenta"
campos(2) = "Cuenta"
campos(3) = "Imp"
campos(4) = "Periodo"
campos(5) = "Importe"
campos(6) = "idPresupuesto"
campos(7) = "idCta"
campos(8) = "Apliaciones"


Me.tab.SelectedItem = 0

'CentrarFormulario (Me)
End Sub

Private Sub grilla_Click()

r = grilla.Row
End Sub

Private Sub grilla_DblClick()
r = grilla.Row
Call PbAcciones3_Click
End Sub

Private Sub PbAcciones_Click(Index As Integer)
On Error Resume Next
Dim vsql, vcampos, vvalores, vupdate As String
Dim vvimporte As Double

If Val(Me.vnuevo_importe) > 0 Then
         vvimporte = vnuevo_importe
        
    Else
        vvimporte = vimporte
End If



vcampos = "idcuentas,periodo,importe"
vvalores = Str(Val(Me.vnombrecta.Tag)) + ",'" + Me.vperiodo + "'," + Me.vimporte
vupdate = "idcuentas=" + Str(Val(Me.vnombrecta.Tag)) + "," + "periodo='" + Me.vperiodo + "', importe=" + Str(vvimporte) + ", historial='" + vhistorial + "'"

If Not ValidarCampos Then Exit Sub

If vMod = True Then
    vsql = "update  presupuesto set " + vupdate + " where idpresupuesto=" + Str(Val(Me.vcodigoCta.Tag))
    Call EjecutarScript(vsql, pathDBMySQL)
    ' no permito modificaciones
Else
    vsql = "insert into presupuesto (" + vcampos + ") values (" + vvalores + ")"
    Call EjecutarScript(vsql, pathDBMySQL)
End If




Call LimpiarCampos

vMod = False

If Err Then Exit Sub
End Sub

Function ValidarCampos() As Boolean
ValidarCampos = True
If Me.vperiodo.Text = "" Or Me.vnombrecta.Text = "" Or Me.vnombrecta.Tag = "" Or Me.vimporte.Text = "" Then
    MsgBox "Hay campos mal cargados", vbInformation, "Validación"
    ValidarCampos = False
End If
End Function

Private Sub LimpiarCampos()
'Me.vperiodo.Text = ""
Me.vnombrecta.Text = ""
Me.vcodigoCta.Text = ""
Me.vimporte.Text = ""
Me.vnuevo_importe.Text = ""
Me.vreducir.Text = ""
Me.vhistorial.Text = ""
vampliar.Text = ""

Me.vperiodo.SetFocus
End Sub


Private Sub PbAcciones2_Click()
Dim v As String


v = "delete from presupuesto where idpresupuesto = " + grilla.TextMatrix(r, 6)

If MsgBox("Está seguro de borrar la línea", vbYesNo) = vbYes Then
    Call EjecutarScript(v, pathDBMySQL)
End If

Me.vbuscar = ""
Call vbuscar_Change

End Sub

Private Sub PbAcciones3_Click()

vMod = True

'vidpresupuesto = grilla.TextMatrix(r, 5)
Me.vnombrecta.Tag = grilla.TextMatrix(r, 7) ' idctacontable
Me.vcodigoCta.Text = grilla.TextMatrix(r, 1)  'CodigoCuenta
Me.vnombrecta = grilla.TextMatrix(r, 2) 'Cuenta
Me.vperiodo = grilla.TextMatrix(r, 4) ' Periodo
Me.vimporte = grilla.TextMatrix(r, 5) ' Importe
Me.vcodigoCta.Tag = grilla.TextMatrix(r, 6) 'idPresupuesto
Me.vhistorial = grilla.TextMatrix(r, 8)  'historial

Call vnombrecta_Change

Me.tab.SelectedItem = 0

End Sub

Private Sub PusBuscarCliente_Click()
Call fbuscarGrilla("(select * from cuentas where Imputable ='S') as t", "Cuenta", "idCuentas", Me.vnombrecta.Name, Me)  ' ema:
End Sub

Private Sub PushButton1_Click()
Call fbuscarGrilla("balances", "codigo", "idBalances", Me.vperiodo.Name, Me)     ' ema:
End Sub

Private Sub PushButton2_Click()
Dim vvhistorial As String


If MsgBox("Confirma la actualiazación presupuestaria ?") = vbNo Then Exit Sub

vvhistorial = Chr(13) & Chr(10) + " > Ant: " + Format(Me.vimporte, "###,###,###.00")

vhistorial = vhistorial + vvhistorial

End Sub


Private Sub historicoActualizar()
'Me.grillaHistorialPresupuesto.Clear

'Me.grillaHistorialPresupuesto.AddItem Me.vhistorial

'Me.vlog3.Caption = Me.vhistorial

End Sub


Private Sub PushButton3_Click()
    Call historicoActualizar
End Sub

Private Sub PushButton4_Click()
Call fbuscarGrilla("balances", "codigo", "idBalances", Me.vperiodobuscar.Name, Me)     ' ema:
End Sub

Private Sub PusImprimir_Click()
Call imprimirGrilla(Me.grilla, 10)

Exit Sub

With Mantenimiento
    Set .rsPresupuesto.DataSource = grilla.DataSource
    drpresupuesto.Show
End With
End Sub

Private Sub tab_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Call vbuscar_Change
End Sub

Private Sub vampliar_Change()
    Me.vnuevo_importe.Text = Val(Replace(Me.vimporte, ",", "")) + Val(Me.vampliar)
End Sub

Private Sub vbuscar_Change()
On Error Resume Next

Dim vsql, vwhere, vcampo, vperiodo  As String



If vperiodobuscar = "" Then
    vperiodo = ""
Else
    vperiodo = " and presupuesto.periodo = '" + vperiodobuscar + "'"
End If


vcampo = "cuentas.CodigoCuenta, cuentas.Cuenta, cuentas.Imputable,presupuesto.periodo, format(presupuesto.importe,'###,###,##0.00') as Importe, presupuesto.idpresupuesto,presupuesto.idcuentas,presupuesto.historial as Ampliaciones"
'vcampo = "cuentas.CodigoCuenta, cuentas.Cuenta, cuentas.Imputable,presupuesto.periodo, presupuesto.importe  as Importe, presupuesto.idpresupuesto,presupuesto.idcuentas,presupuesto.historial as Ampliaciones"
vcampo = "cuentas.CodigoCuenta, cuentas.Cuenta, cuentas.Imputable,presupuesto.periodo, format(presupuesto.importe,2) as Importe, presupuesto.idpresupuesto,presupuesto.idcuentas,presupuesto.historial as Ampliaciones"

Dim rsPresupuesto As New ADODB.Recordset

vwhere = " where cuentas.CodigoCuenta  like '%" + vbuscar + "%' or cuentas.Cuenta like '%" + vbuscar + "%' or presupuesto.periodo like '%" + vbuscar + "%' " + vperiodo

vsql = "SELECT " + vcampo + " From `cuentas` INNER JOIN `presupuesto` ON (`cuentas`.`idCuentas` = `presupuesto`.`idcuentas`) " + vwhere
    
    
    With rsPresupuesto
        Call .Open(vsql, ConnDDBB, adOpenStatic, adLockPessimistic)
        Set grilla.DataSource = .DataSource
    End With
    
Call FormatoGrilla
        
End Sub

Private Sub vcodigoCta_Change()
'Me.vimporte.SetFocus
End Sub

Public Sub vnombrecta_Change()
Dim vsql As String

vsql = "select codigocuenta from cuentas t where t.idcuentas = " + Me.vnombrecta.Tag
Me.vcodigoCta = traerDatos2(vsql, "codigocuenta", pathDBMySQL)
        
    
End Sub

Private Sub vreducir_Change()
    Me.vnuevo_importe.Text = Val(Replace(Me.vimporte, ",", "")) - Val(Me.vreducir)
End Sub
