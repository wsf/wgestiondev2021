VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmGastosFijos 
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
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   14595
      _Version        =   851968
      _ExtentX        =   25744
      _ExtentY        =   12621
      _StockProps     =   68
      ItemCount       =   2
      Item(0).Caption =   "Altas"
      Item(0).ControlCount=   13
      Item(0).Control(0)=   "GroupBox3"
      Item(0).Control(1)=   "GroupBox2"
      Item(0).Control(2)=   "lblCuentaContable"
      Item(0).Control(3)=   "PusBuscarCliente"
      Item(0).Control(4)=   "lblImporte"
      Item(0).Control(5)=   "vimporte"
      Item(0).Control(6)=   "Label1"
      Item(0).Control(7)=   "PushButton1"
      Item(0).Control(8)=   "vperiodo"
      Item(0).Control(9)=   "vnombrecta"
      Item(0).Control(10)=   "vcodigoCta"
      Item(0).Control(11)=   "vprevisionesSemanales"
      Item(0).Control(12)=   "Label2"
      Item(1).Caption =   "Datos"
      Item(1).ControlCount=   7
      Item(1).Control(0)=   "grilla"
      Item(1).Control(1)=   "lblBuscar"
      Item(1).Control(2)=   "vbuscar"
      Item(1).Control(3)=   "PbAcciones2"
      Item(1).Control(4)=   "PbAcciones3"
      Item(1).Control(5)=   "PusImprimir"
      Item(1).Control(6)=   "Exportar"
      Begin XtremeSuiteControls.PushButton Exportar 
         Height          =   345
         Left            =   -65440
         TabIndex        =   19
         Top             =   930
         Visible         =   0   'False
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
         Left            =   -68800
         TabIndex        =   15
         Top             =   450
         Visible         =   0   'False
         Width           =   13275
         _Version        =   851968
         _ExtentX        =   23416
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla 
         Height          =   5595
         Left            =   -69910
         TabIndex        =   13
         Top             =   1350
         Visible         =   0   'False
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   9869
         _Version        =   393216
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   135
         Left            =   30
         TabIndex        =   6
         Top             =   780
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
         Left            =   120
         TabIndex        =   7
         Top             =   360
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
            TabIndex        =   1
            Top             =   60
            Width           =   1455
            _Version        =   851968
            _ExtentX        =   2566
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "Guardar <F2>"
            UseVisualStyle  =   -1  'True
            Picture         =   "frmgGastosFijos.frx":0000
         End
      End
      Begin XtremeSuiteControls.PushButton PusBuscarCliente 
         Height          =   315
         Left            =   2040
         TabIndex        =   3
         Top             =   1560
         Width           =   765
         _Version        =   851968
         _ExtentX        =   1349
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "<F3>"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmgGastosFijos.frx":03CD
         ImageGap        =   2
      End
      Begin XtremeSuiteControls.FlatEdit vimporte 
         Height          =   315
         Left            =   3000
         TabIndex        =   0
         Top             =   1920
         Width           =   2505
         _Version        =   851968
         _ExtentX        =   4419
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   315
         Left            =   2070
         TabIndex        =   4
         Top             =   1140
         Width           =   765
         _Version        =   851968
         _ExtentX        =   1349
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "<F1>"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmgGastosFijos.frx":0967
      End
      Begin XtremeSuiteControls.FlatEdit vperiodo 
         Height          =   285
         Left            =   3000
         TabIndex        =   2
         Top             =   1140
         Width           =   5085
         _Version        =   851968
         _ExtentX        =   8969
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vnombrecta 
         Height          =   285
         Left            =   3000
         TabIndex        =   11
         Top             =   1560
         Width           =   5085
         _Version        =   851968
         _ExtentX        =   8969
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vcodigoCta 
         Height          =   285
         Left            =   8280
         TabIndex        =   12
         Top             =   1560
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
         Left            =   -68800
         TabIndex        =   16
         Top             =   930
         Visible         =   0   'False
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Borrar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmgGastosFijos.frx":0F01
      End
      Begin XtremeSuiteControls.PushButton PbAcciones3 
         Height          =   345
         Left            =   -67690
         TabIndex        =   17
         Top             =   930
         Visible         =   0   'False
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Modificar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmgGastosFijos.frx":149B
      End
      Begin XtremeSuiteControls.PushButton PusImprimir 
         Height          =   345
         Left            =   -66580
         TabIndex        =   18
         Top             =   930
         Visible         =   0   'False
         Width           =   1035
         _Version        =   851968
         _ExtentX        =   1826
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmgGastosFijos.frx":1A35
      End
      Begin XtremeSuiteControls.FlatEdit vprevisionesSemanales 
         Height          =   315
         Left            =   3000
         TabIndex        =   20
         Top             =   2340
         Width           =   2505
         _Version        =   851968
         _ExtentX        =   4419
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   225
         Left            =   780
         TabIndex        =   21
         Top             =   2400
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
         Left            =   -69820
         TabIndex        =   14
         Top             =   510
         Visible         =   0   'False
         Width           =   705
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1140
         Width           =   1575
         _Version        =   851968
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Documento"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblImporte 
         Height          =   225
         Left            =   1170
         TabIndex        =   9
         Top             =   2010
         Width           =   1605
         _Version        =   851968
         _ExtentX        =   2831
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Importe:"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblCuentaContable 
         Height          =   255
         Left            =   210
         TabIndex        =   8
         Top             =   1560
         Width           =   1605
         _Version        =   851968
         _ExtentX        =   2831
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cuanto"
         Alignment       =   1
      End
   End
End
Attribute VB_Name = "frmGastosFijos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim camposTamano(7) As Integer
Dim campos(7) As String
Dim vMod As Boolean
Dim r As Integer


Private Sub Exportar_Click()
Call generarExcel("archivo.xls", Me.grilla)
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

    For i = 0 To 7
        .ColWidth(i) = camposTamano(i)
        .TextMatrix(0, i) = campos(i)
    Next

End With

End Sub

Private Sub init()

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
camposTamano(4) = 1000
camposTamano(5) = 1000
camposTamano(6) = 0
camposTamano(7) = 0

campos(0) = " "
campos(1) = "CodigoCuenta"
campos(2) = "Cuenta"
campos(3) = "Imp"
campos(4) = "Periodo"
campos(5) = "Importe"
campos(6) = "idPresupuesto"
campos(7) = "idCta"



Me.tab.SelectedItem = 0

'CentrarFormulario (Me)
End Sub

Private Sub grilla_Click()
r = grilla.Row
End Sub

Private Sub grilla_DblClick()
Call PbAcciones3_Click
End Sub

Private Sub PbAcciones_Click(Index As Integer)
On Error Resume Next
Dim vsql, vcampos, vvalores, vupdate As String


vcampos = "idcuentas,periodo,importe"
vvalores = Str(Val(Me.vnombrecta.Tag)) + ",'" + Me.vperiodo + "'," + Me.vimporte
vupdate = "idcuentas=" + Str(Val(Me.vnombrecta.Tag)) + "," + "periodo='" + Me.vperiodo + "', importe=" + Me.vimporte

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
Call vnombrecta_Change

Me.tab.SelectedItem = 0

End Sub

Private Sub PusBuscarCliente_Click()
Call fbuscarGrilla("(select * from cuentas where Imputable ='S') as t", "Cuenta", "idCuentas", Me.vnombrecta.Name, Me)  ' ema:
End Sub

Private Sub PushButton1_Click()
Call fbuscarGrilla("balances", "codigo", "idBalances", Me.vperiodo.Name, Me)     ' ema:
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

Private Sub vbuscar_Change()
On Error Resume Next
Dim vsql, vwhere, vcampo  As String

vcampo = "cuentas.CodigoCuenta, cuentas.Cuenta, cuentas.Imputable,presupuesto.periodo, format(presupuesto.importe,'########0,00') as Importe, presupuesto.idpresupuesto,presupuesto.idcuentas"

Dim rsPresupuesto As New ADODB.Recordset

vwhere = " where cuentas.CodigoCuenta  like '%" + vbuscar + "%' or cuentas.Cuenta like '%" + vbuscar + "%' or presupuesto.periodo like '%" + vbuscar + "%' "

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
