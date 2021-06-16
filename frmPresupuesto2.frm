VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmPresupuesto2 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.TabControl tab 
      Height          =   7155
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   11925
      _Version        =   851968
      _ExtentX        =   21034
      _ExtentY        =   12621
      _StockProps     =   68
      ItemCount       =   2
      Item(0).Caption =   "Altas"
      Item(0).ControlCount=   12
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
      Item(0).Control(11)=   "Label2"
      Item(1).Caption =   "Datos"
      Item(1).ControlCount=   6
      Item(1).Control(0)=   "grilla"
      Item(1).Control(1)=   "lblBuscar"
      Item(1).Control(2)=   "vbucar"
      Item(1).Control(3)=   "PbAcciones(0)"
      Item(1).Control(4)=   "PbAcciones(1)"
      Item(1).Control(5)=   "PbAcciones(2)"
      Begin XtremeSuiteControls.FlatEdit vbucar 
         Height          =   315
         Left            =   -69010
         TabIndex        =   15
         Top             =   450
         Visible         =   0   'False
         Width           =   10725
         _Version        =   851968
         _ExtentX        =   18918
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla 
         Height          =   5595
         Left            =   -69940
         TabIndex        =   13
         Top             =   1380
         Visible         =   0   'False
         Width           =   11625
         _ExtentX        =   20505
         _ExtentY        =   9869
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   135
         Left            =   30
         TabIndex        =   1
         Top             =   780
         Width           =   11985
         _Version        =   851968
         _ExtentX        =   21140
         _ExtentY        =   238
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   12105
         _Version        =   851968
         _ExtentX        =   21352
         _ExtentY        =   1085
         _StockProps     =   79
         Appearance      =   2
         BorderStyle     =   2
         Begin XtremeSuiteControls.PushButton PbAcciones 
            Height          =   345
            Index           =   3
            Left            =   0
            TabIndex        =   3
            Top             =   120
            Width           =   1095
            _Version        =   851968
            _ExtentX        =   1931
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "Guardar"
            UseVisualStyle  =   -1  'True
            Picture         =   "frmPresupuesto2.frx":0000
         End
      End
      Begin XtremeSuiteControls.PushButton PusBuscarCliente 
         Height          =   315
         Left            =   1980
         TabIndex        =   5
         Top             =   1560
         Width           =   765
         _Version        =   851968
         _ExtentX        =   1349
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "<F2>"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmPresupuesto2.frx":03CD
      End
      Begin XtremeSuiteControls.FlatEdit vimporte 
         Height          =   315
         Left            =   1920
         TabIndex        =   7
         Top             =   2010
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   315
         Left            =   1950
         TabIndex        =   9
         Top             =   1140
         Width           =   825
         _Version        =   851968
         _ExtentX        =   1455
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "<F1>"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmPresupuesto2.frx":0967
      End
      Begin XtremeSuiteControls.FlatEdit vperiodo 
         Height          =   285
         Left            =   3000
         TabIndex        =   10
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
         Left            =   2940
         TabIndex        =   11
         Top             =   1560
         Width           =   5115
         _Version        =   851968
         _ExtentX        =   9022
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vcodigoCta 
         Height          =   315
         Left            =   8280
         TabIndex        =   12
         Top             =   1560
         Width           =   3315
         _Version        =   851968
         _ExtentX        =   5847
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   0
         Left            =   -69910
         TabIndex        =   16
         Top             =   930
         Visible         =   0   'False
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Guardar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmPresupuesto2.frx":0F01
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   1
         Left            =   -68800
         TabIndex        =   17
         Top             =   930
         Visible         =   0   'False
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Borrar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmPresupuesto2.frx":12CE
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   2
         Left            =   -67690
         TabIndex        =   18
         Top             =   930
         Visible         =   0   'False
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Modificar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmPresupuesto2.frx":1868
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   285
         Left            =   8310
         TabIndex        =   19
         Top             =   1170
         Width           =   3315
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
         Left            =   210
         TabIndex        =   8
         Top             =   1140
         Width           =   1575
         _Version        =   851968
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Período:"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblImporte 
         Height          =   225
         Left            =   210
         TabIndex        =   6
         Top             =   2040
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
         TabIndex        =   4
         Top             =   1590
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
Attribute VB_Name = "frmPresupuesto2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
init
End Sub

Private Sub init()

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2 - 1000

Me.vcodigoCta.Text = 0
Me.vnombrecta.Tag = ""
Me.vnombrecta.Text = ""

'CentrarFormulario (Me)
End Sub

Private Sub PbAcciones_Click(Index As Integer)
On Error Resume Next
Dim vsql, vcampos, vvalores As String


vcampos = "idcuentas,periodo,importe"
vvalores = Str(Val(Me.vnombrecta.Tag)) + ",'" + Me.vperiodo + "'," + Me.vimporte


vsql = "insert into presupuesto (" + vcampos + ") values (" + vvalores + ")"
Call EjecutarScript(vsql, pathDBMySQL)

If Err Then Exit Sub
End Sub

Private Sub PusBuscarCliente_Click()
Call fbuscarGrilla("cuentas", "Cuenta", "idCuentas", Me.vnombrecta.Name, Me)  ' ema:
End Sub

Private Sub PushButton1_Click()
Call fbuscarGrilla("balances", "codigo", "idBalances", Me.vperiodo.Name, Me)     ' ema:
End Sub

Private Sub vbucar_Change()
On Error Resume Next
Dim vsql, vwhere, vCampo  As String

vCampo = "cuentas.CodigoCuenta, cuentas.Cuenta, cuentas.Imputable,presupuesto.periodo, presupuesto.importe"

Dim rspresupuesto As New ADODB.Recordset

vwhere = " where cuentas.CodigoCuenta  like '%" + vbucar + "%' or cuentas.Cuenta like '%" + vbuscar + "%'"

vsql = "SELECT " + vCampo + " From `cuentas` INNER JOIN `presupuesto` ON (`cuentas`.`idCuentas` = `presupuesto`.`idcuentas`) " + vwhere
    
    
    With rspresupuesto
        Call .Open(vsql, ConnDDBB, adOpenStatic, adLockPessimistic)
        Set grilla.DataSource = .DataSource
    End With
        
End Sub

Private Sub vcodigoCta_Change()
'Me.vimporte.SetFocus
End Sub

Private Sub vnombrecta_Change()
Dim vsql As String

vsql = "select codigocuenta from cuentas t where t.idcuentas = " + Me.vnombrecta.Tag
Me.vcodigoCta = traerDatos2(vsql, "codigocuenta", pathDBMySQL)
    
    
End Sub
