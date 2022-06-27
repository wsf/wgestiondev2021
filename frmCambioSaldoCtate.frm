VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmCambioSaldoCtate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos del movimiento de ajuste de Saldo"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.RadioButton rfijo 
      Height          =   345
      Left            =   1530
      TabIndex        =   8
      Top             =   1020
      Width           =   1185
      _Version        =   851968
      _ExtentX        =   2090
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Fijar saldo"
      UseVisualStyle  =   -1  'True
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   465
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7575
      _Version        =   851968
      _ExtentX        =   13361
      _ExtentY        =   820
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   30
         Left            =   0
         TabIndex        =   11
         Top             =   390
         Width           =   7545
         _Version        =   851968
         _ExtentX        =   13309
         _ExtentY        =   53
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   0
         Left            =   30
         TabIndex        =   7
         Top             =   0
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Aplicar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCambioSaldoCtate.frx":0000
      End
   End
   Begin VB.TextBox vsaldo 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   1500
      TabIndex        =   5
      Top             =   1440
      Width           =   2595
   End
   Begin VB.TextBox vcomentario 
      Height          =   315
      Left            =   1470
      TabIndex        =   4
      Top             =   1860
      Width           =   6105
   End
   Begin MSComCtl2.DTPicker vfecha 
      Height          =   315
      Left            =   1530
      TabIndex        =   1
      Top             =   630
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   556
      _Version        =   393216
      Format          =   146997249
      CurrentDate     =   40694
   End
   Begin XtremeSuiteControls.RadioButton rdebito 
      Height          =   345
      Left            =   2760
      TabIndex        =   9
      Top             =   1020
      Width           =   1995
      _Version        =   851968
      _ExtentX        =   3519
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Ingresar como débito"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.RadioButton rcredito 
      Height          =   345
      Left            =   4800
      TabIndex        =   10
      Top             =   1020
      Width           =   1995
      _Version        =   851968
      _ExtentX        =   3519
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Ingresar como Crédito"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FormExtender ForDatosDel 
      Left            =   7050
      Top             =   1290
      _Version        =   851968
      _ExtentX        =   423
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin XtremeSuiteControls.Label lblComentario 
      Height          =   225
      Left            =   60
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
      _Version        =   851968
      _ExtentX        =   1931
      _ExtentY        =   397
      _StockProps     =   79
      Caption         =   "> Comentario:"
      Alignment       =   1
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   225
      Left            =   60
      TabIndex        =   2
      Top             =   1500
      Width           =   1095
      _Version        =   851968
      _ExtentX        =   1931
      _ExtentY        =   397
      _StockProps     =   79
      Caption         =   "> Importe:"
      Alignment       =   1
   End
   Begin XtremeSuiteControls.Label lblFecha 
      Height          =   225
      Left            =   420
      TabIndex        =   0
      Top             =   660
      Width           =   705
      _Version        =   851968
      _ExtentX        =   1244
      _ExtentY        =   397
      _StockProps     =   79
      Caption         =   "> Fecha:"
      Alignment       =   1
   End
End
Attribute VB_Name = "frmCambioSaldoCtate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vcodigo, vnombre, vctacteCP As String
Public vsaldoactual  As Double


Private Sub Form_Load()
vfecha = Date
'vsaldo.Text = 0
vcomentario = "Ajuste de Saldo (Nuevo Sistema)"

End Sub

Private Sub PbAcciones_Click(Index As Integer)
On Error Resume Next
Dim i As Integer
Dim sqlscript, vValor, vcampos As String
Dim vdebito, vcredito, vauxi As Double
Dim vnrointerno As Long


If Index = 1 Then

Unload Me
Exit Sub
End If
i = 1
If vsaldoactual < 0 Then
    i = -1
Else
    i = 1
End If


If (CDbl(vsaldoactual) > CDbl(Val(vsaldo.Text))) Then
    vcredito = Abs(vsaldoactual - Val(vsaldo.Text))
Else
    vdebito = Abs(vsaldoactual - Val(vsaldo.Text))
End If


'If vctacteCP = "pcuentascorrientes" Then

'vauxi = vdebito
'vdebito = vcredito
'vcredito = vauxi

'End If


If Me.rcredito.Value Then
    vcredito = vsaldo
    vdebito = 0
End If '

If Me.rdebito.Value Then
    vdebito = vsaldo
    vcredito = 0
End If


'If rfijo.Value Then'
'
'    If vsaldo > 0 Then
'        vdebito = 0
'        vcredito = vsaldoactual
'    Else
'            vdebito = vsaldoactual
'            vcredito = 0
'    End If

'End If

vnrointerno = UltimoNroInterno2 + 1

vcampos = "TipoMovimiento,fecha,codigo,nombre,debito,credito,comentario,nrointerno"

vValor = "('AJ','" + strfechaMySQL(vfecha) + "','" + vcodigo + "','" + Left(vnombre, 50) + "'," + Str(vdebito) + "," + Str(vcredito) + ",'" + vcomentario.Text + "'," + Str(vnrointerno) + ")"

sqlscript = "insert into " + vctacteCP + " (" + vcampos + ") value " + vValor

Call EjecutarScript(sqlscript)

Call frmCtaCteC.cmdFiltroMovimientos_Click

Unload Me
If Err < 0 Then
    MsgBox "No se pudo modificar el saldo adecuadamente", vbCritical, "Error"
End If
End Sub

Private Sub vcomentario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then PbAcciones(0).SetFocus
End Sub

Private Sub vfecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then vsaldo.SetFocus
End Sub

Private Sub vsaldo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then vcomentario.SetFocus
End Sub
