VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmCargarCodigoBarra 
   Caption         =   "Descarga de servicios por código de barra"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13230
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   13230
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.PushButton PusCancelarPago 
      Height          =   420
      Left            =   0
      TabIndex        =   25
      Top             =   8325
      Width           =   2040
      _Version        =   851968
      _ExtentX        =   3598
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "Anular el pago al recibo seleccionado "
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PusPrueba 
      Height          =   375
      Left            =   3555
      TabIndex        =   24
      Top             =   8370
      Visible         =   0   'False
      Width           =   870
      _Version        =   851968
      _ExtentX        =   1535
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Prueba"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit vt 
      Height          =   285
      Left            =   7650
      TabIndex        =   22
      Top             =   495
      Width           =   3570
      _Version        =   851968
      _ExtentX        =   6297
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "Listados de recibos imputados "
   End
   Begin XtremeSuiteControls.FlatEdit vtotalManual 
      Height          =   285
      Left            =   7965
      TabIndex        =   17
      Top             =   8415
      Width           =   1680
      _Version        =   851968
      _ExtentX        =   2963
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
   End
   Begin MSComCtl2.DTPicker fpago 
      Height          =   330
      Left            =   3105
      TabIndex        =   12
      Top             =   225
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      Format          =   76218369
      CurrentDate     =   43126
   End
   Begin XtremeSuiteControls.GroupBox GroSeleccioneEl 
      Height          =   780
      Left            =   90
      TabIndex        =   8
      Top             =   45
      Width           =   2850
      _Version        =   851968
      _ExtentX        =   5027
      _ExtentY        =   1376
      _StockProps     =   79
      Caption         =   "Seleccione el tipo de servicio:"
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
      Begin XtremeSuiteControls.RadioButton RadComercio 
         Height          =   240
         Left            =   90
         TabIndex        =   9
         Top             =   360
         Width           =   960
         _Version        =   851968
         _ExtentX        =   1693
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Comercio"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RadRural 
         Height          =   240
         Left            =   1125
         TabIndex        =   10
         Top             =   360
         Width           =   645
         _Version        =   851968
         _ExtentX        =   1138
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Rural"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RadUrbana 
         Height          =   240
         Left            =   1935
         TabIndex        =   11
         Top             =   360
         Width           =   780
         _Version        =   851968
         _ExtentX        =   1376
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Urbana"
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   240
      Left            =   0
      TabIndex        =   2
      Top             =   8055
      Width           =   990
      _Version        =   851968
      _ExtentX        =   1746
      _ExtentY        =   423
      _StockProps     =   79
      Caption         =   "Borrar Linea"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit vCodigoBarra 
      Height          =   300
      Left            =   7200
      TabIndex        =   1
      Top             =   45
      Width           =   3630
      _Version        =   851968
      _ExtentX        =   6403
      _ExtentY        =   529
      _StockProps     =   77
      BackColor       =   -2147483643
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla 
      Height          =   7005
      Left            =   45
      TabIndex        =   0
      Top             =   855
      Width           =   13140
      _ExtentX        =   23178
      _ExtentY        =   12356
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      WordWrap        =   -1  'True
      GridLinesUnpopulated=   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   330
      Left            =   10935
      TabIndex        =   3
      Top             =   8415
      Width           =   2205
      _Version        =   851968
      _ExtentX        =   3889
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Aceptar los datos cargados"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PushButton3 
      Height          =   240
      Left            =   1035
      TabIndex        =   4
      Top             =   8055
      Width           =   1035
      _Version        =   851968
      _ExtentX        =   1826
      _ExtentY        =   423
      _StockProps     =   79
      Caption         =   "Limpiar Grilla"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PusValidar 
      Height          =   375
      Left            =   2610
      TabIndex        =   7
      Top             =   8370
      Visible         =   0   'False
      Width           =   900
      _Version        =   851968
      _ExtentX        =   1587
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Validar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit vnro_recibo 
      Height          =   300
      Left            =   11745
      TabIndex        =   15
      Top             =   45
      Width           =   1335
      _Version        =   851968
      _ExtentX        =   2355
      _ExtentY        =   529
      _StockProps     =   77
      BackColor       =   -2147483643
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton PusImprimirCerificado 
      Height          =   330
      Left            =   10980
      TabIndex        =   21
      Top             =   7920
      Width           =   2160
      _Version        =   851968
      _ExtentX        =   3810
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Imprimir Cerificado:"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PusRecalcular 
      Height          =   330
      Left            =   11745
      TabIndex        =   26
      Top             =   450
      Width           =   1440
      _Version        =   851968
      _ExtentX        =   2540
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Recalcular"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblTìtuloDel 
      Height          =   285
      Left            =   5445
      TabIndex        =   23
      Top             =   495
      Width           =   2175
      _Version        =   851968
      _ExtentX        =   3836
      _ExtentY        =   503
      _StockProps     =   79
      Caption         =   "Tìtulo del listado para anexar:"
   End
   Begin XtremeSuiteControls.Label lblPrimerVencimiento 
      Height          =   240
      Left            =   4185
      TabIndex        =   20
      Top             =   585
      Width           =   1140
      _Version        =   851968
      _ExtentX        =   2011
      _ExtentY        =   423
      _StockProps     =   79
      Caption         =   "Primer Vnto."
      ForeColor       =   16777215
      BackColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
   End
   Begin XtremeSuiteControls.Label lblVencimiento2 
      Height          =   240
      Left            =   3105
      TabIndex        =   19
      Top             =   585
      Width           =   1005
      _Version        =   851968
      _ExtentX        =   1773
      _ExtentY        =   423
      _StockProps     =   79
      Caption         =   "Seg. Vnto."
      ForeColor       =   16777215
      BackColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
   End
   Begin XtremeSuiteControls.Label lblIngImporte 
      Height          =   330
      Left            =   6165
      TabIndex        =   18
      Top             =   8415
      Width           =   1725
      _Version        =   851968
      _ExtentX        =   3043
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Ing. importe manual:"
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   240
      Left            =   10890
      TabIndex        =   16
      Top             =   90
      Width           =   870
      _Version        =   851968
      _ExtentX        =   1535
      _ExtentY        =   423
      _StockProps     =   79
      Caption         =   "Nro Recibo:"
   End
   Begin XtremeSuiteControls.Label lblIngresarCódigo 
      Height          =   285
      Left            =   5400
      TabIndex        =   14
      Top             =   45
      Width           =   1770
      _Version        =   851968
      _ExtentX        =   3122
      _ExtentY        =   503
      _StockProps     =   79
      Caption         =   "Ingresar código de barra:"
   End
   Begin XtremeSuiteControls.Label lblFechaDe 
      Height          =   195
      Left            =   3105
      TabIndex        =   13
      Top             =   0
      Width           =   1455
      _Version        =   851968
      _ExtentX        =   2566
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Fecha de Pago:"
   End
   Begin VB.Label lblTotalImportes 
      Caption         =   "Total importes recibos: _________"
      Height          =   225
      Left            =   6300
      TabIndex        =   6
      Top             =   8055
      Width           =   3285
   End
   Begin VB.Label lblCantidadDe 
      Caption         =   "Cantidad de recibos cargados: ____________"
      Height          =   225
      Left            =   2565
      TabIndex        =   5
      Top             =   8055
      Width           =   3285
   End
End
Attribute VB_Name = "frmCargarCodigoBarra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vtabla, vwhere As String

Dim arr(100) As String
Dim j As Integer
Dim total As Double
Dim cantidad2 As Integer

Dim pagarRecibo(100) As String

Dim arecibos() As String


Dim vcolor2 As Variant


Private Sub MSHFlexGrid1_Click()

End Sub

Private Sub Form_Load()
Me.grilla.Clear

init
End Sub


Private Sub init()
On Error Resume Next
Dim i As Integer


Set ConnComunaDB2 = New ADODB.Connection
ConnComunaDB2.ConnectionString = LeerXml("ComunaCnn")


'Me.vCodigoBarra.SetFocus
j = 0

Me.grilla.Clear
Me.grilla.Rows = 1
        

Me.fpago.Value = Date

grilla.ColWidth(0) = 3500
grilla.ColWidth(1) = 3500
grilla.ColWidth(2) = 1500
grilla.ColWidth(3) = 1500

Me.vCodigoBarra.SetFocus

grilla.Rows = 1

If Err Then Exit Sub
End Sub

Public Sub setrecibos(recibos() As String)
    arecibos = recibos()
    
    Call pasar_arecibosAgrilla
    
    Call CalcularTotal
    
End Sub


Sub CalcularTotal()
Dim i As Integer
Dim vtotal As Double

vtotal = 0

For i = 0 To grilla.Rows - 1
    vtotal = vtotal + Val(grilla.TextMatrix(i, 4))
Next

Me.lblCantidadDe.Caption = "> Cantidad de recibos: " + Str(grilla.Rows - 1)
vtotalManual.Text = vtotal

End Sub

Sub pasar_arecibosAgrilla()

Dim i As Integer

With grilla

    .Clear
    .Rows = 0
    For i = 0 To UBound(arecibos)
        If Not arecibos(i) = "" Then
            .AddItem arecibos(i)
        End If
        'i = i + 1
        
    Next
    
End With

Me.Show
End Sub


Private Sub grilla_DblClick()
Dim r As Integer

Dim vnro, vc, vsql   As String

r = grilla.Row
vnro = grilla.TextMatrix(r, 5)

vsql = fsql2 + " and nro_recibo = '" + vnro + "' order by fecha_emision desc limit 1"

Me.Caption = vsql + "  ///  "

'MsgBox ("2 " + pathcomunadb)

vc = traerDatos2(vsql, "cod_barra", ConnComunaDB2)

grilla.TextMatrix(r, 0) = vc

'Call vnro_recibo_KeyPress(13)

End Sub

Private Sub PusCancelarPago_Click()
Dim vnro, vsql, vcancelar  As String
Dim r As Integer

r = grilla.Row
vnro = grilla.TextMatrix(r, 5)
 
vcancelar = InputBox("Si quere cancelar el pago realizado al recibo nro: " + vnro + Chr(13) + _
"Debe escribir la palabra: anular")


If vcancelar = "anular" Then

    vsql = "update recibo_resumen set id_estados = 'IM', importe_pagado = 0 , pagado_en = '', fecha_pago= Null where nro_recibo = '" + vnro + "'"
    
    Call EjecutarScript(vsql, ConnComunaDB2)
    
    MsgBox "El pago del recibo fue cancelado"
    
    grilla.RemoveItem (r)
    
Else

    MsgBox "No hubo cambios registrados" + Chr(13) + "El recibo quedará sin cambios"
                                        
End If

End Sub

Private Sub PushButton1_Click()
On Error Resume Next

    grilla.RemoveItem (grilla.Row)

If Err Then Exit Sub
End Sub

Private Sub PushButton2_Click()
Dim vmensaje As String

vmensaje = "Acepta la imputación de los datos de la grilla"
'Call validar


If InputBox("Para ejecutar el pago debe escribir la palabra: " + Chr(13) + "pagar") = "pagar" Then
                If MsgBox(vmensaje, vbYesNo) = vbYes Then
                    
                    Call ejecutar_pago
                    
                End If
                
                If MsgBox("Quiere verificar el nuevo saldo", vbYesNo) = vbYes Then
                    Call frmDeudasServicios.PusGenerarListado_Click
                    
                End If
                
                
                Unload Me
End If

End Sub



Private Sub acpectoar()

End Sub

Private Sub PushButton3_Click()

If MsgBox("Está seguro ?", vbYesNo) = vbYes Then
        Me.grilla.Clear
        Me.grilla.Rows = 1
End If

End Sub


Function fengrilla(ByVal grilla As MSHFlexGrid, ByVal v As String, ByVal Col As Integer) As Boolean

Dim i As Integer

i = 0

fengrilla = False

For i = 0 To grilla.Rows - 1
   
    If grilla.TextMatrix(i, Col) = v Then fengrilla = True
    
     
Next

End Function


Function fenarray(ByVal c) As Boolean


Dim f As Boolean

f = False

Dim i As Integer

For i = 1 To UBound(arr())
    If arr(i) = c Then
        f = True
        fenarray = f
    End If
    i = i + 1
Next

j = j + 1
arr(j) = c

fenarray = f

End Function


Function calTotal() As Double

Dim i As Integer


total = 0

For i = 0 To grilla.Rows - 1

    total = total + Val(grilla.TextMatrix(i, 4))

Next

calTotal = total

End Function



Function calCantidad() As Double

Dim i As Integer


cantidad2 = 0

For i = 1 To grilla.Rows - 1

    cantidad2 = cantidad2 + 1

Next

calCantidad = cantidad2

End Function


Private Sub ejecutar_pago()
On Error Resume Next
Dim i As Integer

For i = 0 To grilla.Rows - 1
      Call pagar(grilla.TextMatrix(i, 0), grilla.TextMatrix(i, 5))
Next

MsgBox "Los pagos fueron imputados al mòdulo de servicio" + Chr(13) + "Continuamos al siguiente paso"


With frmIngresosEgresos

    Call .initIngreso
    .txtAlta(12).Text = Me.vtotalManual.Text
    .Show

End With


If Err Then Exit Sub

End Sub


Private Sub pagar(vcode As String, vnro_recibo As String)
On Error Resume Next

Dim vsql2, vsql, vcampo, vValor As String

If vcode = "" And vnro_recibo = "" Then Exit Sub


If Not vcode = "" Then
    vsql = "update recibo_resumen set fecha_pago = now(), pagado_en = 'C', id_estados = 'PA' where cod_barra = '" + vcode + "'"
Else
    vsql = "update recibo_resumen set fecha_pago = now(), pagado_en = 'C', id_estados = 'PA' where nro_recibo = '" + vnro_recibo + "'"
End If


Call EjecutarScript(vsql, ConnComunaDB2)

vcampo = "(proceso,hora,formulario,fecha,comentario)"

vValor = "'pagar servicio',"
vValor = vValor + "0" + ","
vValor = vValor + "'frmCargarCodigoBarra'" + ","
vValor = vValor + "'" + Replace((Str(Year(Date)) + "-" + Str(Month(Date)) + "-" + Str(Day(Date))), " ", "") + "',"
vValor = vValor + "'" + vcode + "'"

vsql2 = "insert into log " + vcampo + " values " + "(" + vValor + ")"

Call EjecutarScript(vsql2, pathDBMySQL)

If Err < 0 Then
    MsgBox "Ocurrió un error con el código: " + vcode
    Exit Sub
End If
End Sub


Private Sub actualizarTabla()

If Me.RadComercio.Value Then vtabla = "recibo_comercio_resumen"
If Me.RadRural Then vwhere = "id_tipos_zonas= 2"
If Me.RadUrbana Then vwhere = "id_tipos_zonas= 1"

vwhere = " 1=1 "

End Sub



Private Sub gettrecibo(ByVal v1, ByRef v2, ByRef v3, ByRef v4, ByRef v5)

Dim vsql, nro_recibo, periodo_anomes, Nombre As String
Dim importe_total As Double
Dim id_contribuyentes As Long

Dim fecha_pago As String


actualizarTabla


If fengrilla(Me.grilla, v1, 0) = True Then
    MsgBox "Este recibo está cargado en la grilla", vbInformation
    Exit Sub
End If



vsql = fsql2 + " and cod_barra = '" + v1 + "'"


' Me.Caption = vsql + "  ///  " + ConnComunaDB2

nro_recibo = traerDatos2(vsql, "nro_recibo", ConnComunaDB2)

If nro_recibo = "" Then
    MsgBox "Este recibo no fue encontrado"
    Exit Sub
End If

fecha_pago = traerDatos2(vsql, "fecha_pago", ConnComunaDB2)


periodo_anomes = traerDatos2(vsql, "periodo_anomes", ConnComunaDB2)

importe_total = traerDatos2(vsql, "importe_total", ConnComunaDB2)

id_contribuyentes = Val(traerDatos2(vsql, "id_contribuyentes", ConnComunaDB2))



Dim importe2, importe3 As Double

importe2 = traerDatos2(vsql, "importe_total2", ConnComunaDB2)

importe3 = traerDatos2(vsql, "importe_total3", ConnComunaDB2)



Dim vence1, vence2, vence3 As Date

vence1 = CDate(traerDatos2(vsql, "fecha_vencimiento", ConnComunaDB2))

vence2 = CDate(traerDatos2(vsql, "fecha_vencimiento2", ConnComunaDB2))


vence3 = CDate(traerDatos2(vsql, "fecha_vencimiento3", ConnComunaDB2))



vsql = "select concat(apellido, ', ',nombre) as nombre from  contribuyentes c inner join personas p  on p.id_personas = c.id_personas where  c.id_contribuyentes =" + Str(id_contribuyentes)


Nombre = traerDatos2(vsql, "nombre", ConnComunaDB2)

v1 = id_contribuyentes
v2 = Nombre
v3 = periodo_anomes


If vence3 < Me.fpago.Value Then

    v5 = importe3
    vcolor2 = vbRed
    
Else

    If vence2 < Me.fpago.Value Then
    
        v5 = importe2
        vcolor2 = vbBlue
        
    Else
    
        v5 = importe_total
        vcolor2 = vbWhite
       

    End If

End If


v5 = importe_total


If Not fecha_pago = "" Then
    MsgBox " Este recibo ya se encuentra pago " + Chr(13) + " Verifique este dato con el encargado de Servicios "
    v1 = ""
End If


End Sub

Private Sub PusImprimirCerificado_Click()
    Call grillaToExcel(Me.grilla, vt)
End Sub

Private Sub PusPrueba_Click()
Dim vsql, v1, v2, v As String

Set ConnComunaDB2 = New ADODB.Connection

ConnComunaDB2.ConnectionString = LeerXml("ComunaCnn")

vsql = "select * from recibo_resumen limit 1"

v1 = traerDatos2(vsql, "nro_recibo", ConnComunaDB2)


MsgBox (v1 + Chr(13) + ConnComunaDB2)


v = "driver={MySQL ODBC 3.51 Driver};server=pc-servicios;port=3306;uid=root;pwd=root.2009;database=comunadb;OPTION=8"

v2 = traerDatos2(vsql, "nro_recibo", ConnComunaDB2)

MsgBox (v2 + Chr(13) + ConnComunaDB2)

End Sub

Private Sub PusRecalcular_Click()
Dim r As Integer

Dim vnro, vc, vsql   As String

r = grilla.Row

Dim i As Integer


For i = 0 To grilla.Rows - 1

vnro = grilla.TextMatrix(i, 5)

vsql = fsql2 + "  nro_recibo = '" + vnro + "' order by fecha_emision desc limit 1"

Me.Caption = vsql + "  ///  "

'MsgBox ("2 " + pathcomunadb)

vc = traerDatos2(vsql, "cod_barra", ConnComunaDB2)

grilla.TextMatrix(i, 0) = vc
grilla.TextMatrix(i, 6) = traerDatos2(vsql, "periodo_anomes", ConnComunaDB2)
grilla.TextMatrix(i, 7) = traerDatos2(vsql, "id_contribuyentes", ConnComunaDB2)

Next

'Call vnro_recibo_KeyPress(13)
End Sub

Private Sub vCodigoBarra_KeyPress(KeyAscii As Integer)
On Error Resume Next

Dim vCodigoBarra, vcodigo, vnombre, vperiodo, v As String

Dim vsql As String

Dim vimporte As Double

vcodigo = ""

Call actualizarTabla

If KeyAscii = 13 Then

    Call gettrecibo(Trim(Me.vCodigoBarra.Text), vcodigo, vnombre, vperiodo, vimporte)

    'If Not v = "" Then
        v = Me.vCodigoBarra.Text + vbTab + vcodigo + vbTab + vnombre + vbTab + vperiodo + vbTab + Str(vimporte) + vbTab + Me.vnro_recibo.Text
    'End If
    
     If Not vcodigo = "" Then
        grilla.AddItem v
        grilla.Row = grilla.Rows - 1
        grilla.Col = 1
        grilla.CellBackColor = vcolor2
       
        
    End If
     

    Me.vCodigoBarra.Text = ""
    
    Dim vtotaltemp As Double
    
    vtotaltemp = Me.calTotal()
    
    Me.lblTotalImportes.Caption = "Total importe: " + Str(vtotaltemp)
    Me.lblCantidadDe.Caption = "Total recibos: " + Str(Me.calCantidad())
    
    Me.vtotalManual.Text = vtotaltemp
    
End If


'j = j + 1

'arr(j) = vCodigoBarra

vCodigoBarra.SetFocus



If Err Then Exit Sub
End Sub

Function fsql2() As String
    fsql2 = "select * from recibo_resumen where " + vwhere
End Function



Function pathcomunadb() As String
    pathcomunadb = ConnComunaDB.ConnectionString
End Function


Private Sub vnro_recibo_KeyPress(KeyAscii As Integer)
Dim vsql, vc As String


If Not KeyAscii = 13 Then Exit Sub

Call actualizarTabla

vsql = fsql2 + " and  nro_recibo like '%" + vnro_recibo + "' order by fecha_emision desc limit 1"

Me.Caption = vsql + "  ///  "

'MsgBox ("2 " + pathcomunadb)

vc = traerDatos2(vsql, "cod_barra", ConnComunaDB2)


If vc = "" Then
    MsgBox "No encontrado"
   Exit Sub
End If

Me.vCodigoBarra.Text = vc

Call vCodigoBarra_KeyPress(13)

vnro_recibo.Text = ""

End Sub
