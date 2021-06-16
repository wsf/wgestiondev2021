VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{9746E3DA-06E1-4D26-9CE4-D9F6411A9C70}#1.0#0"; "SMGA_OcxTxt2009.ocx"
Begin VB.Form frmVolquetesAdmin 
   Caption         =   "Cambios de estados de documento:"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.GroupBox GroupBox4 
      Height          =   2745
      Left            =   30
      TabIndex        =   14
      Top             =   570
      Width           =   9795
      _Version        =   851968
      _ExtentX        =   17277
      _ExtentY        =   4842
      _StockProps     =   79
      Caption         =   "Documentos seleccionados:"
      Appearance      =   2
      Begin XtremeSuiteControls.ProgressBar barra 
         Height          =   225
         Left            =   60
         TabIndex        =   31
         Top             =   2400
         Width           =   9675
         _Version        =   851968
         _ExtentX        =   17066
         _ExtentY        =   397
         _StockProps     =   93
         Text            =   "Barra"
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.ListBox vlista 
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFF00&
         Height          =   2010
         Left            =   60
         TabIndex        =   15
         Top             =   210
         Width           =   9705
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   1935
      Left            =   30
      TabIndex        =   11
      Top             =   3450
      Width           =   9765
      _Version        =   851968
      _ExtentX        =   17224
      _ExtentY        =   3413
      _StockProps     =   79
      Caption         =   "Documento: "
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
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.ComboBox cmb_estadopedidoCambio 
         Height          =   315
         Left            =   7200
         TabIndex        =   0
         Top             =   390
         Width           =   2535
         _Version        =   851968
         _ExtentX        =   4471
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "Cmb_estadopedido cambio"
      End
      Begin XtremeSuiteControls.FlatEdit txt_montoParcial 
         Height          =   315
         Left            =   7170
         TabIndex        =   1
         Top             =   810
         Width           =   2565
         _Version        =   851968
         _ExtentX        =   4524
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit vsaldo 
         Height          =   315
         Left            =   7170
         TabIndex        =   2
         Top             =   1170
         Width           =   2535
         _Version        =   851968
         _ExtentX        =   4471
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin Aplisoft_CajasDeTexto.TxF vfechapago2 
         Height          =   285
         Left            =   2430
         TabIndex        =   28
         Top             =   1590
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Aplisoft_CajasDeTexto.TxF vfechapago 
         Height          =   285
         Left            =   7140
         TabIndex        =   30
         Top             =   1620
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label Label11 
         Height          =   315
         Left            =   60
         TabIndex        =   29
         Top             =   1560
         Width           =   2235
         _Version        =   851968
         _ExtentX        =   3942
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Fecha estimada de pago:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha estimada de pago:"
         Height          =   255
         Left            =   4530
         TabIndex        =   27
         Top             =   1650
         Width           =   2505
      End
      Begin XtremeSuiteControls.Label vtotal 
         Height          =   315
         Left            =   2460
         TabIndex        =   26
         Top             =   990
         Width           =   1305
         _Version        =   851968
         _ExtentX        =   2302
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "0.00"
         ForeColor       =   255
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   315
         Left            =   210
         TabIndex        =   25
         Top             =   960
         Width           =   2145
         _Version        =   851968
         _ExtentX        =   3784
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Total del documento:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Falta pagar:"
         Height          =   255
         Left            =   4590
         TabIndex        =   24
         Top             =   1170
         Width           =   2505
      End
      Begin XtremeSuiteControls.Label vsumaDePago 
         Height          =   315
         Left            =   2340
         TabIndex        =   19
         Top             =   660
         Width           =   1425
         _Version        =   851968
         _ExtentX        =   2514
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "0.00"
         ForeColor       =   255
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   315
         Left            =   90
         TabIndex        =   18
         Top             =   630
         Width           =   2385
         _Version        =   851968
         _ExtentX        =   4207
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Suma de pagos anteriores:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label vestadoAnterior 
         Height          =   315
         Left            =   2340
         TabIndex        =   17
         Top             =   300
         Width           =   1425
         _Version        =   851968
         _ExtentX        =   2514
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "xx"
         ForeColor       =   255
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label vestado 
         Height          =   315
         Left            =   930
         TabIndex        =   16
         Top             =   300
         Width           =   1425
         _Version        =   851968
         _ExtentX        =   2514
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Estado anterior:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Cambiar el estado del  documento:"
         Height          =   255
         Left            =   4560
         TabIndex        =   13
         Top             =   450
         Width           =   2535
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Cambiar Importe pago:"
         Height          =   255
         Left            =   4590
         TabIndex        =   12
         Top             =   810
         Width           =   2505
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   975
      Left            =   30
      TabIndex        =   8
      Top             =   5640
      Width           =   9765
      _Version        =   851968
      _ExtentX        =   17224
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   "Volquetes:"
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
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.ComboBox cmbCambiaestadopedido 
         Height          =   315
         Left            =   7140
         TabIndex        =   3
         Top             =   210
         Width           =   2595
         _Version        =   851968
         _ExtentX        =   4577
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "No Retirado"
      End
      Begin XtremeSuiteControls.FlatEdit vVolquetesDevueltos 
         Height          =   315
         Left            =   7140
         TabIndex        =   4
         Top             =   600
         Width           =   2595
         _Version        =   851968
         _ExtentX        =   4577
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label vsumaVolquetesDevuelto 
         Height          =   315
         Left            =   3300
         TabIndex        =   23
         Top             =   600
         Width           =   1425
         _Version        =   851968
         _ExtentX        =   2514
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "0.00"
         ForeColor       =   255
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label9 
         Height          =   315
         Left            =   150
         TabIndex        =   22
         Top             =   570
         Width           =   2805
         _Version        =   851968
         _ExtentX        =   4948
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Suma de volquetes devueltos:"
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label vestadoVolqueteAnterior 
         Height          =   315
         Left            =   3270
         TabIndex        =   21
         Top             =   210
         Width           =   1425
         _Version        =   851968
         _ExtentX        =   2514
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "xx"
         ForeColor       =   255
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label7 
         Height          =   315
         Left            =   1530
         TabIndex        =   20
         Top             =   240
         Width           =   1425
         _Version        =   851968
         _ExtentX        =   2514
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Estado anterior:"
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
         Alignment       =   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Cambio de estado del volquete:"
         Height          =   195
         Left            =   4710
         TabIndex        =   10
         Top             =   270
         Width           =   2385
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Definir Volquetes devueltos:"
         Height          =   255
         Left            =   4680
         TabIndex        =   9
         Top             =   660
         Width           =   2415
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   585
      Left            =   0
      TabIndex        =   5
      Top             =   -60
      Width           =   9795
      _Version        =   851968
      _ExtentX        =   17277
      _ExtentY        =   1032
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   375
         Left            =   60
         TabIndex        =   6
         Top             =   150
         Width           =   2025
         _Version        =   851968
         _ExtentX        =   3572
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Ejecutar cambios [F2]"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Picture         =   "frm_VolquetesAdmin.frx":0000
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   375
         Left            =   2070
         TabIndex        =   7
         Top             =   150
         Width           =   2025
         _Version        =   851968
         _ExtentX        =   3572
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Imprimir Comprobante"
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Picture         =   "frm_VolquetesAdmin.frx":059A
      End
   End
End
Attribute VB_Name = "frmVolquetesAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vtablaFactura, vIdFactura, vtablaCtaCte As String



Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
 
    If KeyCode = 13 Then
        Call Form_KeyUp(vbTab, 1)
    End If

    If KeyCode = vbKeyF2 Then
        Call PushButton3_Click
    End If



End Sub

Private Sub Form_Load()
finit
End Sub


Private Sub actualizarDocMarcados()
Dim i As Integer
Dim vid As Long
Dim vsql
      
        If Not (MsgBox("Está seguro de cambiar el tipo de estado de toda las facturras marcadas ?" _
            + "Cantidad: " + Str(frmBuscarFactura.vdocmarcados), vbYesNo)) = vbYes Then
            Exit Sub
        End If
        
' ----------------------

With frmBuscarFactura.KlexDocumentos
barra.Max = .Rows - 1
barra.Value = 0

    For i = 1 To .Rows - 1
        barra.Value = i
        If Not Trim(.TextMatrix(i, 0)) = "" Then
            vid = .TextMatrix(i, 1)
            
            If Me.cmb_estadopedidoCambio = "No Admitido" Then
                        .TextMatrix(i, 15) = Me.cmb_estadopedidoCambio
                        .CellBackColor = vbRed + 12
            Else
                       .TextMatrix(i, 15) = Me.cmb_estadopedidoCambio
                       .CellBackColor = &H747474 + 15
            End If
            
            .Row = i
            .Col = 15
    
            vsql = "update  " + Me.vtablaFactura + "  set estadodocumento='" + Trim(Me.cmb_estadopedidoCambio) + "' where " + Me.vIdFactura + "=" + Str(vid)
            Call EjecutarScript(vsql, pathDBMySQL)
            
            vsql = "update  " + Me.vtablaCtaCte + "  set estadoadmicion=" + festadoadmicion(cmb_estadopedidoCambio.Text) + " where nrointerno = " + Str(.TextMatrix(i, 11))
            Call EjecutarScript(vsql, pathDBMySQL)
            
            
        End If
    Next
End With

        
End Sub


Function festadoadmicion(vestado As String) As String
festadoadmicion = "0"

If vestado = "No Admitido" Then
    festadoadmicion = "1"
End If

End Function

Private Sub PushButton3_Click()
Dim id As Long
Dim vFila As Integer
Dim vsql As String
Dim vfdetalle, vid  As String
Dim vremito As Long


If frmBuscarFactura.vdocmarcados > 0 Then
    actualizarDocMarcados
    Exit Sub
End If


vFila = frmBuscarFactura.KlexDocumentos.Row

id = frmBuscarFactura.KlexDocumentos.TextMatrix(vFila, 1)

vsql = "select * from " + vtablaFactura + " Where " + vIdFactura + "=" + Str(id)

vremito = traerDatos2(vsql, "remito", pathDBMySQL)

If vtablaFactura = "pfactura" Then
    vfdetalle = "PFDetalle"
    vid = "idPFDetalle"
End If

If vtablaFactura = "factura" Then
    vfdetalle = "FDetalle"
    vid = "idFDetalle"
End If



If Not Validar Then Exit Sub


If (MsgBox("Está seguro de cambiar el tipo de estado del documento ?", vbYesNo)) = vbYes Then

    If Not cmbCambiaestadopedido.Text = "" Then
        vsql = "update " + vtablaFactura + " set tipopedido='" + Me.cmbCambiaestadopedido.Text + "' where " + vIdFactura + "=" + Str(id)
        Call EjecutarScript(vsql, pathDBMySQL)
    End If
    
    
    If Not Me.cmb_estadopedidoCambio.Text = "" Then
        vsql = "update  " + vtablaFactura + "  set estadodocumento='" + Me.cmb_estadopedidoCambio + "' where " + vIdFactura + "=" + Str(id)
        Call EjecutarScript(vsql, pathDBMySQL)
    End If
    
    
    If Not Me.txt_montoParcial = "" Then
        vsql = "update  " + vtablaFactura + "  set pagoparcial='" + Me.txt_montoParcial + "' where " + vIdFactura + "=" + Str(id)
        Call EjecutarScript(vsql, pathDBMySQL)
    End If
    
    If Not vVolquetesDevueltos.Text = "" Then
            vsql = "update  " + vtablaFactura + "  set cantidadvolquetedevuelto='" + vVolquetesDevueltos + "' where " + vIdFactura + "=" + Str(id)
            Call EjecutarScript(vsql, pathDBMySQL)
    
    End If
    
    If Not Me.vsaldo.Text = "" Then
            vsql = "update  " + vtablaFactura + "  set saldos=" + Me.vsaldo + " where " + vIdFactura + "=" + Str(id)
            Call EjecutarScript(vsql, pathDBMySQL)
    
    End If
    
    
     If Not Me.vfechapago = Me.vfechapago2 Then
            vsql = "update  " + vtablaFactura + "  set fechapago='" + strfechaMySQL(Me.vfechapago) + "' where " + vIdFactura + "=" + Str(id)
            Call EjecutarScript(vsql, pathDBMySQL)
    
    End If
    
    
    If Me.cmb_estadopedidoCambio = "Anulado" Then Call stockAnular(vremito, vfdetalle, vid)
    
    finit
    
End If

Call frmBuscarFactura.cmdFiltrar_Click

Unload Me

End Sub

Function Validar() As Boolean
Dim vmensaje As String

vmensaje = ""

If (Me.vestadoAnterior.Caption = "Anulado") And (Me.cmb_estadopedidoCambio.Text = "Anulado") Then vmensaje = vmensaje + Chr(13) + " - Este documento ya fue anulado"

If Not vmensaje = "" Then
    MsgBox vmensaje
    Validar = False
Else
    Validar = True
End If

End Function


Private Sub finit()

On Error Resume Next
'---------------------------
Me.cmbCambiaestadopedido.Clear
Me.cmbCambiaestadopedido.AddItem "Retirado"
Me.cmbCambiaestadopedido.AddItem "No Retirado"
Me.cmbCambiaestadopedido.AddItem "Recambio"
Me.cmbCambiaestadopedido.AddItem "Todos"
Me.cmbCambiaestadopedido.Text = ""
'---------------------------


'---------
Me.cmb_estadopedidoCambio.Clear
Me.cmb_estadopedidoCambio.AddItem "No Admitido"
Me.cmb_estadopedidoCambio.AddItem "Admitido"
Me.cmb_estadopedidoCambio.AddItem "Adeudado"
Me.cmb_estadopedidoCambio.AddItem "Pagado"
Me.cmb_estadopedidoCambio.AddItem "Quebranto"
Me.cmb_estadopedidoCambio.AddItem "Pendiente"
Me.cmb_estadopedidoCambio.AddItem "Anulado"
Me.cmb_estadopedidoCambio.AddItem "Todos los estados"

Me.cmb_estadopedidoCambio.Text = ""

Dim vFila As Integer



With frmBuscarFactura

vFila = .KlexDocumentos.Row

Me.vestadoAnterior.Caption = .KlexDocumentos.TextMatrix(vFila, 15)
Me.vestadoVolqueteAnterior.Caption = .KlexDocumentos.TextMatrix(vFila, 19)
Me.vsumaVolquetesDevuelto.Caption = .KlexDocumentos.TextMatrix(vFila, 18)
Me.vsumaDePago.Caption = .KlexDocumentos.TextMatrix(vFila, 20)
Me.vsaldo.Text = .KlexDocumentos.TextMatrix(vFila, 22)
Me.vtotal.Caption = .KlexDocumentos.TextMatrix(vFila, 10)

Me.vfechapago2.Value = .KlexDocumentos.TextMatrix(vFila, 23)

End With

If Err Then Exit Sub

End Sub

Private Sub txt_montoParcial_Change()
Me.vsaldo.Text = Val(vtotal.Caption) - Val(Me.txt_montoParcial)
End Sub

