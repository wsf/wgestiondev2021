VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.ShortcutBar.v13.0.0.Demo.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "Copia de KlexGrid.ocx"
Begin VB.Form frmInconsistencias 
   Caption         =   "Inconsistencias de datos..."
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4815
   ScaleWidth      =   11445
   Begin VB.Frame Frame1 
      Caption         =   "Datos inconsistentes: "
      Height          =   4545
      Left            =   6060
      TabIndex        =   1
      Top             =   180
      Width           =   5385
      Begin Grid.KlexGrid grilla 
         Height          =   3315
         Left            =   90
         TabIndex        =   2
         Top             =   180
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   5847
         EnterKeyBehaviour=   0
         BackColorAlternate=   12632256
         GridLinesFixed  =   2
         BackColor       =   14737632
         BackColorFixed  =   -2147483626
         Cols            =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8388608
         GridColorFixed  =   8421504
         MouseIcon       =   "frmInconsistencias.frx":0000
         Rows            =   10
      End
      Begin XtremeShortcutBar.ShortcutCaption vdisplay 
         Height          =   795
         Left            =   150
         TabIndex        =   19
         Top             =   3630
         Width           =   5175
         _Version        =   851968
         _ExtentX        =   9128
         _ExtentY        =   1402
         _StockProps     =   14
         ForeColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Sylfaen"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
         Alignment       =   1
         GradientColorLight=   12648384
         GradientColorDark=   32768
         ForeColor       =   14737632
      End
   End
   Begin VB.Frame FraComprobanciónDe 
      Caption         =   "Comprobanción de inconsistencias :"
      Height          =   4545
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   5955
      Begin XtremeSuiteControls.PushButton bDocIva 
         Height          =   225
         Left            =   5010
         TabIndex        =   3
         Top             =   240
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Ver Datos"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   225
         Left            =   5010
         TabIndex        =   5
         Top             =   900
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Ver Datos"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   225
         Left            =   5010
         TabIndex        =   9
         Top             =   1560
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Ver Datos"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   225
         Left            =   5010
         TabIndex        =   10
         Top             =   2220
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Ver Datos"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   225
         Left            =   5040
         TabIndex        =   15
         Top             =   2880
         Width           =   825
         _Version        =   851968
         _ExtentX        =   1455
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Ver Datos"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton PushButton5 
         Height          =   225
         Left            =   5010
         TabIndex        =   16
         Top             =   3210
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Ver Datos"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton PushButton6 
         Height          =   195
         Left            =   5010
         TabIndex        =   17
         Top             =   3630
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Ver Datos"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton PushButton7 
         Height          =   225
         Left            =   5010
         TabIndex        =   18
         Top             =   3990
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Ver Datos"
         Appearance      =   6
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption6 
         Height          =   255
         Left            =   150
         TabIndex        =   24
         Top             =   4230
         Width           =   5775
         _Version        =   851968
         _ExtentX        =   10186
         _ExtentY        =   450
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   12632256
         GradientColorDark=   8421376
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption5 
         Height          =   225
         Left            =   150
         TabIndex        =   23
         Top             =   2490
         Width           =   5745
         _Version        =   851968
         _ExtentX        =   10134
         _ExtentY        =   397
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   12632256
         GradientColorDark=   49344
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption4 
         Height          =   255
         Left            =   150
         TabIndex        =   22
         Top             =   1830
         Width           =   5745
         _Version        =   851968
         _ExtentX        =   10134
         _ExtentY        =   450
         _StockProps     =   14
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
         GradientColorDark=   49152
         ForeColor       =   16777215
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   255
         Left            =   150
         TabIndex        =   21
         Top             =   1140
         Width           =   5745
         _Version        =   851968
         _ExtentX        =   10134
         _ExtentY        =   450
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   12632256
         GradientColorDark=   16744576
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   255
         Left            =   150
         TabIndex        =   20
         Top             =   510
         Width           =   5745
         _Version        =   851968
         _ExtentX        =   10134
         _ExtentY        =   450
         _StockProps     =   14
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
         GradientColorDark=   192
         ForeColor       =   16777215
      End
      Begin XtremeShortcutBar.ShortcutCaption vbancoasiento 
         Height          =   285
         Left            =   150
         TabIndex        =   14
         Top             =   3960
         Width           =   5775
         _Version        =   851968
         _ExtentX        =   10186
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "vbancoasiento"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   12632256
         GradientColorDark=   8421376
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   285
         Left            =   150
         TabIndex        =   13
         Top             =   3570
         Width           =   5745
         _Version        =   851968
         _ExtentX        =   10134
         _ExtentY        =   503
         _StockProps     =   14
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
         GradientColorDark=   33023
         ForeColor       =   16777215
      End
      Begin XtremeShortcutBar.ShortcutCaption vctacteasiento 
         Height          =   285
         Left            =   150
         TabIndex        =   12
         Top             =   3180
         Width           =   5745
         _Version        =   851968
         _ExtentX        =   10134
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "vctacteasiento"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   12632256
         GradientColorDark=   8421504
      End
      Begin XtremeShortcutBar.ShortcutCaption vasientoctate 
         Height          =   285
         Left            =   150
         TabIndex        =   11
         Top             =   2880
         Width           =   5745
         _Version        =   851968
         _ExtentX        =   10134
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "vasientoctate"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
         GradientColorDark=   12583104
         ForeColor       =   16777215
      End
      Begin XtremeShortcutBar.ShortcutCaption vdocctacte 
         Height          =   285
         Left            =   150
         TabIndex        =   8
         Top             =   2190
         Width           =   5745
         _Version        =   851968
         _ExtentX        =   10134
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "vdocctacte"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   12632256
         GradientColorDark=   49344
      End
      Begin XtremeShortcutBar.ShortcutCaption vctactedoc 
         Height          =   285
         Left            =   150
         TabIndex        =   7
         Top             =   1530
         Width           =   5745
         _Version        =   851968
         _ExtentX        =   10134
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "vctactedoc"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
         GradientColorDark=   49152
         ForeColor       =   16777215
      End
      Begin XtremeShortcutBar.ShortcutCaption vivadoc 
         Height          =   285
         Left            =   150
         TabIndex        =   6
         Top             =   870
         Width           =   5745
         _Version        =   851968
         _ExtentX        =   10134
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "vivadoc"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   12632256
         GradientColorDark=   16744576
      End
      Begin XtremeShortcutBar.ShortcutCaption vdociva 
         Height          =   315
         Left            =   150
         TabIndex        =   4
         Top             =   210
         Width           =   5745
         _Version        =   851968
         _ExtentX        =   10134
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "vdociva"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
         GradientColorDark=   192
         ForeColor       =   16777215
      End
   End
End
Attribute VB_Name = "frmInconsistencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vvcant As Integer

Private Sub bDocIva_Click()
Dim vsql As String
Dim vcant As Integer
vsql = "select * from factura where tipoMovimiento='FC' and tipoMovimiento='CD' and fecha >= '2011-05-01' and remito not in (select remito from ivafacturaventa where fecha >= '2011-05-01')"

If Val(fCantInco(vsql)) = 0 Then
    MsgBox "No hay errores para mostrar", vbInformation, "Documentos en Libro Iva"
    Exit Sub
End If

Call LlenarGrilla("nada", Me.grilla, vsql, "1")
End Sub

Private Sub Form_Load()
init
End Sub
Private Sub init()

Dim vsql As String
Dim vcant As Integer

vvcant = 0

Me.Height = 5220
Me.Width = 11565


' --------------- pctacte sin bancoscaja  -----------------




' -----------------------------------------------------------



'---------------------------- doc iva  ------------------------------------

vsql = "select * from factura where tipoMovimiento='FC' and tipoMovimiento='CD' and fecha >= '2011-05-01' and remito not in (select remito from ivafacturaventa where fecha >= '2011-05-01')"

vcant = Val(fCantInco(vsql))


If vcant > 0 Then
    Me.vdociva.Caption = "Hay " + fCantInco(vsql) + " Documentos s/ Iva Venta"

Else
    Me.vdociva.Caption = "OK ! Documentos en Iva Venta"
End If


'------------------------------------------------------------------------------


'---------------------------- iva doc  ------------------------------------

vsql = "select * from ivafacturaventa where ivafacturaventa.`idIvaFacturaVenta` >= 2898 and remito not in (select factura.remito  from factura inner join ivafacturaventa on  (factura.remito = ivafacturaventa.remito) where fecha >= '2011-05-01');"

vcant = Val(fCantInco(vsql))

If vcant > 0 Then
    Me.vivadoc.Caption = "Hay " + Str(vcant) + " Libro Iva s/ Documentos"

Else
    Me.vivadoc.Caption = "OK ! Libro Iva s/ Documentos"
End If


'------------------------------------------------------------------------------


'---------------------------- doc ctacte  ------------------------------------

vsql = "select * from factura where factura.fecha >= '2011-05-01' and remito not in (select remito  from cuentascorrientes where cuentascorrientes.fecha  >= '2011-05-01')"

vcant = Val(fCantInco(vsql))

If vcant > 0 Then
    Me.vdocctacte.Caption = "Hay " + fCantInco(vsql) + " Documentos s/ Ctacte"

Else
    Me.vdocctacte.Caption = "OK ! Documentos s/ Ctacte"
End If


'------------------------------------------------------------------------------

'---------------------------- factura ctacte  ------------------------------------

vsql = "select * from factura where factura.fecha >= '2011-06-10' and remito not in (select remito  from cuentascorrientes where cuentascorrientes.fecha  >= '2011-06-10')"
'vsql = "select * from factura where factura.fecha >= '2011-01-01' and remito not in (select remito  from cuentascorrientes where cuentascorrientes.fecha  >= '2011-01-01')"



vcant = Val(fCantInco(vsql))

If vcant > 0 Then
    Me.vdocctacte.Caption = "Hay " + fCantInco(vsql) + " Facturas s/ Ctacte"

Else
    Me.vdocctacte.Caption = "OK ! Documentos s/ Ctacte"
End If


'------------------------------------------------------------------------------


'----------------------------  ctacte  asiento ------------------------------------

'vsql = "select * from cuentascorrientes where fecha >= '2011-06-10' and nroasiento not in (select numero  from asientos where asientos.fecha  >= '2011-06-10')"

vsql = "select * from cuentascorrientes where fecha >= '2011-01-01' and NroInterno not in (select NroInterno  from asientos where asientos.fecha  >= '2011-01-01')"


vcant = Val(fCantInco(vsql))

If vcant > 0 Then
    Me.vctacteasiento.Caption = "Hay " + Str(vcant) + " Documentos s/ Asientos"

Else
    Me.vctacteasiento.Caption = "OK ! Documentos s/ Asientos"
End If


'------------------------------------------------------------------------------


'----------------------------  banco movi  asiento ------------------------------------


'vsql = "select * from bancosmovimientos where bancosmovimientos.NroAsiento > 0 and bancosmovimientos.fecha >= '2011-06-01' and bancosmovimientos.nroasiento not in (select numero  from asientos where asientos.fecha >= '2011-06-01')"
vsql = "select * from bancosmovimientos where bancosmovimientos.NroInterno > 0 and bancosmovimientos.fecha >= '2011-06-01' and bancosmovimientos.NroInterno not in (select NroInterno  from asientos where asientos.fecha >= '2011-06-01')"


vcant = Val(fCantInco(vsql))

If vcant > 0 Then
    Me.vbancoasiento.Caption = "Hay " + fCantInco(vsql) + " Banco/Caja s/ asientos"

Else
    Me.vbancoasiento.Caption = "OK ! Banco/Caja s/ asientos"
End If


'------------------------------------------------------------------------------

If vvcant > 0 Then
MsgBox "Hay alerta de inconsistencias de datos!" + Chr(13) + "Reporte este problema al servicio técnico", vbCritical, "Alerta importante..."

vdisplay.Caption = "Hay " + Str(vvcant) + " datos inconsistentes"
vdisplay.GradientColorDark = &HFF&
vdisplay.GradientColorLight = &HC0FFC0

Else
vdisplay.Caption = "No se registran problemas de datos"
vdisplay.GradientColorDark = &H8000&
vdisplay.GradientColorLight = &HC0FFC0
End If


End Sub


Function fCantInco(ByVal vsql As String) As String
vsql = "select count(*) as cant from (" + vsql + ") as t"
fCantInco = EsNulo((traerDatos2(vsql, "cant", pathDBMySQL)))
vvcant = vvcant + Val(fCantInco)
End Function

Private Sub PushButton1_Click()
Dim vsql As String
Dim vcant As Integer
vsql = "select * from ivafacturaventa where ivafacturaventa.`idIvaFacturaVenta` >= 2898 and remito not in (select factura.remito  from factura inner join ivafacturaventa on  (factura.remito = ivafacturaventa.remito) where fecha >= '2011-05-01');"
If Val(fCantInco(vsql)) = 0 Then
    MsgBox "No hay errores para mostrar", vbInformation, "Documentos en Libro Iva"
    Exit Sub
End If

Call LlenarGrilla("nada", Me.grilla, vsql, "1")

End Sub

Private Sub PushButton2_Click()
Dim vsql As String
Dim vcant As Integer
vsql = "(select * from ivafacturaventa where Remito >=2542 and Remito not in (select Remito from factura where fecha >= '2011-05-01' and remito >=2542))"

If Val(fCantInco(vsql)) = 0 Then
    MsgBox "No hay errores para mostrar", vbInformation, "Documentos en Libro Iva"
    Exit Sub
End If

Call LlenarGrilla("nada", Me.grilla, vsql, "1")
End Sub

Private Sub PushButton3_Click()
Dim vsql As String
Dim vcant As Integer
vsql = "select * from factura where factura.fecha >= '2011-05-01' and remito not in (select remito  from cuentascorrientes where cuentascorrientes.fecha  >= '2011-05-01')"

If Val(fCantInco(vsql)) = 0 Then
    MsgBox "No hay errores para mostrar", vbInformation, "Documentos en Libro Iva"
    Exit Sub
End If

Call LlenarGrilla("nada", Me.grilla, vsql, "1")
End Sub

Private Sub PushButton5_Click()
Dim vsql As String
Dim vcant As Integer
vsql = "select * from cuentascorrientes where fecha >= '2011-05-01' and nroasiento not in (select numero  from asientos where asientos.fecha  >= '2011-05-01')"
If Val(fCantInco(vsql)) = 0 Then
    MsgBox "No hay errores para mostrar", vbInformation, "Documentos en Libro Iva"
    Exit Sub
End If

Call LlenarGrilla("nada", Me.grilla, vsql, "1")

End Sub

Private Sub PushButton7_Click()
Dim vsql As String
Dim vcant As Integer
vsql = "select * from bancosmovimientos where bancosmovimientos.`idBancosMovimientos` > 23416 and nroasiento not in (select numero  from asientos where asientos.`idAsientos`   >= 32645)"
If Val(fCantInco(vsql)) = 0 Then
    MsgBox "No hay errores para mostrar", vbInformation, "Documentos en Libro Iva"
    Exit Sub
End If

Call LlenarGrilla("nada", Me.grilla, vsql, "1")
End Sub

