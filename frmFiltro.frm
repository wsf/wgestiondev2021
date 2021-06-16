VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{9746E3DA-06E1-4D26-9CE4-D9F6411A9C70}#1.0#0"; "SMGA_OcxTxt2009.ocx"
Begin VB.Form frmFiltro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parámetros para el filtro:"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   12255
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.PushButton PusExcel 
      Height          =   405
      Left            =   1950
      TabIndex        =   8
      Top             =   6240
      Width           =   1755
      _Version        =   851968
      _ExtentX        =   3096
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Excel"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PusRefresh 
      Height          =   345
      Left            =   10950
      TabIndex        =   7
      Top             =   90
      Width           =   1095
      _Version        =   851968
      _ExtentX        =   1931
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Actualizar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PusImprimir 
      Height          =   405
      Left            =   210
      TabIndex        =   6
      Top             =   6240
      Width           =   1695
      _Version        =   851968
      _ExtentX        =   2990
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Imprimir"
      UseVisualStyle  =   -1  'True
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla 
      Height          =   5595
      Left            =   210
      TabIndex        =   5
      Top             =   540
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   9869
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin Aplisoft_CajasDeTexto.TxF fdesde 
      Height          =   345
      Left            =   1170
      TabIndex        =   0
      Top             =   60
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   609
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
   Begin XtremeSuiteControls.PushButton PusAceptar 
      Height          =   375
      Left            =   7620
      TabIndex        =   4
      Top             =   60
      Width           =   885
      _Version        =   851968
      _ExtentX        =   1561
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Filtar"
      UseVisualStyle  =   -1  'True
   End
   Begin Aplisoft_CajasDeTexto.TxF fhasta 
      Height          =   345
      Left            =   4950
      TabIndex        =   1
      Top             =   60
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   609
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
   Begin XtremeSuiteControls.Label lblFechaHasta 
      Height          =   285
      Left            =   3960
      TabIndex        =   3
      Top             =   90
      Width           =   945
      _Version        =   851968
      _ExtentX        =   1667
      _ExtentY        =   503
      _StockProps     =   79
      Caption         =   "Fecha hasta:"
   End
   Begin XtremeSuiteControls.Label lblFechaDesde 
      Height          =   225
      Left            =   150
      TabIndex        =   2
      Top             =   120
      Width           =   975
      _Version        =   851968
      _ExtentX        =   1720
      _ExtentY        =   397
      _StockProps     =   79
      Caption         =   "Fecha desde:"
   End
End
Attribute VB_Name = "frmFiltro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vcomando As String
Dim vcondi As String
Dim rs As New ADODB.Recordset


Private Sub fdesde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    fhasta.SetFocus
End If
End Sub

Private Sub fhasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.PusAceptar.SetFocus
End If

End Sub

Private Sub Form_Load()


' ---- init

Me.fdesde.Value = Date
Me.fhasta.Value = Date

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2 - 1000

If vcomando = "fae" Then Me.Caption = "Listado para la rendición del F.A.E."

End Sub

Private Sub PusAceptar_Click()
On Error Resume Next
Dim vcondi2 As String
Dim vsql As String
vcondi2 = " fecha>='" + strfechaMySQL(fdesde) + "' and  fecha <= '" + strfechaMySQL(fhasta) + "' "
vsql = fsqlFaeDetalle2(vcondi2)

With rs

Screen.MousePointer = vbHourglass
    Call .Open(vsql, ConnDDBB, adOpenStatic, adLockReadOnly)
Screen.MousePointer = vbDefault

setgrilla

    Set grilla.DataSource = .DataSource
    
End With


Call PusRefresh_Click
If Err Then Exit Sub
End Sub

Private Sub setgrilla()
grilla.ColWidth(0) = 100
grilla.ColWidth(1) = 1200
grilla.ColWidth(2) = 1200
grilla.ColWidth(3) = 1200
grilla.ColWidth(4) = 5000
grilla.ColWidth(5) = 1700

grilla.ColAlignment(1) = 0

End Sub

Public Sub faedetalle()
Dim vul, sqlSaldoCaja As String
On Error Resume Next
    
    
    Unload Mantenimiento
    Load Mantenimiento
    
    With Mantenimiento.rsFAEDetalle
        If .State = 1 Then .Close
        
      .Source = fsqlFaeDetalle(vcondi)


        
        If Not .State = 1 Then .Open
        .Close
        .Open
        
    End With
    
    With drFaeDetalle
    .Show
    End With

If Err Then Exit Sub

End Sub

Private Sub PusExcel_Click()
On Error Resume Next

Call grillaToExcel(Me.grilla)

Exit Sub


Dim iErr As Integer
            iErr = 0
            On Error GoTo Proc_Err
            Screen.MousePointer = vbHourglass
            Dim i, II  As Long
            Dim X  As Long
            Dim Cols As Integer
            Dim Rows As Integer
            Dim sLine As String
            Open App.Path + "\l.csv" For Output As #1
            Cols = Me.grilla.Cols
            Rows = rs.RecordCount
            
           rs.MoveFirst
            
           For II = 0 To grilla.Rows - 1
            
            'Do While Not rs.EOF
               sLine = ""
               For i = 1 To Cols - 1
                  'sLine = sLine & rs.Fields(i).Value & IIf(i < Cols - 1, ";", "")
                  sLine = sLine & grilla.TextMatrix(II, i) & IIf(i < Cols - 1, ";", "")
               Next i
               Print #1, sLine
              ' rs.MoveNext
            'Loop
            
            Next
            
            Close #1
    
    Call Shell("excel.bat", 1)

Proc_Exit:
            Screen.MousePointer = vbDefault
            Exit Sub
Proc_Err:
            If iErr > 3 Then
               ' Log your error here...
               Resume Proc_Exit
            Else
               iErr = iErr + 1
               Resume
            End If
            
End Sub

Private Sub PusImprimir_Click()
    Call imprimirGrilla(Me.grilla, 8)
End Sub

Private Sub PusRefresh_Click()
Dim i As Integer
Dim valor, vlinea As String
Dim vt1, vt2, vt3 As Double

vt1 = 0
vt2 = 0
vt3 = 0


For i = 1 To grilla.Rows - 1

        valor = grilla.TextMatrix(i, 1)
        vt1 = vt1 + Val(Replace(valor, ",", ""))
        grilla.TextMatrix(i, 1) = Format(valor, "###,###,##0.00")
        
        valor = grilla.TextMatrix(i, 2)
        vt2 = vt2 + Val(Replace(valor, ",", ""))
        grilla.TextMatrix(i, 2) = Format(valor, "###,###,##0.00")
        
        valor = grilla.TextMatrix(i, 3)
        vt3 = vt3 + Val(Replace(valor, ",", ""))
        grilla.TextMatrix(i, 3) = Format(valor, "###,###,##0.00")

Next


vlinea = "" + vbTab + Format(vt1, "###,###,##0.00") + vbTab + Format(vt2, "###,###,##0.00") + vbTab + Format(vt3, "###,###,##0.00") + vbTab + "TOTALES"
grilla.AddItem vlinea

vlinea = ""
grilla.AddItem vlinea

vlinea = "" + vbTab + "Periodo: "
grilla.AddItem vlinea

vlinea = "" + vbTab + "F.Desde:" + vbTab + Format(Me.fdesde, "dd/mm/yyyy") + vbTab + " F.Hasta :  " + vbTab + Format(Me.fhasta, "dd/mm/yyyy")
grilla.AddItem vlinea

End Sub
