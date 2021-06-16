VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "KlexGrid.ocx"
Begin VB.Form frmNroFactNC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de Factura y artículos para la Nota de Crédito"
   ClientHeight    =   7245
   ClientLeft      =   2760
   ClientTop       =   3675
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   12135
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox vPuntoDeVenta 
      Height          =   315
      Left            =   3960
      TabIndex        =   16
      Top             =   2340
      Width           =   1455
   End
   Begin VB.TextBox vLetra 
      Height          =   315
      Left            =   2520
      TabIndex        =   15
      Top             =   2340
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Height          =   675
      Left            =   60
      TabIndex        =   7
      Top             =   6060
      Width           =   12045
      Begin VB.Label vdisplay 
         ForeColor       =   &H000000C0&
         Height          =   345
         Left            =   5610
         TabIndex        =   12
         Top             =   210
         Width           =   2715
      End
      Begin VB.Label lblTotalFactura 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   3
         Left            =   4260
         TabIndex        =   11
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label lblTotalFactura 
         Caption         =   "Total Seleccionado:"
         Height          =   195
         Index           =   2
         Left            =   2730
         TabIndex        =   10
         Top             =   270
         Width           =   1605
      End
      Begin VB.Label lblTotalFactura 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   1
         Left            =   1350
         TabIndex        =   9
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label lblTotalFactura 
         Caption         =   "Total factura:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   270
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Facturas del cliente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12075
      Begin Grid.KlexGrid KlexFactura 
         Height          =   1935
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   3413
         EnterKeyBehaviour=   0
         BackColorAlternate=   0
         GridLines       =   0
         GridLinesFixed  =   2
         BackColorFixed  =   -2147483626
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColorFixed  =   8421504
         MouseIcon       =   "frmNroFactNC.frx":0000
         ScrollBars      =   2
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Detalles de las facturas seleccionadas:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3285
      Left            =   30
      TabIndex        =   4
      Top             =   2730
      Width           =   12075
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla2 
         Height          =   2775
         Left            =   90
         TabIndex        =   13
         Top             =   450
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   4895
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   27
         BackColorFixed  =   -2147483648
         ForeColorFixed  =   4210752
         BackColorSel    =   255
         FocusRect       =   2
         GridLines       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   27
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label lblSeleccioneLos 
         Caption         =   "Seleccione los artículos para la Nota de Créditos"
         Height          =   255
         Left            =   3570
         TabIndex        =   5
         Top             =   180
         Width           =   8115
      End
   End
   Begin VB.TextBox txtNroFactura 
      Height          =   315
      Left            =   5640
      TabIndex        =   2
      Top             =   2340
      Width           =   2385
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   345
      Left            =   90
      TabIndex        =   1
      Top             =   6840
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   1500
      TabIndex        =   0
      Top             =   6840
      Width           =   1365
   End
   Begin VB.Label Label11 
      Caption         =   "Factura seleccionada: "
      Height          =   225
      Left            =   60
      TabIndex        =   3
      Top             =   2370
      Width           =   2325
   End
End
Attribute VB_Name = "frmNroFactNC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsNotaC As ADODB.Recordset, sqlNotaC As String
Dim rsDetalle As ADODB.Recordset, sqlDetalle As String
Private Sub cmdAceptar_Click()
On Error Resume Next

    If Not Trim(txtNroFactura.Text) = "" Then
        frmRemito.vNroFacturaNotaC = Val(txtNroFactura.Text)
        frmRemito.vLetraNotaC = vLetra.Text
        frmRemito.vPuntoDeVentaNotaC = vPuntoDeVenta.Text
    
        rsNotaC.Close
    
        GrabarRemito
    
        rsDetalle.Close
        Unload Me
    End If

If Err Then
    Unload Me
    Exit Sub
    Unload Me
End If
End Sub
Private Sub GrabarRemito()
On Error Resume Next

    Dim i As Integer

    With grilla2
    
        For i = 1 To .Rows - 1
            .Col = 1
            .Row = i
        
            If .CellBackColor = &HFFC0C0 Then
                
                frmRemito.txtDetalle(0).Text = .TextMatrix(i, 2)
                frmRemito.txtDetalle(1).Tag = .TextMatrix(i, 3)
                frmRemito.txtDetalle(1).Text = .TextMatrix(i, 4)
                frmRemito.txtDetalle(2).Text = .TextMatrix(i, 5)
                frmRemito.txtDetalle(3).Text = .TextMatrix(i, 6)
                frmRemito.txtDetalle(4).Text = .TextMatrix(i, 7)
                frmRemito.txtDetalle(5).Text = 0
                frmRemito.txtDetalle(6).Text = .TextMatrix(i, 8)
                frmRemito.GrabarRenglon
        
            End If
    
        Next

    End With

If Err Then GrabarLog "GrabarRemito", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Load()
On Error Resume Next

    Set rsNotaC = New ADODB.Recordset
    
    sqlNotaC = "SELECT codigo,Letra, PuntoDeVenta, ncomprobante, fecha, cuit, subtotal, total, tipo, remito FROM factura WHERE ((tipo = 'Fact A') OR (tipo = 'Fact B')) AND (codigo = '" & frmRemito.txtCliente(0).Tag & "') order by fecha desc"
    
    rsNotaC.Open sqlNotaC, ConnDDBB, adOpenKeyset, adLockOptimistic

    If rsNotaC.EOF Then
        MsgBox "No existen facturas para este cliente"
       ' rsNotaC.Close
       ' Unload Me
        Exit Sub
    End If
    
    ConfigurarGrilla
    ConfigurarGrilla2

    Set KlexFactura.Recordset = rsNotaC
    grilla2.Clear

If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub ConfigurarGrilla()
On Error Resume Next

    Dim i As Integer
    
    With KlexFactura
  
    
        .Cols = 12

        .TextMatrix(0, 1) = "Sel."
        .TextMatrix(0, 2) = "Codigo"
        .TextMatrix(0, 3) = "Letra"
        .TextMatrix(0, 4) = "PVenta"
        .TextMatrix(0, 5) = "NroFact"
        .TextMatrix(0, 6) = "Fecha"
        .TextMatrix(0, 7) = "Cuit"
        .TextMatrix(0, 8) = "Sutotal"
        .TextMatrix(0, 9) = "Total"
        .TextMatrix(0, 10) = "Doc."
        .TextMatrix(0, 11) = "Remito"

        .ColWidth(0) = 400
        .ColWidth(1) = 100
        .ColWidth(2) = 0
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
        
        .ColWidth(8) = 2000
        .ColDisplayFormat(8) = "######0.00"
                
        
        .ColWidth(9) = 1200
        .ColDisplayFormat(9) = "######0.00"
                
        .ColWidth(10) = 1200
        .ColWidth(11) = 0

        'For i = 0 To .Cols - 1
        '    .ColAlignmentFixed(i) = 4
        'Next i

    End With

If Err Then GrabarLog "ConfigurarGrilla", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub formatgrilla2()
Dim j As Integer
  
With grilla2

For j = 1 To .Rows - 1
            .TextMatrix(j, 7) = Format(.TextMatrix(j, 7), "#######0.0000")
            .TextMatrix(j, 6) = Format(.TextMatrix(j, 6), "#######0.0000")
            .TextMatrix(j, 5) = Format(.TextMatrix(j, 5), "#######0.0000")
Next j
 
 End With
            
End Sub
Private Sub ConfigurarGrilla2()
Dim i, j As Integer

    With grilla2
        .Refresh
        .Clear

        .Cols = 8
        .TextMatrix(0, 0) = "Sel."
        .TextMatrix(0, 1) = "Remito"
        .TextMatrix(0, 2) = "Cantidad"
        .TextMatrix(0, 3) = "Código"
        .TextMatrix(0, 4) = "Detalle"
        .TextMatrix(0, 5) = "Precio"
        .TextMatrix(0, 6) = "% Iva"
        .TextMatrix(0, 7) = "Total"
        
        .ColWidth(0) = 400
        .ColWidth(1) = 0
        .ColWidth(2) = 1000
        .ColWidth(3) = 1500
        .ColWidth(4) = 3000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
        
        
 
    End With
    
If Err Then GrabarLog "configurarGrilla2", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Terminate()
On Error Resume Next
    
    rsNotaC.Close
    rsDetalle.Close

If Err Then GrabarLog "Form_Terminate", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    rsNotaC.Close
    rsDetalle.Close

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub KlexFactura_Click()
On Error Resume Next
    
    'Le Paso el nro de Remito
    txtNroFactura.Text = KlexFactura.TextMatrix(KlexFactura.Row, 4)
    vLetra.Text = KlexFactura.TextMatrix(KlexFactura.Row, 2)
    vPuntoDeVenta.Text = KlexFactura.TextMatrix(KlexFactura.Row, 3)
    'cmdncok_Click
    'sql = "Select * from fdetalle where remito=" + Str(Rec!remito)
    
    Set rsDetalle = New ADODB.Recordset
    
    sqlDetalle = "SELECT remito, cantidad, codigo, detalle,   precio  , descuento, tiva ,  total FROM fdetalle WHERE (remito = " & Val(KlexFactura.TextMatrix(KlexFactura.Row, 10)) & ")"

    With rsDetalle
        If .State = 1 Then .Close
        Call .Open(sqlDetalle, ConnDDBB, adOpenKeyset, adLockOptimistic)

        Set grilla2.DataSource = rsDetalle

        '.Close
    End With
    
    formatgrilla2
    
If Err Then GrabarLog "KlexFactura_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub grilla2_Click()
On Error Resume Next

    Dim i As Integer
    
    vdisplay.Caption = "Seleccionando artículos ..."
            
    With grilla2
        .Col = 0
        .SetFocus
            
        For i = 1 To 6
            .Col = i

            If .CellBackColor = &HFFC0C0 Then
                .CellBackColor = vbWhite
            Else
                .CellBackColor = &HFFC0C0
            End If

        Next
    '    .Refresh
    End With
    
If Err Then GrabarLog "grilla2_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtNroFactura_keyPress(Keyascii As Integer)
On Error Resume Next

    If Keyascii = 13 Then
        sqlNotaC = "SELECT codigo, ncomprobante, fecha, cuit, subtotal, total, tipo, nrofactnc  FROM factura WHERE NOT (nrofactnc ='usada') AND (tipo = 'Fact A') AND (tipo = 'Fact B') AND (NComprobante = " & Val(txtNroFactura.Text) & ")"
        
        rsNotaC.Close
        rsNotaC.Open sqlNotaC, ConnDDBB, adOpenKeyset, adLockOptimistic

        If rsNotaC.EOF Then
            MsgBox "No existen facturas para este cliente"
            rsNotaC.Close
            Unload Me
            Exit Sub
        End If

        ConfigurarGrilla
        
        Set KlexFactura.Recordset = rsNotaC
        rsNotaC.Close
    End If

If Err Then GrabarLog "txtNroFactura_keyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
