VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "Copia de KlexGrid.ocx"
Begin VB.Form frmAsientosTipo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asientos Tipo"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8805
   Begin XtremeSuiteControls.GroupBox GBSeleccionAsiento 
      Height          =   795
      Left            =   0
      TabIndex        =   1
      Top             =   5040
      Width           =   8745
      _Version        =   851968
      _ExtentX        =   15425
      _ExtentY        =   1402
      _StockProps     =   79
      Caption         =   "Seleccion de Asiento"
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtAsiento 
         Height          =   315
         Index           =   0
         Left            =   3180
         TabIndex        =   2
         Top             =   285
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
         MaxLength       =   10
      End
      Begin XtremeSuiteControls.FlatEdit txtAsiento 
         Height          =   315
         Index           =   1
         Left            =   6840
         TabIndex        =   4
         Top             =   285
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
         MaxLength       =   10
      End
      Begin XtremeSuiteControls.Label lblLuegoPresione 
         Height          =   255
         Index           =   2
         Left            =   90
         TabIndex        =   10
         Top             =   480
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3096
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Luego presione <Enter>"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLuegoPresione 
         Height          =   255
         Index           =   1
         Left            =   5160
         TabIndex        =   5
         Top             =   315
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Importe del Asiento :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblLuegoPresione 
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   3
         Top             =   285
         Width           =   3135
         _Version        =   851968
         _ExtentX        =   5530
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Seleccione un Numero de Asiento Tipo :"
         Transparent     =   -1  'True
      End
   End
   Begin Grid.KlexGrid KlexAsientosTipo 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   570
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7858
      EnterKeyBehaviour=   0
      BackColorAlternate=   12632256
      GridLinesFixed  =   2
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
      GridColorFixed  =   8421504
      MouseIcon       =   "frmAsientosTipo.frx":0000
      Rows            =   10
   End
   Begin XtremeSuiteControls.PushButton cmdAcciones 
      Height          =   375
      Index           =   2
      Left            =   7530
      TabIndex        =   6
      Top             =   30
      Width           =   1245
      _Version        =   851968
      _ExtentX        =   2196
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Cerrar"
      UseVisualStyle  =   -1  'True
      Picture         =   "frmAsientosTipo.frx":001C
   End
   Begin XtremeSuiteControls.PushButton cmdAcciones 
      Height          =   375
      Index           =   0
      Left            =   30
      TabIndex        =   7
      Top             =   30
      Width           =   1455
      _Version        =   851968
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Modificar"
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
      Picture         =   "frmAsientosTipo.frx":041C
   End
   Begin XtremeSuiteControls.PushButton cmdAcciones 
      Height          =   375
      Index           =   1
      Left            =   1500
      TabIndex        =   8
      Top             =   30
      Width           =   1455
      _Version        =   851968
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Borrar"
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
      Picture         =   "frmAsientosTipo.frx":09B6
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   135
      Left            =   -30
      TabIndex        =   9
      Top             =   360
      Width           =   8835
      _Version        =   851968
      _ExtentX        =   15584
      _ExtentY        =   238
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
   End
   Begin XtremeSuiteControls.Label lblLuegoPresione 
      Height          =   255
      Index           =   3
      Left            =   3720
      TabIndex        =   11
      Top             =   30
      Width           =   3165
      _Version        =   851968
      _ExtentX        =   5583
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Doble clic para seleccionar un asiento tipo"
      Transparent     =   -1  'True
   End
End
Attribute VB_Name = "frmAsientosTipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAcciones_Click(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 0
        
        Case 1
        
        Case 2
            If Val(txtAsiento(1).Text) = 0 Then
                If Not MsgBox("El importe para el Asiento Tipo Nº : " & EsNulo(txtAsiento(0).Text) & " Es 0 " & vbCrLf & " Desea cargar el asiento de todas maneras?", vbExclamation + vbYesNo, "Mensaje ...") = vbYes Then
                
                
                    Exit Sub
                End If
            End If
    
            Call frmAsientosAlta.CargarDetalleAsiento(Trim(txtAsiento(0).Text), Val(txtAsiento(1).Text))
            Unload Me
    
    End Select

If Err Then GrabarLog "cmdAcciones_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Load()
On Error Resume Next
    
    With Me
        .Show
        
    End With
    
    Call CargarAsietosTipo

    KlexAsientosTipo.TopRow = Val(KlexAsientosTipo.Rows - 1)

    txtAsiento(0).SetFocus

If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CargarAsietosTipo()
    On Error Resume Next
    
    Dim rsAsientosTipo As New ADODB.Recordset, sqlAsientosTipo As String, i As Integer
    
    sqlAsientosTipo = "SELECT idAsientosTipo, Numero, AsientosTipo.CodigoCuenta, Cuentas.Cuenta, DebeHaber, Porcentaje FROM AsientosTipo INNER JOIN Cuentas ON AsientosTipo.CodigoCuenta=Cuentas.CodigoCuenta ORDER BY Numero;"
    
    With rsAsientosTipo
        .CursorLocation = adUseClient
        Call .Open(sqlAsientosTipo, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then
            .MoveFirst
            FormatoGrilla (.RecordCount)
        Else
            FormatoGrilla (1)
        End If
        
        i = 1
        
        Do Until .EOF = True
            
            
            'If Not "[" & KlexAsientosTipo.TextMatrix(i, 2) & "]" = "[" & EsNulo(.Fields("Numero").Value) & "]" Then
            
            
            KlexAsientosTipo.TextMatrix(i, 0) = ""
            KlexAsientosTipo.TextMatrix(i, 1) = EsNulo(.Fields("idAsientosTipo").Value)
            KlexAsientosTipo.TextMatrix(i, 2) = "[" & EsNulo(.Fields("Numero").Value) & "]"
            KlexAsientosTipo.TextMatrix(i, 3) = EsNulo(.Fields("CodigoCuenta").Value)
            KlexAsientosTipo.TextMatrix(i, 4) = EsNulo(.Fields("Cuenta").Value)
            KlexAsientosTipo.TextMatrix(i, 5) = EsNulo(.Fields("DebeHaber").Value)
            KlexAsientosTipo.TextMatrix(i, 6) = EsNulo(.Fields("Porcentaje").Value)

            .MoveNext
        
            i = i + 1
        Loop
        
    End With
    
    sqlAsientosTipo = ""
    
    If rsAsientosTipo.State = 1 Then
        rsAsientosTipo.Close
        Set rsAsientosTipo = Nothing
    End If
    
If Err Then GrabarLog "CargarAsietosTipo", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub FormatoGrilla(vCantidadRenglones As Integer)
On Error Resume Next

    Dim i As Integer

    With KlexAsientosTipo
        .FixedRows = 1
        .FixedCols = 1
    
        .Cols = 7
        .Rows = vCantidadRenglones + 1
        
        If vCantidadRenglones = 1 Then
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
                .ColWidth(i) = 0
            Next
        End If
        
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 400
        
        .TextMatrix(0, 1) = "idAsientosTipo"
        .ColWidth(1) = 0
               
        .TextMatrix(0, 2) = "Nº"
        .ColWidth(2) = 1000
        
        .TextMatrix(0, 3) = "Cuenta"
        .ColWidth(3) = 1000
        
        .TextMatrix(0, 4) = "Descripcion"
        .ColWidth(4) = 2250
        
        .TextMatrix(0, 5) = "Debe/Haber"
        .ColWidth(5) = 1000
        
        .TextMatrix(0, 6) = "Porcentaje"
        .ColWidth(6) = 1000
        
    End With
    
If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub KlexAsientosTipo_DblClick()
Dim vlen As Integer
Dim v As String

v = Me.KlexAsientosTipo.TextMatrix(Me.KlexAsientosTipo.Row, 2)
vlen = Len(v) 'Me.txtAsiento(0).Text = v

Me.txtAsiento(0).Text = Right(Left(v, vlen - 1), vlen - 2)

Call txtAsiento_KeyPress(0, 13)

End Sub

Private Sub txtAsiento_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    
    If KeyAscii = 13 Then
    
        Select Case Index
    
            Case 0
                BusquedaAsientoTipo
                txtAsiento(1).SetFocus
            
            Case 1
                cmdAcciones(2).SetFocus
        
        End Select
    End If
    
If Err Then GrabarLog "txtAsiento_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub BusquedaAsientoTipo()
On Error Resume Next

    Dim i As Integer, j As Integer


    If Not Trim(txtAsiento(0).Text) = "" Then
    
        With KlexAsientosTipo
            .Row = 1
            For i = 1 To Val(.Rows - 1)
                If .TextMatrix(i, 2) = "[" & Trim(txtAsiento(0).Text) & "]" Then
                    .Row = i
                    .TopRow = i
                    
                    For j = 1 To Val(.Cols - 1)
                        .Col = j
                        .CellBackColor = vbGreen
                    Next
                    
                    
                Else
                    'If chkOcultarNoBuscados.Value = xtpChecked Then
                    '    .RowHeight(i) = 0
                    'End If
                End If
            
            Next
        
        
            'Randomize
            'vColor = Val(Rnd * vColorInicial)

        End With
        
    
    End If


If Err Then GrabarLog "BusquedaAsientoTipo", Err.Number & " " & Err.Description, Me.Caption
End Sub
