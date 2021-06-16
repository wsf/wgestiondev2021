VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "KlexGrid.ocx"
Object = "{9746E3DA-06E1-4D26-9CE4-D9F6411A9C70}#1.0#0"; "SMGA_OcxTxt2009.ocx"
Begin VB.Form frmBorrarBases 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Borrado de Datos Almacenados"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   13530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   13530
   Begin Aplisoft_CajasDeTexto.TxF dtpFecha 
      Height          =   375
      Left            =   3120
      TabIndex        =   30
      Top             =   360
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
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
   Begin XtremeSuiteControls.PushButton PusEjecutar 
      Height          =   375
      Left            =   3090
      TabIndex        =   29
      Top             =   5400
      Width           =   1545
      _Version        =   851968
      _ExtentX        =   2725
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Ejecutar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit vnrointerno 
      Height          =   345
      Left            =   1410
      TabIndex        =   28
      Top             =   5370
      Width           =   1545
      _Version        =   851968
      _ExtentX        =   2725
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.GroupBox GBClientesADepurar 
      Height          =   5805
      Left            =   4680
      TabIndex        =   19
      Top             =   120
      Width           =   8775
      _Version        =   851968
      _ExtentX        =   15478
      _ExtentY        =   10239
      _StockProps     =   79
      Caption         =   "Clientes a Depurar"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton PBSeleccionar 
         Height          =   345
         Index           =   0
         Left            =   5460
         TabIndex        =   21
         Top             =   5340
         Width           =   1515
         _Version        =   851968
         _ExtentX        =   2672
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Seleccionar Todos"
         UseVisualStyle  =   -1  'True
      End
      Begin Grid.KlexGrid KlexClientes 
         Height          =   4995
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   8811
         EnterKeyBehaviour=   0
         BackColorAlternate=   0
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
         MouseIcon       =   "frmBorrarBases.frx":0000
         Rows            =   10
      End
      Begin XtremeSuiteControls.PushButton PBSeleccionar 
         Height          =   345
         Index           =   1
         Left            =   7020
         TabIndex        =   22
         Top             =   5340
         Width           =   1605
         _Version        =   851968
         _ExtentX        =   2831
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Seleccionar Ninguno"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin VB.Frame fraFecha 
      Caption         =   "Fecha a Depurar"
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   4515
      Begin VB.Label lblFecha 
         AutoSize        =   -1  'True
         Caption         =   "> Borrar desde esta Fecha ( Menor )"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   330
         Width           =   2640
      End
   End
   Begin VB.Frame Frame4 
      Height          =   525
      Left            =   90
      TabIndex        =   6
      Top             =   6030
      Width           =   4515
      Begin VB.OptionButton o 
         Caption         =   "Borrar Todos los datos del programa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   180
         Width           =   3555
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Información que desea eliminar :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4365
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4515
      Begin VB.CheckBox c 
         Caption         =   "Eliminar Log de caja"
         Height          =   255
         Index           =   10
         Left            =   300
         TabIndex        =   24
         Top             =   4050
         Width           =   3015
      End
      Begin VB.CheckBox c 
         Caption         =   "Eliminar asientos"
         Height          =   255
         Index           =   9
         Left            =   300
         TabIndex        =   23
         Top             =   3720
         Width           =   3015
      End
      Begin VB.Frame Frame7 
         Height          =   30
         Left            =   60
         TabIndex        =   16
         Top             =   3270
         Width           =   4395
      End
      Begin VB.Frame Frame6 
         Height          =   30
         Left            =   60
         TabIndex        =   15
         Top             =   2850
         Width           =   4395
      End
      Begin VB.Frame Frame5 
         Height          =   30
         Left            =   60
         TabIndex        =   14
         Top             =   2130
         Width           =   4395
      End
      Begin VB.Frame Frame3 
         Height          =   30
         Left            =   60
         TabIndex        =   13
         Top             =   630
         Width           =   4395
      End
      Begin VB.Frame Frame2 
         Height          =   30
         Left            =   60
         TabIndex        =   12
         Top             =   1380
         Width           =   4395
      End
      Begin VB.CheckBox c 
         Caption         =   "Eliminar todos los movimientos de Caja."
         Height          =   255
         Index           =   8
         Left            =   300
         TabIndex        =   11
         Top             =   3390
         Width           =   3495
      End
      Begin VB.CheckBox c 
         Caption         =   "Eliminar todos los cheques."
         Height          =   255
         Index           =   7
         Left            =   300
         TabIndex        =   10
         Top             =   2970
         Width           =   3495
      End
      Begin VB.CheckBox c 
         Caption         =   "Eliminar Cuentas Corrientes Proveedores."
         Height          =   255
         Index           =   6
         Left            =   300
         TabIndex        =   9
         Top             =   2550
         Width           =   3495
      End
      Begin VB.CheckBox c 
         Caption         =   "Eliminar todos los documentos de Ventas."
         Height          =   255
         Index           =   5
         Left            =   300
         TabIndex        =   8
         Top             =   1500
         Width           =   4065
      End
      Begin VB.CheckBox c 
         Caption         =   "Eliminar Cuentas Corrientes Clientes."
         Height          =   255
         Index           =   4
         Left            =   300
         TabIndex        =   5
         Top             =   2220
         Width           =   3135
      End
      Begin VB.CheckBox c 
         Caption         =   "Eliminar todos los documentos de  Compras."
         Height          =   255
         Index           =   3
         Left            =   300
         TabIndex        =   4
         Top             =   1800
         Width           =   3825
      End
      Begin VB.CheckBox c 
         Caption         =   "Eliminar todos los Proveedores. "
         Height          =   255
         Index           =   2
         Left            =   300
         TabIndex        =   3
         Top             =   1050
         Width           =   3135
      End
      Begin VB.CheckBox c 
         Caption         =   "Eliminar todos los Clientes."
         Height          =   255
         Index           =   1
         Left            =   300
         TabIndex        =   2
         Top             =   750
         Width           =   4035
      End
      Begin VB.CheckBox c 
         Caption         =   "Eliminar todos los Artículos."
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   1
         Top             =   360
         Width           =   2415
      End
   End
   Begin XtremeSuiteControls.PushButton PbAcciones 
      Height          =   375
      Index           =   0
      Left            =   11010
      TabIndex        =   25
      Top             =   6120
      Width           =   1215
      _Version        =   851968
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Depurar"
      UseVisualStyle  =   -1  'True
      Picture         =   "frmBorrarBases.frx":001C
      BorderGap       =   10
   End
   Begin XtremeSuiteControls.PushButton PbAcciones 
      Height          =   375
      Index           =   1
      Left            =   12210
      TabIndex        =   26
      Top             =   6120
      Width           =   1215
      _Version        =   851968
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Cerrar"
      UseVisualStyle  =   -1  'True
      Picture         =   "frmBorrarBases.frx":045E
   End
   Begin XtremeSuiteControls.Label lblNroInternos 
      Height          =   285
      Left            =   150
      TabIndex        =   27
      Top             =   5400
      Width           =   1065
      _Version        =   851968
      _ExtentX        =   1879
      _ExtentY        =   503
      _StockProps     =   79
      Caption         =   "Nro. Internos:"
   End
End
Attribute VB_Name = "frmBorrarBases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub FormatoGrilla(vCantidadRenglones As Integer)
On Error Resume Next

    Dim k As Integer

    With KlexClientes
        .FixedRows = 1
        .FixedCols = 1
    
        .Cols = 8
        .Rows = vCantidadRenglones + 1
        
        If vCantidadRenglones = 1 Then
            For k = 0 To .Cols - 1
                .TextMatrix(1, k) = ""
                .ColWidth(k) = 0
            Next
        End If
        
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 400
        
        .TextMatrix(0, 1) = "idClientes"
        .ColWidth(1) = 0
               
        .TextMatrix(0, 2) = "Codigo"
        .ColWidth(2) = 850
        
        .TextMatrix(0, 3) = "Nombre"
        .ColWidth(3) = 2500
        
        .TextMatrix(0, 4) = "Direccion"
        .ColWidth(4) = 1250

        
        .TextMatrix(0, 5) = "Localidad"
        .ColWidth(5) = 1000
        
        .TextMatrix(0, 6) = "Telefono"
        .ColWidth(6) = 1000
        
        .TextMatrix(0, 7) = "Borrar"
        .ColWidth(7) = 750

        .BackColorAlternate = 14737632
    End With
    
If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub CargarClientesEnGrilla()
On Error Resume Next

    Dim rsClientes As New ADODB.Recordset, sqlClientes As String
    
    sqlClientes = "SELECT * FROM VistaClientes ORDER BY Codigo ASC"
            
    With rsClientes
        .CursorLocation = adUseClient
        
        Call .Open(sqlClientes, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then
            .MoveFirst
            FormatoGrilla (.RecordCount)
        Else
            FormatoGrilla (1)
        End If
        
        Do Until .EOF = True
            KlexClientes.TextMatrix(.AbsolutePosition, 1) = EsNulo(.Fields("idClientes").Value)
            KlexClientes.TextMatrix(.AbsolutePosition, 2) = "[" & EsNulo(.Fields("Codigo").Value) & "]"
            KlexClientes.TextMatrix(.AbsolutePosition, 3) = EsNulo(.Fields("Nombre").Value)
            KlexClientes.TextMatrix(.AbsolutePosition, 4) = EsNulo(.Fields("Direccion").Value)
            KlexClientes.TextMatrix(.AbsolutePosition, 5) = EsNulo(.Fields("Localidad").Value)
            KlexClientes.TextMatrix(.AbsolutePosition, 6) = EsNulo(.Fields("Telefono").Value)
            KlexClientes.TextMatrix(.AbsolutePosition, 7) = ""
        
            .MoveNext

        Loop

    End With

    sqlClientes = ""
    
    If rsClientes.State = 1 Then
        rsClientes.Close
        Set rsClientes = Nothing
    End If
    
If Err Then GrabarLog "CargarClientesEnGrilla", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Depurar()
On Error Resume Next
Dim vdato As String

vdato = InputBox("Ingresar la palabra de seguridad: ")

If Not vdato = "dalas" Then Exit Sub


    If MsgBox("Está realmente seguro de borrar los datos seleccionados ?", vbYesNo, "Consulta ...") = vbYes Then
        
        MousePointer = vbHourglass
        
        If c(0).Value = 1 Then Call BorrarBase("articulos", pathDBMySQL)
        If c(1).Value = 1 Then Call BorrarBase("Clientes", pathDBMySQL)
        If c(2).Value = 1 Then Call BorrarBase("proveedores", pathDBMySQL)
        
        If c(5).Value = 1 Then
            Call BorrarBase("Factura", pathDBMySQL)
            Call BorrarBase("Fdetalle", pathDBMySQL)
            Call BorrarBase("IvaVenta", pathDBMySQL)
        End If
        
        If c(3).Value = 1 Then
            Call BorrarBase("pfactura", pathDBMySQL)
            Call BorrarBase("pfdetalle", pathDBMySQL)
            Call BorrarBase("ivacompra", pathDBMySQL)
        End If

        If c(4).Value = 1 Then
            Call MigraCtaCteCliente
        End If
        
        If c(6).Value = 1 Then Call BorrarBase("pcuentascorrientes", pathDBMySQL)
        If c(7).Value = 1 Then Call BorrarBase("cheques", pathDBMySQL)
        If c(8).Value = 1 Then Call BorrarBase("bancosmovimientos", pathDBMySQL)
        
        If c(9).Value = 1 Then
            Call BorrarBase("asientos", pathDBMySQL)
            Call BorrarBase("asientosdetalle", pathDBMySQL)
        End If
        
        
        If c(10).Value = 1 Then Call BorrarBase("t_logcaja", pathDBMySQL)
        
        MousePointer = vbDefault
        
        MsgBox "Los datos fueron depurados correctamente.", vbInformation, "Mensaje..."
    
    End If
    
If Err Then GrabarLog "Depurar", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Load()
On Error Resume Next

    With Me
        .Top = 0
        .Left = 0
        .Height = 6950
        .Width = 13500
        .dtpFecha.Value = Date - 90
        .Show
    End With
    
    Call CargarClientesEnGrilla
    
    vnrointerno.Text = UltimoNroInterno2
    
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub KlexClientes_DblClick()
On Error Resume Next

    With KlexClientes
        If .TextMatrix(.Row, 7) = "X" Then
            .TextMatrix(.Row, 7) = ""
        Else
            .TextMatrix(.Row, 7) = "X"
        End If
    End With

If Err Then GrabarLog "KlexClientes_DblClick", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub o_Click()
On Error Resume Next

    Dim i As Integer

    If o.Value Then
        For i = 0 To 10
            c(i).Value = 1
        Next
    Else
        For i = 0 To 10
            c(i).Value = 0
        Next
    End If
    
If Err Then GrabarLog "o_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub MigraCtaCteCliente()
On Error Resume Next
    
    Dim j As Integer, vcliente As String
    
    With KlexClientes
        
        For j = 1 To Val(.Rows - 1)
            
            If .TextMatrix(j, 7) = "X" Then
                
                vcliente = Replace(Replace(EsNulo(.TextMatrix(j, 2)), "[", ""), "]", "")
                
                Call MigrarSaldoCliente(vcliente, EsNulo(.TextMatrix(j, 3)))
                
                .TextMatrix(j, 7) = "Depurado"
                
                vcliente = ""
            
            Else
                
                .TextMatrix(j, 7) = "Omitido"
            
            End If
        
        Next j
    
    End With

If Err Then GrabarLog "MigraCtaCteCliente", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub MigrarSaldoCliente(vCodigoCliente As String, vNombreCliente As String)
On Error Resume Next
    
    Dim rsCtaCteC As New ADODB.Recordset, sqlCtaCteC As String, vSaldoCliente As Double
    
    sqlCtaCteC = "SELECT * FROM cuentascorrientes WHERE (Codigo = '" & vCodigoCliente & "' AND Fecha < '" + strfechaMySQL(dtpFecha.Value) + "')"
    
    With rsCtaCteC
        .CursorLocation = adUseClient
        
        Call .Open(sqlCtaCteC, ConnDDBB, adOpenDynamic, adLockPessimistic)
        
        If Not .RecordCount = 0 Then
            .MoveFirst
            vSaldoCliente = 0
        Else
            vSaldoCliente = 0
            Exit Sub
        End If
        
        Do Until .EOF = True
            vSaldoCliente = vSaldoCliente & Val(Format(.Fields("debito").Value, "#######0.000")) - Val(Format(.Fields("credito").Value, "#######0.000"))
            .MoveNext
        Loop
        
        sqlCtaCteC = ""
    
        If Not rsCtaCteC.State = 1 Then
            rsCtaCteC.Close
            Set rsCtaCteC = Nothing
        End If
            
    End With
    
    'Borro los registros que voy a Agrupar
    Call BorrarBase("cuentascorrientes WHERE (Fecha < '" & strfechaMySQL(dtpFecha.Value) & "') AND (Codigo = '" & vCodigoCliente & "')", pathDBMySQL)
    
    If vSaldoCliente >= 0 Then
        Call EjecutarScript("INSERT INTO CuentasCorrientes (Codigo, Nombre, Fecha, Debito, Credito, Saldo, Comentario) VALUES ('" & vCodigoCliente & "','" & vNombreCliente & "','" & strfechaMySQL(dtpFecha.Value - 1) & "'," & Val(vSaldoCliente) & ", 0,0,'Saldo anterior al " & dtpFecha.Value & "')")
    Else
        Call EjecutarScript("INSERT INTO CuentasCorrientes (Codigo, Nombre, Fecha, Debito, Credito, Saldo, Comentario) VALUES ('" & vCodigoCliente & "','" & vNombreCliente & "','" & strfechaMySQL(dtpFecha.Value - 1) & "',0," & Val(Format(vSaldoCliente * (-1), "######0.000")) & ", 0,'Saldo anterior al " & dtpFecha.Value & "')")
    End If

If Err Then GrabarLog "MigrarSaldoCliente", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub PbAcciones_Click(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 0
            Depurar
            
        Case 1
            Unload Me
        
    End Select

If Err Then GrabarLog "PbAcciones_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub PBSeleccionar_Click(Index As Integer)
On Error Resume Next

    Dim m As Integer
    
    With KlexClientes
        For m = 1 To (.Rows - 1)
        
            Select Case Index
    
                Case 0
                    .TextMatrix(m, 7) = "X"
                Case 1
                    .TextMatrix(m, 7) = ""
            
            End Select
        Next
    End With

If Err Then GrabarLog "PBSeleccionar_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub PusEjecutar_Click()
On Error Resume Next
Dim vsql As String
Dim i As Integer

For i = Val(UltimoNroInterno2) To vnrointerno
    vsql = "insert into t_nrointerno (auxiliar) values (1)"
    Call EjecutarScript(vsql, pathDBMySQL)
Next

MsgBox "Trabajo completado"

If Err Then Exit Sub
End Sub
