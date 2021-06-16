VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmControlFacturacion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Control de Facturación"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   13260
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab TabControlFacturacion 
      Height          =   5175
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Control gral de Detalles de Documentos de Ventas"
      TabPicture(0)   =   "frmControlFacturacion.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Dgfdetalle"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin MSDataGridLib.DataGrid Dgfdetalle 
         Bindings        =   "frmControlFacturacion.frx":001C
         Height          =   4635
         Left            =   45
         TabIndex        =   4
         Top             =   480
         Width           =   11805
         _ExtentX        =   20823
         _ExtentY        =   8176
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483624
         Enabled         =   -1  'True
         HeadLines       =   2
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   15
         BeginProperty Column00 
            DataField       =   "Fecha"
            Caption         =   "Fecha"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "CodCli"
            Caption         =   "Cod."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Nombre"
            Caption         =   "Nombre"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Codigo"
            Caption         =   "Cod. Art."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Descripcion"
            Caption         =   "Articulo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Cantidad"
            Caption         =   "Vendido"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Impuesto"
            Caption         =   "Devol."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "Repartidor"
            Caption         =   "Repar."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "confirmado"
            Caption         =   "confirmado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "Domicilio"
            Caption         =   "Domicilio"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "Remito"
            Caption         =   "Remito"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "Cventa"
            Caption         =   "C. Venta"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column12 
            DataField       =   "total_cdo"
            Caption         =   "total_cdo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column13 
            DataField       =   "Total_ctacte"
            Caption         =   "Total_ctacte"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column14 
            DataField       =   "CodRep"
            Caption         =   "fdetalle.repartidor"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1934.929
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1904.882
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1604.976
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1590.236
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   14.74
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1500
      Left            =   0
      TabIndex        =   5
      Top             =   5280
      Width           =   11895
      Begin VB.Frame Frame2 
         Caption         =   "Ordenar"
         Height          =   735
         Left            =   8970
         TabIndex        =   18
         Top             =   0
         Width           =   2925
         Begin VB.ComboBox cboOrdenar 
            Height          =   315
            ItemData        =   "frmControlFacturacion.frx":003C
            Left            =   240
            List            =   "frmControlFacturacion.frx":003E
            TabIndex        =   19
            Text            =   "Fecha"
            Top             =   280
            Width           =   2415
         End
      End
      Begin XtremeSuiteControls.PushButton PBAcciones 
         Height          =   495
         Index           =   0
         Left            =   9120
         TabIndex        =   16
         Top             =   840
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "&Buscar!"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmControlFacturacion.frx":0040
      End
      Begin VB.TextBox txtEmpleado 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   3015
      End
      Begin VB.CheckBox chkFechas 
         BackColor       =   &H8000000C&
         Caption         =   "Anular Fechas"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   4320
         TabIndex        =   11
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox txtcliente 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txtarticulo 
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   960
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker vfdesde 
         Height          =   315
         Left            =   5640
         TabIndex        =   9
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   64552961
         CurrentDate     =   39356
      End
      Begin MSComCtl2.DTPicker vfhasta 
         Height          =   315
         Left            =   5640
         TabIndex        =   10
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   64552961
         CurrentDate     =   39379
      End
      Begin VB.CheckBox chkCVenta 
         Caption         =   "Contado"
         Height          =   255
         Index           =   0
         Left            =   7440
         TabIndex        =   15
         Top             =   240
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.CheckBox chkCVenta 
         Caption         =   "Cuenta Corriente"
         Height          =   255
         Index           =   1
         Left            =   7440
         TabIndex        =   14
         Top             =   480
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin XtremeSuiteControls.PushButton PBAcciones 
         Height          =   495
         Index           =   1
         Left            =   10440
         TabIndex        =   17
         Top             =   840
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "&Imprimir!"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmControlFacturacion.frx":0512
      End
      Begin VB.Label lblFdesde 
         AutoSize        =   -1  'True
         Caption         =   "> Fecha Desde :"
         Height          =   195
         Left            =   4320
         TabIndex        =   13
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label lblFHasta 
         AutoSize        =   -1  'True
         Caption         =   "> Fecha Hasta :"
         Height          =   195
         Left            =   4320
         TabIndex        =   12
         Top             =   600
         Width           =   1140
      End
      Begin VB.Label lblCliente 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "> Cliente :"
         Height          =   195
         Left            =   30
         TabIndex        =   8
         Top             =   640
         Width           =   1000
      End
      Begin VB.Label lblArticulo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "> Articulo :"
         Height          =   195
         Left            =   30
         TabIndex        =   7
         Top             =   1000
         Width           =   1000
      End
      Begin VB.Label lblRepartidor 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "> Empleado :"
         Height          =   195
         Left            =   105
         TabIndex        =   6
         Top             =   285
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmControlFacturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsFacturaDetalle As ADODB.Recordset
Dim vSQL As String
Dim vMostrarFiltro As String
Private Sub Buscar()
On Error Resume Next
    
    MousePointer = vbHourglass
    
    Set rsFacturaDetalle = New ADODB.Recordset
    Dim sqlFacturaDetalle As String
    'Dim vMostrarTipoVenta As String
    
    Dim i As Integer
    
    vSQL = ""
    vMostrarFiltro = ""
    
    If Not Trim(txtEmpleado.Text) = "" Then
        vSQL = vSQL + " AND (CodRep = '" + Trim(txtEmpleado.Tag) + "')"
        vMostrarFiltro = vMostrarFiltro & "Empleado: " & txtEmpleado.Text
    End If
    
    If chkFechas.Value = 0 Then
        vSQL = vSQL + " AND (fecha >= '" & strfechaMySQL(vfdesde.Value) + "') and (fecha <= '" & strfechaMySQL(vfhasta.Value) + "')"
        vMostrarFiltro = vMostrarFiltro & " / Rango de Fecha: " & vfdesde.Value & " Hasta: " & vfhasta.Value
    End If
    
    If Not txtCliente.Text = "" Then
        vSQL = vSQL + " AND (CodCli = '" + txtCliente.Tag + "')"
        vMostrarFiltro = vMostrarFiltro & " / Cliente: " & txtCliente.Text
    End If
    
    If Not txtarticulo.Text = "" Then
        vSQL = vSQL + " AND (Codigo = '" + txtarticulo.Tag + "')"
        vMostrarFiltro = vMostrarFiltro & " / Articulo: " & txtarticulo.Text
    End If
    
    For i = 0 To 1
        If (chkCVenta(0).Value = 1) And (chkCVenta(1).Value = 1) Then
            vMostrarFiltro = vMostrarFiltro & " / Tipo Venta: " & chkCVenta(0).Caption & " - " & chkCVenta(1).Caption
            Exit For
        Else
            If chkCVenta(i).Value = 1 Then
                vSQL = vSQL + " and (CVenta = '" + chkCVenta(i).Caption + "')"
                vMostrarFiltro = vMostrarFiltro & " / Tipo Venta: " & chkCVenta(i).Caption
            End If
        End If
    Next
    
    sqlFacturaDetalle = "SELECT * FROM factura_fdetalle WHERE 1=1" + vSQL + " ORDER BY " & Ordenar(cboOrdenar)
    
    With rsFacturaDetalle
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        
        Call .Open(sqlFacturaDetalle, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then
            .MoveLast
            Set Dgfdetalle.DataSource = rsFacturaDetalle
        Else
            Set Dgfdetalle.DataSource = Nothing
        End If
    End With
    
    MousePointer = vbDefault

If Err Then GrabarLog "Buscar", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub chkFechas_Click()
On Error Resume Next
    
    vfdesde.Enabled = CBool(chkFechas.Value - 1)
    vfhasta.Enabled = CBool(chkFechas.Value - 1)
        
If Err Then GrabarLog "chkFechas_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Imprimir()
    On Error Resume Next
    
    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox "   Prepare la Impresora   ", vbInformation, "Mensaje ..."
    
    With Mantenimiento.rsFactura_fdetalle
        If Not .State = 0 Then .Close
        
        .Source = rsFacturaDetalle.Source
        
        If Not .State = 1 Then .Open
        .Close
        .Open
    End With
    
    With drcontrol_facturacion
        .Sections("PieInforme").Controls("lblFiltro").Caption = vMostrarFiltro
        .Show
    End With
    
    If Err Then
        MsgBox "Los datos del reporte generado no son correctos debido" & vbCrLf & " a entradas no Validas, verifiquelas y vuelva a generarlo.", vbInformation, "Mensaje ..."
        GrabarLog "cmdImprimir_Click", Err.Number & " " & Err.Description, Me.Name
    End If
End Sub
Private Sub DgFdetalle_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next

    OrdenarDataGrid ColIndex, rsFacturaDetalle, Dgfdetalle

If Err Then GrabarLog "Dgfdetalle_HeadClick", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Function Ordenar(vOrdena As String) As String
On Error Resume Next

    Select Case vOrdena

        Case "Fecha"
            Ordenar = "Fecha"

        Case "Codigo"
            Ordenar = "CodCli"

        Case "Nombre"
            Ordenar = "Nombre"

        Case "C. articulo"
            Ordenar = "Codigo"

        Case "Articulo"
            Ordenar = "Descripcion"

        Case "Repartidor"
            Ordenar = "CodRep"

        Case "Confirmado"
            Ordenar = "Confirmado"
    
    End Select
    
If Err Then GrabarLog "Ordenar", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub Form_Load()
On Error Resume Next

    Width = 12045
    Height = 7890
    
    vfdesde.Value = Date
    vfhasta.Value = Date
    
    With cboOrdenar
        .AddItem "Fecha"
        .AddItem "Nombre"
        .AddItem "C. articulo"
        .AddItem "Articulo"
        .AddItem "Repartidor"
        .AddItem "Confirmado"
    End With
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)
    
    If rsFacturaDetalle.State = 1 Then
        rsFacturaDetalle.Close
        Set rsFacturaDetalle = Nothing
    End If

If Err Then GrabarLog "", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub PBAcciones_Click(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 0
            Buscar
        
        Case 1
            Imprimir
        
        Case 2
            
    End Select

If Err Then GrabarLog "PBAcciones_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub txtarticulo_keypress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
    
        Dim rsArticulos As New ADODB.Recordset, sqlArticulos As String

        sqlArticulos = "SELECT * FROM articulos WHERE (Codigo = '" & Trim(txtarticulo.Text) & "') OR (Descrip LIKE '%" & Trim(txtarticulo.Text) & "%')"
        
        With rsArticulos
            .CursorLocation = adUseClient
            Call .Open(sqlArticulos, ConnDDBB, adOpenStatic, adLockReadOnly)
            
            If Not .EOF = True Then
                txtarticulo.Text = Trim(.Fields("Descrip").Value)
                txtarticulo.Tag = Trim(.Fields("Codigo").Value)
            Else
                txtarticulo.Text = ""
                txtarticulo.Tag = ""
            End If

        End With
        
        sqlArticulos = ""
        
        If rsArticulos.State = 1 Then
            rsArticulos.Close
            Set rsArticulos = Nothing
        End If

    End If

If Err Then GrabarLog "txtarticulo_keypress", Err.Number & " " & Err.Description, Caption
End Sub
Public Sub txtCliente_Keypress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        
        Dim rsClientes As New ADODB.Recordset, sqlClientes As String
        
        sqlClientes = "SELECT * FROM clientes WHERE (Codigo = '" & Trim(txtCliente.Text) & "') OR (Nombre LIKE '%" & Trim(txtCliente.Text) & "%')"
        
        With rsClientes
            .CursorLocation = adUseClient
            Call .Open(sqlClientes, ConnDDBB, adOpenStatic, adLockReadOnly)

            If Not .EOF = True Then
                txtCliente.Text = .Fields("Nombre").Value
                txtCliente.Tag = .Fields("Codigo").Value
                txtarticulo.SetFocus
            Else
                With frmBuscarCliente
                    .Show
                    .o = 2
                    .txtClientes.Text = txtCliente.Text
                End With
            End If

        End With
        
        sqlClientes = ""
        
        If rsClientes.State = 1 Then
            rsClientes.Close
            Set rsClientes = Nothing
        End If

    End If
    
If Err Then GrabarLog "txtcliente_Keypress", Err.Number & " " & Err.Description, Caption
End Sub
Private Sub txtEmpleado_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        
        Dim rsempleados As New ADODB.Recordset, sqlEmpleados As String
        
        sqlEmpleados = "SELECT * FROM Empleados WHERE (Codigo = '" + Trim(txtEmpleado.Text) + "')"
        
        With rsempleados
            .CursorLocation = adUseClient
            Call .Open(sqlEmpleados, ConnDDBB, adOpenStatic, adLockReadOnly)
            
            If Not .EOF = True Then
                txtEmpleado.Text = .Fields("Nombre").Value
                txtEmpleado.Tag = .Fields("Codigo").Value
                txtCliente.SetFocus
            Else
                txtEmpleado.Text = ""
                txtEmpleado.Tag = ""
            End If

        End With
        
        sqlEmpleados = ""
        
        If rsempleados.State = 1 Then
            rsempleados.Close
            Set rsempleados = Nothing
        End If

    End If
    
If Err Then GrabarLog "txtEmpleado_KeyPress", Err.Number & " " & Err.Description, Caption
End Sub
