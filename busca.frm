VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmBuscarArticulo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Búsqueda de Artículos"
   ClientHeight    =   4590
   ClientLeft      =   2040
   ClientTop       =   2280
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   10440
      TabIndex        =   7
      Top             =   4305
      Width           =   10500
   End
   Begin VB.TextBox txtProveedor 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2370
      TabIndex        =   5
      Top             =   3930
      Width           =   6975
   End
   Begin VB.CommandButton cmdSeleccionar 
      Caption         =   "Select"
      Height          =   285
      Left            =   9510
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Volver al módulo anterior"
      Top             =   3600
      Width           =   945
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   315
      Left            =   9510
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3930
      Visible         =   0   'False
      Width           =   945
   End
   Begin MSDataGridLib.DataGrid dgArticulos 
      Bindings        =   "busca.frx":0000
      Height          =   3375
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   16777215
      HeadLines       =   2
      RowHeight       =   15
      RowDividerStyle =   4
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      ColumnCount     =   21
      BeginProperty Column00 
         DataField       =   "Codigo"
         Caption         =   "Codigo"
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
         DataField       =   "Descrip"
         Caption         =   "Descrip"
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
         DataField       =   "Rubro"
         Caption         =   "Rubro"
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
         DataField       =   "Stock"
         Caption         =   "Stock"
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
         DataField       =   "Pcosto"
         Caption         =   "Pcosto"
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
         DataField       =   "Pventa1"
         Caption         =   "Pventa1"
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
         DataField       =   "Pventa2"
         Caption         =   "Pventa2"
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
         DataField       =   "Pventa3"
         Caption         =   "Pventa3"
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
         DataField       =   "Pventa4"
         Caption         =   "Pventa4"
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
         DataField       =   "Pventa5"
         Caption         =   "Pventa5"
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
         DataField       =   "Pventa6"
         Caption         =   "Pventa6"
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
         DataField       =   "Faltante"
         Caption         =   "Faltante"
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
         DataField       =   "Ganancia"
         Caption         =   "Ganancia"
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
         DataField       =   "Proveedor"
         Caption         =   "Proveedor"
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
         DataField       =   "Envase"
         Caption         =   "Envase"
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
      BeginProperty Column15 
         DataField       =   "id"
         Caption         =   "id"
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
      BeginProperty Column16 
         DataField       =   "Ganancia_Vendedor"
         Caption         =   "Ganancia_Vendedor"
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
      BeginProperty Column17 
         DataField       =   "Rubro_descrip"
         Caption         =   "Rubro_descrip"
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
      BeginProperty Column18 
         DataField       =   "codigo_num"
         Caption         =   "codigo_num"
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
      BeginProperty Column19 
         DataField       =   "pventat"
         Caption         =   "pventat"
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
      BeginProperty Column20 
         DataField       =   "porcentaje"
         Caption         =   "porcentaje"
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
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
         EndProperty
         BeginProperty Column12 
         EndProperty
         BeginProperty Column13 
         EndProperty
         BeginProperty Column14 
         EndProperty
         BeginProperty Column15 
         EndProperty
         BeginProperty Column16 
         EndProperty
         BeginProperty Column17 
         EndProperty
         BeginProperty Column18 
         EndProperty
         BeginProperty Column19 
         EndProperty
         BeginProperty Column20 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtArticulo 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2370
      TabIndex        =   0
      Top             =   3570
      Width           =   7005
   End
   Begin MSAdodcLib.Adodc barticulos_clientes 
      Height          =   345
      Left            =   5160
      Top             =   3000
      Visible         =   0   'False
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   609
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "barticulos_clientes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc barticulo 
      Height          =   345
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   609
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "barticulo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblDatos 
      Appearance      =   0  'Flat
      Caption         =   "Proveedor :"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   6
      Top             =   3960
      Width           =   885
   End
   Begin VB.Label lblDatos 
      Appearance      =   0  'Flat
      Caption         =   "Escriba código o descripción :"
      ForeColor       =   &H00000000&
      Height          =   165
      Index           =   0
      Left            =   150
      TabIndex        =   4
      Top             =   3630
      Width           =   2205
   End
End
Attribute VB_Name = "frmBuscarArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Buscar As String
Public o As Integer
Public vlista As Integer

Public busca As Integer '1 Codigo, 2 Descrip
Dim vsql As String
Private Sub cmdSalir_Click()
On Error Resume Next
    
    Unload Me

If Err Then GrabarLog "cmdSalir_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Seleccionar()
On Error Resume Next

    With barticulo
        If .Recordset.EOF Then cmdSalir_Click
    
        Select Case o
            Case 0
                With frmArticulosAlta
                    .ModificarArticulo (barticulo.Recordset("idArticulos").Value)
                    Unload Me
                End With
                
            Case 1
                With barticulos_clientes
                    If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
                    .RecordSource = "SELECT * FROM articulos_clientes WHERE (codigo_cliente = '" & Trim(frmRemito.txtCliente(0).Tag) & "') AND (Articulo = '" & Trim(barticulo.Recordset("codigo").Value) & "')"
                    .Refresh
    
                    If Not .Recordset.EOF Then
                        frmRemito.txtDetalle(2).Text = .Recordset(4).Value
                        frmRemito.txtDetalle(2).Text = .Recordset(5).Value
                        frmRemito.ElegirTipoPrecio
                        frmBuscarArticulo.Visible = False
                        frmRemito.txtDetalle(2).SetFocus
                    Else
                        frmRemito.txtDetalle(1).Text = barticulo.Recordset("Codigo").Value
                        frmRemito.txtDetalle(1).Text = barticulo.Recordset("descrip").Value
                        frmRemito.vvcodigo = barticulo.Recordset("codigo").Value
                        frmRemito.txtDetalle(2).Text = barticulo.Recordset("pventa" + Trim(Me.vlista)).Value
                        frmRemito.ElegirTipoPrecio
                        frmBuscarArticulo.Visible = False
                        frmRemito.txtDetalle(2).SetFocus
                    End If
                End With
            Case 9

                With barticulo
                    frmCompras.vvdescrip = .Recordset("descrip").Value
                    frmCompras.vvcodigo = .Recordset("codigo").Value
                    frmCompras.vvpventa = .Recordset("pventa1").Value
                End With

                frmCompras.CargarBien
                frmBuscarArticulo.Visible = False
                frmCompras.f(2).SetFocus
            
            Case 10
                'With frmAddCargacamion
                '    .txtCodigo.Text = barticulo.Recordset("codigo").value
                '    .txtArticulo.Text = barticulo.Recordset("descrip").value
                '    .txtCantidad.SetFocus
                '    Unload Me
                'End With
    
            Case 11
                'With frmEstadisticaProducto
                '    .txtArticulo.Text = barticulo.Recordset(0).value
                '    .txtarticulo_keypress 13
                '    .WindowState = vmaximizar
                '    Unload Me
                'End With
        End Select

    End With

If Err Then GrabarLog "cmdSeleccionar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdSeleccionar_Click()
On Error Resume Next
    
    Seleccionar

If Err Then GrabarLog "cmdSeleccionar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Load()
On Error Resume Next

    With barticulos_clientes
        .ConnectionString = pathDBMySQL
        .RecordSource = "Articulos_Clientes"
        .Refresh
    End With

    With barticulo
        .ConnectionString = pathDBMySQL
        .RecordSource = "Articulos"
        .Refresh
    End With
    
    With Me
        .Top = 1400
        .Left = 500
        .Width = 10590
        .Height = 4920
    End With
     Call CentrarFormulario(Me)
If Err Then GrabarLog "cmdSeleccionar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub dgArticulos_DblClick()
On Error Resume Next

    Seleccionar

If Err Then GrabarLog "dgArticulos_DblClick", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub dgArticulos_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next

    Call OrdenarDataGrid(ColIndex, barticulo.Recordset, dgArticulos)

If Err Then GrabarLog "dgArticulos_HeadClick", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub txtarticulo_keypress(Keyascii As Integer)
On Error Resume Next

    If Keyascii = 27 Then Visible = False

    vsql = ""

    MousePointer = vbHourglass

    If Keyascii = 13 Then
        If Not txtArticulo.Text = "" Then vsql = vsql + " and (codigo like '%" + txtArticulo.Text + "%' or descrip like '%" + txtArticulo.Text + "%')"
        If Not txtProveedor.Text = "" Then vsql = vsql + " and proveedor like '%" + txtProveedor.Text + "%'"
  
        With barticulo
            If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
            .RecordSource = "SELECT * FROM articulos WHERE 1=1 " & vsql
            .Refresh
        End With
        
        txtProveedor.SetFocus

    End If

    MousePointer = vbDefault
    
If Err Then GrabarLog "txtArticulo_keypress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtProveedor_KeyPress(Keyascii As Integer)
On Error Resume Next

    If Keyascii = 27 Then Me.Visible = False
    

    vsql = ""

    MousePointer = vbHourglass

    If Keyascii = 13 Then
        If Not txtArticulo.Text = "" Then vsql = vsql + "and (codigo like '%" + txtArticulo.Text + "%' or descrip like '%" + txtArticulo.Text + "%')"
        If Not txtProveedor.Text = "" Then vsql = vsql + "and proveedor like '%" + txtProveedor.Text + "%'"
  
        With barticulo
            If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
            .RecordSource = "SELECT * FROM articulos WHERE 1=1 (" & vsql & ")"
            .Refresh
        End With
    
        cmdSeleccionar.SetFocus

    End If

    MousePointer = vbDefault
    
If Err Then GrabarLog "txtProveedor_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
