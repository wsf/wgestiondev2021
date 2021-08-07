VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#13.0#0"; "Codejock.CommandBars.v13.0.0.Demo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmCuentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alta & Mantenimiento de Cuentas"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   13995
   Begin VB.PictureBox PicInferior 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   5640
      Picture         =   "frmCuentas.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   5505
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   5505
   End
   Begin XtremeSuiteControls.TabControl TabCuentas 
      Height          =   7935
      Left            =   30
      TabIndex        =   1
      Top             =   600
      Width           =   13965
      _Version        =   851968
      _ExtentX        =   24633
      _ExtentY        =   13996
      _StockProps     =   68
      ItemCount       =   2
      Item(0).Caption =   "Vista : Listado"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "dgGeneral"
      Item(0).Control(1)=   "Frame1"
      Item(1).Caption =   "Vista: Arbol"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "TVCuentas"
      Begin MSDataGridLib.DataGrid dgGeneral 
         Height          =   6810
         Left            =   90
         TabIndex        =   2
         Top             =   990
         Width           =   13725
         _ExtentX        =   24209
         _ExtentY        =   12012
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483628
         BorderStyle     =   0
         HeadLines       =   2
         RowHeight       =   15
         RowDividerStyle =   3
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0-00-000-0000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   60
         TabIndex        =   5
         Top             =   330
         Width           =   8655
         Begin XtremeSuiteControls.GroupBox GroupBox1 
            Height          =   405
            Left            =   -1920
            TabIndex        =   8
            Top             =   390
            Width           =   11655
            _Version        =   851968
            _ExtentX        =   20558
            _ExtentY        =   714
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            BorderStyle     =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtBuscar 
            Height          =   285
            Left            =   1050
            TabIndex        =   6
            Top             =   90
            Width           =   7575
            _Version        =   851968
            _ExtentX        =   13361
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.Label lblBuscar 
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   60
            Width           =   975
            _Version        =   851968
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Buscar :"
            Transparent     =   -1  'True
         End
      End
      Begin MSComctlLib.TreeView TVCuentas 
         Height          =   6000
         Left            =   -69880
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   8500
         _ExtentX        =   15002
         _ExtentY        =   10583
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ImageList2"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2400
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCuentas.frx":50B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCuentas.frx":598D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCuentas.frx":5E3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCuentas.frx":62DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCuentas.frx":67A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCuentas.frx":6CF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCuentas.frx":71C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCuentas.frx":76CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCuentas.frx":7B84
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCuentas.frx":80BF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.ProgressBar Barra 
      Height          =   210
      Left            =   60
      TabIndex        =   0
      Top             =   8550
      Width           =   13755
      _Version        =   851968
      _ExtentX        =   24262
      _ExtentY        =   370
      _StockProps     =   93
      Scrolling       =   1
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   0
      Top             =   0
      _Version        =   851968
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   3
   End
   Begin XtremeCommandBars.ImageManager ImageManager1 
      Left            =   480
      Top             =   0
      _Version        =   851968
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmCuentas.frx":85D0
   End
End
Attribute VB_Name = "frmCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCuentas As ADODB.Recordset
Dim sqlBusqueda As String
Dim vIDCuenta As Long
Dim tNodo As Node ' sacado
Public Sub Buscar()
    On Error Resume Next
        
    sqlBusqueda = ""

    'If Not (Val(txtCodigoDesde.Text) = 0) And Not (Val(txtCodigoHasta.Text) = 0) Then
    '    sql = sql + " AND ((CodigoCuenta >= '" & Val(txtCodigoDesde.Text) & "') AND (CodigoCuenta <= '" & Val(txtCodigoHasta.Text) & "'))"
    'End If
    
    If Not Trim(txtBuscar.Text) = "" Then sqlBusqueda = sqlBusqueda + " AND (Cuenta LIKE '%" + Trim(txtBuscar.Text) + "%' or CodigoCuenta like '%" + Trim(txtBuscar.Text) + "%')"
    
    'If Not RBImputable(0).Value = True Then
    '    If RBImputable(1).Value = True Then
    '        sql = " AND (Imputable = 'S')"
    '    Else
    '        sql = " AND (Imputable = 'N')"
    '    End If
    'End If
    
    Set rsCuentas = New ADODB.Recordset
    Dim sqlCuentas As String
    
    sqlCuentas = "SELECT *  FROM Cuentas WHERE (1=1) " & sqlBusqueda & " order by CodigoCuenta"
    
    With rsCuentas
        If .State = 1 Then .Close
        
        Call .Open(sqlCuentas, ConnDDBB, adOpenStatic, adLockReadOnly)

        If Not .EOF = True Then .MoveLast
    End With
    
    FormatoGrilla
    
If Err Then GrabarLog "Buscar", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Imprimir(Index As Integer)
    On Error Resume Next

    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox "Prepare la Impresora!!!", vbInformation, "Mensaje ..."
    
    Select Case Index
    
        Case 2
            With Mantenimiento.rsCuentasContables
                If .State = 1 Then .Close
        
                .Source = rsCuentas.Source
        
                If Not .State = 1 Then .Open
                .Close
                .Open
            End With
    Mantenimiento.rsCuentasContables.Sort = "CodigoCuenta"
    
            With drCuentas
                .Sections(2).Controls("snombre").Caption = vDatosEmpresa.Nombre
                .Sections(2).Controls("sdirtel").Caption = vDatosEmpresa.Direccion & "  /  " & vDatosEmpresa.Telefono
                .Sections(2).Controls("slocalidad").Caption = vDatosEmpresa.Localidad
                .Sections(2).Controls("semail").Caption = vDatosEmpresa.Email
                
                .Show
        
            End With
        
        Case 5
    
    End Select
    


    If Err Then GrabarLog "Imprimir", Err.Number & " " & Err.Description, Me.Name
    
End Sub
Private Sub ArmarArbolDeCuentas()
On Error Resume Next

    TVCuentas.Enabled = Not True
    InicializarLlenadoArbol
    RecorrerCuentas
    'cmdAcciones(4).Enabled = True
    'cmdAcciones(5).Enabled = True
    TVCuentas.Enabled = True
            
            
Exit Sub
TVCuentas.Enabled = Not True
CalcularSaldoPorNivel (0)
TVCuentas.Enabled = True
    
If Err Then GrabarLog "ArmarArbolDeCuentas", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdCuentas_Click(Index As Integer)
On Error Resume Next

        ' Sustituir el texto del nodo seleccionado
        ' por el contenido del cboConcepto
        
        'TVCuentas.SelectedItem.Text = txtCodigo(1).Text

        'Dim rsAgregar As New ADODB.Recordset, sqlAgregar As String
        
        'sqlAgregar = "SELECT * FROM Cuentas WHERE (CodigoCuenta = '" & Mid(tNodo.Key, 3, ContarCaracteres(tNodo.Key, "-") - 2) & "')"

        'With rsAgregar
        '    If .State = 1 Then .Close
        '    .CursorLocation = adUseClient
            
        '    Call .Open(sqlAgregar, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        '    If .RecordCount = 1 Then
        '        .MoveFirst
                
        '        .Fields("CodigoCuenta").Value = Left(Trim(txtCodigo(0).Text), 10)
        '        .Fields("Cuenta").Value = Left(Trim(txtCodigo(1).Text), 255)
                        
        '        If chkImputable.Value = xtpChecked Then
        '            .Fields("Imputable").Value = "N"
        '        Else
        '            .Fields("Imputable").Value = "S"
        '        End If
                        
        '        .Update
        '    End If
        'End With

If Err Then GrabarLog "", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub BorrarNodo()
    Dim tNodo As Node
    Dim i As Long
    
    ' El nodo que está actualmente seleccionado
    Set tNodo = TVCuentas.SelectedItem
    i = tNodo.children
    Dim padre As String
    padre = tNodo.Parent.key
    
    ' Avisar que se va a borrar un nodo que tiene hijos
    If i > 0 Then
        If MsgBox("¿Esta seguro que quiere borrar el concepto con " & CStr(i) & " subconceptos?", vbQuestion Or vbYesNo, "Borrar nodos") = vbNo Then
            Exit Sub
        End If
    End If
    TVCuentas.Nodes.Remove tNodo.Index
    Dim padreAActualizar As Long
    If i = 0 Then
        
        'padreAActualizar = EliminarUnico(tNodo.Key)
        
    Else
        'EliminarConcepto Mid(tNodo.Key, 3, ContarCaracteres(tNodo.Key, "-") - 2), True
        'padreAActualizar = EliminarUnico(tNodo.Key)
    End If
    
    'Actualizar tieneconcepto del padre
    'ActualizarTieneSubconceptoEnEliminacion (padreAActualizar)
End Sub
Private Sub CommandBars_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error Resume Next

    Select Case control.Index
                
        Case 1
            With frmCuentasAltas
                .Show
                .vaccion = "Nuevo"
                .Caption = "Agregar Cuenta"
                .txtAlta(0).SetFocus
            End With
            
        Case 2
            If Not (rsCuentas.EOF = True) And Not (rsCuentas.BOF = True) Then
                With frmCuentasAltas
                    .Show
                    .ModificarCuenta (rsCuentas.Fields("idCuentas").Value)
                    .vaccion = "Modificar"
                    .Caption = "Modificar Cuenta"
                End With
            Else
                MsgBox "Debe elegir una cuenta para poder modificarla!!!", vbExclamation, "Mensaje ..."
            End If
            
        Case 3
            'Duplicar
        
        Case 4
            With rsCuentas
                If Not (.EOF = True) And Not (.BOF = True) Then
                    If MsgBox("Esta seguro que desea borrar este registro?", vbInformation + vbYesNo, "Mensaje ...") = vbYes Then
                        Call BorrarBase(vConfigGral.vempresa & ".Cuentas WHERE (idCuentas = " & .Fields("idCuentas").Value & ")", pathDBMySQL)
                        'Call BorrarBase(vConfigGral.vEmpresa & ".Articulosclientes WHERE (CodigoCliente = " & .Fields("Codigo").Value & ")", pathDBMySQL)
                        'Call BorrarBase(vConfigGral.vEmpresa & ".FacturaAutomatica WHERE (CodigoCliente = " & .Fields("Codigo").Value & ")", pathDBMySQL)
                        
                        Buscar
                    End If
                End If
            End With

        
        Case 5
            Buscar
                
        Case 6
            
            Call Imprimir(2)
            
            'vVieneImpresion = Me.Name
            'frmImprimir.Show
            
        Case 7
            ArmarArbolDeCuentas
            
        Case 8
            Unload Me
            
    End Select

If Err Then GrabarLog "CommandBars_Execute", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub dgGeneral_DblClick()
On Error Resume Next
    
    If Not (rsCuentas.EOF = True) And Not (rsCuentas.BOF = True) Then
        With frmCuentasAltas
            .Show
            .ModificarCuenta (rsCuentas.Fields("idCuentas").Value)
            .vaccion = "Modificar"
            .Caption = "Modificar Cuenta"
        End With
    Else
        MsgBox "Debe elegir una cuenta para poder modificarla!!!", vbExclamation, "Mensaje ..."
    End If

    
    Exit Sub
    
    With rsCuentas
        If Not (.EOF = True) And Not (.BOF = True) Then
            'Nuevo
            'txtCodigo(0).Text = .Fields("Codigo").Value
            'txtCodigo(1).Text = .Fields("Cuenta").Value
            
            'If .Fields("Imputable").Value = "N" Then
                
            '    chkImputable.Value = xtpChecked
            
            'Else
            
            '    chkImputable.Value = xtpUnchecked
            
           ' End If
            
            'vIDCuenta = .Fields("idCuentas").Value
            
            'TabBusqueda.SelectedItem = 0
        'Else
        
        End If
    End With
    
If Err Then GrabarLog "DgGeneral_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub dgGeneral_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next
    
    Call OrdenarDataGrid(ColIndex, rsCuentas, dgGeneral)
    
If Err Then GrabarLog "DgGeneral_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub FormatoGrilla()
    On Error Resume Next
    
    With dgGeneral
        
        Set .DataSource = rsCuentas
        
        .Columns(0).Width = 0
        .Columns(0).Locked = True
        .Columns(1).Width = 2000
        .Columns(1).Caption = "Codigo"
        .Columns(2).Width = 7000
        .Columns(2).Caption = "Nombre de la Cuenta"
        .Columns(3).Width = 0
        .Columns(4).Width = 950
        .Columns(4).Alignment = dbgRight
    End With
    
    If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Load()
    On Error Resume Next
        
    CargarBotonera

    With Me
        .Show
        '.Top = 0
        '.Left = 0
        '.Width = 9000
        '.Height = 7850 ' + 350
        '.PicInferior.Top = -45
        '.PicInferior.Left = 4025
    End With
   
    Call Buscar

    CentrarFormulario Me

    If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CargarBotonera()
On Error Resume Next
    
    CommandBarsGlobalSettings.App = App
     
    Dim control As CommandBarControl
    Dim ToolBar As CommandBar
    Set ToolBar = CommandBars.Add("Standard", xtpBarTop)
    
    AddControl ToolBar.Controls, xtpControlButton, 2, "&Nuevo", False, "Crea una Nueva Cuenta"
    AddControl ToolBar.Controls, xtpControlButton, 11, "&Modificar", False, "Modifica una Cuenta Seleccionada"
    AddControl ToolBar.Controls, xtpControlButton, 5, "&Duplicar", False, ""
    AddControl ToolBar.Controls, xtpControlButton, 6, "&Borrar", False, "Borra una Cuenta Seleccionada"
    ToolBar.Closeable = True
    AddControl ToolBar.Controls, xtpControlButton, 14, "Bu&scar", True, ""
    AddControl ToolBar.Controls, xtpControlButton, 26, "&Imprimir", True, ""
    AddControl ToolBar.Controls, xtpControlButton, 19, "&Armar Arbol", True, "Arma el Arbol de Niveles de Cuentas"
    ToolBar.Closeable = True
    AddControl ToolBar.Controls, xtpControlButton, 16, "&Salir", False, ""
    
    ToolBar.CommandBars.VisualTheme = xtpThemeVisualStudio2008
    
    CommandBars.Options.LargeIcons = True
    
    Call CommandBars.DockToolBar(ToolBar, 0, 0, xtpBarTop)
    
    'Disable MenuBar Docking
    CommandBars.ActiveMenuBar.EnableDocking xtpFlagStretched
    
    'Disable ToolBar Docking
    ToolBar.EnableDocking xtpFlagHideWrap
    CommandBars.ActiveMenuBar.ShowGripper = False
    
    Set CommandBars.Icons = ImageManager1.Icons
    CommandBars.Options.UseDisabledIcons = True
    'UseDisabledIcons = True
    CommandBars.Options.SetIconSize True, 24, 24
    CommandBars.Options.ShowExpandButtonAlways = False
    
If Err Then GrabarLog "CargarBotonera", Err.Number & "-" & Err.Description, Me.Name
End Sub
Private Sub RecorrerCuentas()
On Error Resume Next

    Dim nodoPadre As Node
    
    With TVCuentas
        .ZOrder (0)
        .Visible = True
        
        .Nodes.Clear
        .LineStyle = tvwTreeLines
        .Style = tvwTreelinesPlusMinusPictureText
        Set nodoPadre = .Nodes.Add(, , "R " & CStr(-1), "Plan De Cuenta", 1, 1)
        .Nodes("R " & CStr(-1)).Expanded = True
    End With

    Dim rsCuentas As New ADODB.Recordset, sqlCuentas As String
    Dim vCodigoCuenta As String, vMostrarCodigo As String, vCodigoPadre As String
    Dim vSaldoCuenta As Double
    
    sqlCuentas = "SELECT * FROM cuentas ORDER BY CodigoCuenta ASC"
    
    With rsCuentas
        .CursorLocation = adUseClient
        
        Call .Open(sqlCuentas, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If .State = 1 Then
            If Not .EOF = True Then
                .MoveFirst
                barra.Max = .RecordCount
            End If
            
            Call BorrarBase("TempSaldosCuentas", pathDBMySQL)
            
            Do Until .EOF = True
                DoEvents
          
                Call GuardarTempSaldosCuentas(.Fields("CodigoCuenta").Value, LCase(.Fields("Cuenta").Value))
                
                .MoveNext
                
                barra.Value = .AbsolutePosition
            Loop

            
        Else
            Exit Sub
        End If
    
    End With


    sqlCuentas = ""
    
    If rsCuentas.State = 1 Then
        rsCuentas.Close
        Set rsCuentas = Nothing
    End If

If Err Then GrabarLog "CargarCuentas", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub GuardarTempSaldosCuentas(vCodigoCuenta As String, vNombreCuenta As String)
On Error Resume Next

    Dim rsTSC As New ADODB.Recordset, sqlTSC As String
    Dim nodoPadre As Node
            
    sqlTSC = "SELECT * FROM TempSaldosCuentas"
    
    With rsTSC
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        
        Call .Open(sqlTSC, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If .State = 1 Then
        
            .AddNew
            
            .Fields("CodigoCuenta").Value = vCodigoCuenta
            .Fields("CodigoMostrar").Value = MostrarCodigoCuenta(vCodigoCuenta)

            .Fields("Nivel").Value = VerNivelCuenta(vCodigoCuenta)
            
            .Fields("CodigoPadre").Value = VerCodigoPadre(vCodigoCuenta, .Fields("Nivel").Value)
            
            If .Fields("Nivel").Value = 5 Then
                .Fields("Saldo").Value = 0 'CalcularSaldoCuenta(.Fields("CodigoCuenta").Value)
            Else
                .Fields("Saldo").Value = 0
            End If
            
            .Update
            
            Set nodoPadre = TVCuentas.Nodes.Add("R " & .Fields("CodigoPadre").Value, tvwChild, "R " & CStr(.Fields("CodigoCuenta").Value), vNombreCuenta, Len(.Fields("CodigoCuenta").Value) + 1, Len(.Fields("CodigoCuenta").Value) + 1)
            TVCuentas.Nodes("R " & CStr(.Fields("CodigoCuenta").Value)).Expanded = True
            
        End If
            
    End With
    
    sqlTSC = ""
    
    If rsTSC.State = 1 Then
        rsTSC.Close
        Set rsTSC = Nothing
    End If

If Err Then GrabarLog "GuardarTempSaldosCuentas", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub InicializarLlenadoArbol()

    barra.Value = 0

    ' Configuramos manualmente el Treeview
    With TVCuentas
        .Style = tvwTreelinesPlusMinusText
        .LineStyle = tvwRootLines
        .PathSeparator = "\"
        .Indentation = Screen.TwipsPerPixelX * 5 '256
        '
        ' No permitir la edición automática del texto
        .LabelEdit = tvwManual
        ' Para que se pueda expandir al seleccionar un nodo,
        ' cambia este valor a True,
        ' si se deja en False, se expande al hacer doble-click
        .SingleSel = False
        ' Para que al perder el foco,
        ' se siga viendo el que está seleccionado
        .HideSelection = False
        '
        .Refresh
    End With
    '
    PrepararImageList
    ' Llenar el Treeview con los nodos de la tabla Concepto
    
    DoEvents
    
    barra.Value = 0
    MousePointer = vbDefault

If Err Then GrabarLog "InicializarLlenadoArbol", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub PrepararImageList()
   Set TVCuentas.ImageList = ImageList2
End Sub
Private Function CalcularSaldoPorNivel(vnivel As Integer) As Double
On Error Resume Next

    Dim rsSaldoPorNivel As New ADODB.Recordset, sqlSaldoPorNivel As String
    
    sqlSaldoPorNivel = "SELECT * FROM TempSaldosCuentas ORDER BY CodigoCuenta"
    
    With rsSaldoPorNivel
        .CursorLocation = adUseClient
        
        Call .Open(sqlSaldoPorNivel, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If .State = 1 Then
            If Not .EOF = True Then
                .MoveFirst
                barra.Max = .RecordCount
                barra.Value = 0
            End If
            
            Do Until .EOF = True
                AsignarSaldoEnNodo ("R " & .Fields("CodigoCuenta").Value)
                .MoveNext
            
                barra.Value = .AbsolutePosition
            Loop
        
        End If
    
    End With
    
    sqlSaldoPorNivel = ""

    If rsSaldoPorNivel.State = 1 Then
        rsSaldoPorNivel.Close
        Set rsSaldoPorNivel = Nothing
    End If
If Err Then GrabarLog "CalcularSaldoPorNivel", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub AsignarSaldoEnNodo(vNodo As String)
On Error Resume Next
    
    Dim vSaldoCuenta As Double, vCodigoCuenta As String
    Dim i As Integer
    
    vCodigoCuenta = Replace(vNodo, "R ", "")
    
    Select Case Len(vCodigoCuenta)
        
        Case 1
            vSaldoCuenta = Val(Format(GenerarDato("SELECT Sum(AsientosDetalle.Debe) AS SumaDeDebe, Sum(AsientosDetalle.Haber) AS SumaDeHaber, Sum(AsientosDetalle.Debe-AsientosDetalle.Haber) AS Saldo FROM AsientosDetalle WHERE (Mid(CodigoCuenta,1,1) = '" & Mid(vCodigoCuenta, 1, 1) & "')", "Saldo"), "#######0.00"))
                
        Case 2
            vSaldoCuenta = Val(Format(GenerarDato("SELECT Sum(AsientosDetalle.Debe) AS SumaDeDebe, Sum(AsientosDetalle.Haber) AS SumaDeHaber, Sum(AsientosDetalle.Debe-AsientosDetalle.Haber) AS Saldo FROM AsientosDetalle WHERE (Mid(CodigoCuenta, 1,2) = '" & Mid(vCodigoCuenta, 1, 2) & "')", "Saldo"), "#######0.00"))
        
        Case 4
            vSaldoCuenta = Val(Format(GenerarDato("SELECT Sum(AsientosDetalle.Debe) AS SumaDeDebe, Sum(AsientosDetalle.Haber) AS SumaDeHaber, Sum(AsientosDetalle.Debe-AsientosDetalle.Haber) AS Saldo FROM AsientosDetalle WHERE (Mid(CodigoCuenta,1 ,4) = '" & Mid(vCodigoCuenta, 1, 4) & "')", "Saldo"), "#######0.00"))
        
        Case 6
            vSaldoCuenta = Val(Format(GenerarDato("SELECT Sum(AsientosDetalle.Debe) AS SumaDeDebe, Sum(AsientosDetalle.Haber) AS SumaDeHaber, Sum(AsientosDetalle.Debe-AsientosDetalle.Haber) AS Saldo FROM AsientosDetalle WHERE (Mid(CodigoCuenta,1,6) = '" & Mid(vCodigoCuenta, 1, 6) & "')", "Saldo"), "#######0.00"))
        
        Case 8
            vSaldoCuenta = Val(Format(GenerarDato("SELECT Sum(AsientosDetalle.Debe) AS SumaDeDebe, Sum(AsientosDetalle.Haber) AS SumaDeHaber, Sum(AsientosDetalle.Debe-AsientosDetalle.Haber) AS Saldo FROM AsientosDetalle WHERE (mid(CodigoCuenta,1,8) = '" & Mid(vCodigoCuenta, 1, 8) & "')", "Saldo"), "#######0.00"))
            
    End Select
            
                
    With TVCuentas
        For i = 2 To Val(.Nodes.Count)
            If .Nodes.Item(i).key = vNodo Then
                .Nodes.Item(i).Text = .Nodes.Item(i).Text & "...$ " & Abs(vSaldoCuenta)
                Exit For
            End If
        
        Next
    
    End With

If Err Then GrabarLog "AsignarSaldoEnNodo", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub TVCuentas_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim S As String
    
    S = Node.Text
    If Node.children > 0 Then
        S = S & ", tiene " & Node.children & " hijos"
    Else
        S = S & ", no tiene hijos"
    End If
        
    Dim cmdConcepto As New ADODB.Command
    Dim sqlConcepto As String
    cmdConcepto.ActiveConnection = ConnDDBB
      
    sqlConcepto = "SELECT * FROM Cuentas WHERE (CodigoCuenta = '" & Mid(Node.key, 3, ContarCaracteres(Node.key, "-") - 2) & "')"
    
    Dim rsConcepto As New ADODB.Recordset
      
    If rsConcepto.State = 0 Then
        rsConcepto.Open sqlConcepto, ConnDDBB, 3, 3
    Else
        Set rsConcepto = ConnDDBB.Execute(sqlConcepto)
    End If
      
    If Not rsConcepto.EOF Then
        rsConcepto.MoveFirst
        'txtCodigo(0).Text = rsConcepto("CodigoCuenta").Value
        'txtCodigo(1).Text = rsConcepto("Cuenta").Value
        
        'If rsConcepto("Imputable").Value = "S" Then
        '    chkImputable.Value = xtpUnchecked
        'Else
        '    chkImputable.Value = xtpChecked
        'End If
    End If
    
    ' El nodo que está actualmente seleccionado
    Set tNodo = TVCuentas.SelectedItem
    
End Sub
Private Sub AgregarNodo()
On Error Resume Next
    
    'Dim tNodo As Node
    'Dim sP As String, sH As String
    'Dim i As Long
    
    ' El nodo que está actualmente seleccionado
    'Set tNodo = TVCuentas.SelectedItem
   
    'If Me.TreeView1.SelectedItem Is Not Null Then
    
    'sP = tNodo.Key
    ' Cantidad de hijos
    'i = tNodo.Children
    
    '    Do
    '        i = i + 1
    '        sH = sP & "-" & CStr(i)
    '        Err = 0
            ' Añadirlo como nodo hijo del seleccionado
    '        TVCuentas.Nodes.Add sP, tvwChild, sH, Trim(txtCodigo(1).Text), 4, 4
            ' Si no da error, salir del bucle
        
   '         If Err.Description <> "" Then
   '             MsgBox "Debe seleccionar un nodo padre para el nuevo concepto"
   '             Exit Sub
   '         Else
   '             Exit Do
   '         End If
   '
   '         Exit Do
   '     Loop
   '
   '
   ' Dim ingresoEgreso As Boolean
   ' Dim spSinLetras As String
   ' spSinLetras = Mid(sP, 3, ContarCaracteres(sP, "-") - 2)
   '
   '
If Err Then
    'MsgBox "Error al guardar el concepto " & Trim(txtCodigo(0).Text)
    'GrabarLog "AgregarNodo", Err.Number & " " & Err.Description, Me.Name
End If

End Sub
Private Sub txtBuscar_Change()
    Buscar
End Sub
