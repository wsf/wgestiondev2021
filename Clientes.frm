VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#13.0#0"; "Codejock.CommandBars.v13.0.0.Demo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de clientes"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   270
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   11850
   Begin VB.PictureBox PicInferior 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   3270
      Picture         =   "Clientes.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   9075
      TabIndex        =   2
      Top             =   -30
      Visible         =   0   'False
      Width           =   9075
      Begin WGestion.AlphaIcon IconoFormulario 
         Height          =   555
         Left            =   7500
         Top             =   0
         Width           =   750
         _extentx        =   1323
         _extenty        =   979
         bdata           =   -1  'True
         bvdata          =   "Clientes.frx":50B3
         iconwidth       =   48
         iconheight      =   48
         stretch         =   -1  'True
      End
   End
   Begin XtremeSuiteControls.TabControl TabClientes 
      Height          =   5835
      Left            =   0
      TabIndex        =   0
      Top             =   705
      Width           =   11970
      _Version        =   851968
      _ExtentX        =   21114
      _ExtentY        =   10292
      _StockProps     =   68
      Color           =   8
      ItemCount       =   3
      Item(0).Caption =   "Todos"
      Item(0).ControlCount=   7
      Item(0).Control(0)=   "dgClientes"
      Item(0).Control(1)=   "txtBuscar"
      Item(0).Control(2)=   "lblBuscar"
      Item(0).Control(3)=   "PusExcel"
      Item(0).Control(4)=   "vlocalidad"
      Item(0).Control(5)=   "lblLocalidad"
      Item(0).Control(6)=   "PusFiltar"
      Item(1).Caption =   ""
      Item(1).ControlCount=   0
      Item(2).Caption =   ""
      Item(2).ControlCount=   0
      Begin XtremeSuiteControls.PushButton PusFiltar 
         Height          =   345
         Left            =   10530
         TabIndex        =   8
         Top             =   450
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Filtar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PusExcel 
         Height          =   285
         Left            =   10830
         TabIndex        =   5
         Top             =   0
         Width           =   1005
         _Version        =   851968
         _ExtentX        =   1773
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Excel"
         UseVisualStyle  =   -1  'True
      End
      Begin MSDataGridLib.DataGrid dgClientes 
         Height          =   4785
         Left            =   180
         TabIndex        =   1
         Top             =   900
         Width           =   11550
         _ExtentX        =   20373
         _ExtentY        =   8440
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
      Begin XtremeSuiteControls.FlatEdit txtBuscar 
         Height          =   285
         Left            =   1470
         TabIndex        =   3
         Top             =   480
         Width           =   3525
         _Version        =   851968
         _ExtentX        =   6218
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vlocalidad 
         Height          =   285
         Left            =   6000
         TabIndex        =   6
         Top             =   480
         Width           =   3705
         _Version        =   851968
         _ExtentX        =   6535
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.Label lblLocalidad 
         Height          =   255
         Left            =   5160
         TabIndex        =   7
         Top             =   510
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Localidad:"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblBuscar 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3087
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Buscar :"
         Transparent     =   -1  'True
      End
   End
   Begin XtremeCommandBars.ImageManager ImageManager1 
      Left            =   480
      Top             =   0
      _Version        =   851968
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "Clientes.frx":602F
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
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsClientes As ADODB.Recordset
Dim vsql As String, vSQLOrden As String
Dim sqlClientes As String
Public vieneCobro As Boolean
Dim md5PasswordCliente As MD5
Private Sub CargarBotonera()
On Error Resume Next
    
    CommandBarsGlobalSettings.App = App
     
    Dim control As CommandBarControl
    Dim ToolBar As CommandBar
    Set ToolBar = CommandBars.Add("Standard", xtpBarTop)
    
    AddControl ToolBar.Controls, xtpControlButton, 2, "&Nuevo", False, "Crea un Nuevo Proveedor"
    AddControl ToolBar.Controls, xtpControlButton, 11, "&Modificar", False, ""
    AddControl ToolBar.Controls, xtpControlButton, 5, "&Duplicar", False, ""
    AddControl ToolBar.Controls, xtpControlButton, 6, "&Borrar", False, ""
    ToolBar.Closeable = True
    AddControl ToolBar.Controls, xtpControlButton, 14, "Bu&scar", True, ""
    AddControl ToolBar.Controls, xtpControlButton, 26, "&Imprimir", True, ""
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
Public Sub Buscar()
    On Error Resume Next
    
    If 1 = 0 Then
    
    Else
        
        vsql = ""
        sqlClientes = ""
        vSQLOrden = "Codigo ASC"

        If Not txtBuscar.Text = "" Then
            If Val(txtBuscar.Text) > 0 Then
                vsql = vsql + " AND ((codigo = '" + Trim(txtBuscar.Text) + "'))"
            Else
                vsql = vsql + " AND ((nombre LIKE '%" + Trim(txtBuscar.Text) + "%') OR (codigo LIKE '%" + Trim(txtBuscar.Text) + "%')) "
            End If
        End If
        
        If Me.vlocalidad.Text Then
            vsql = vsql + " AND ((localidad like '%" + Trim(vlocalidad.Text) + "%'))"
        End If
        
        
        Set rsClientes = New ADODB.Recordset
        
        If vConfigGral.vUsarEstaEmpresa = True Then
            sqlClientes = "SELECT * FROM " & vConfigGral.vempresa & ".Clientes C LEFT JOIN " & vConfigGral.vempresa & ".TipoIva Ti ON C.idTipoIva=Ti.idTipoIva WHERE (1=1 " & vsql & ") ORDER BY " & Trim(vSQLOrden)
        Else
            sqlClientes = "SELECT * FROM " & vConfigGral.vempresa & ".Clientes C LEFT JOIN " & vConfigGral.vempresa & ".TipoIva Ti ON C.idTipoIva=Ti.idTipoIva WHERE (1=1 " & vsql & ") "
        End If
        
        With rsClientes
            If .State = 1 Then .Close
            .CursorLocation = adUseClient
            Call .Open(sqlClientes, ConnDDBB, adOpenStatic, adLockPessimistic)
            
            If .State = 1 Then
                If Not .EOF = True Then .MoveLast
                
                FormatoGrilla
            End If
            
        End With
    
    End If
    
    If Err Then GrabarLog "Buscar", Err.Number & "-" & Err.Description, Me.Name
End Sub
Private Sub FormatoGrilla()
On Error Resume Next
    
    Dim i As Integer
    
    With dgClientes
        Set .DataSource = rsClientes
        .HeadLines = 2
        
        
        For i = 0 To .Columns.Count - 1
            .Columns(i).Width = 0
        Next
        
        .Columns(0).Width = 0
        .Columns(1).Width = 1000
        .Columns(2).Width = 0
        .Columns(3).Width = 2500
        .Columns(4).Width = 2500
        .Columns(6).Width = 1500
        .Columns(8).Width = 1500
        .Columns(11).Width = 1000
        .Columns(12).Width = 1000
        .Columns(13).Width = 1000
    
        .Columns(39).Caption = "Cond. Iva"
        .Columns(39).Width = 1750
    End With
    
    


If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub CommandBars_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error Resume Next

    Select Case control.Index
                
        Case 1
            
            
            With frmClientesAlta
                .Show
                .vaccion = "Nuevo"
                .Caption = "Agregar cliente"
                .vVieneClientesAlta = Me.Name
            End With
            
        Case 2
            If Not (rsClientes.EOF = True) And Not (rsClientes.BOF = True) Then
                With frmClientesAlta
                    .Show
                    .ModificarCliente (rsClientes.Fields("idClientes").Value)
                    .vaccion = "Modificar"
                    .Caption = "Modificar cliente"
                End With
            End If
            
        Case 3
            'Duplicar
        
        Case 4
            With rsClientes
                If Not (.EOF = True) And Not (.BOF = True) Then
                    If MsgBox("Esta seguro que desea borrar este registro?", vbInformation + vbYesNo, "Mensaje ...") = vbYes Then
                        Call BorrarBase(vConfigGral.vempresa & ".Clientes WHERE (idClientes = " & .Fields("idClientes").Value & ")", pathDBMySQL)
                        Call BorrarBase(vConfigGral.vempresa & ".Articulosclientes WHERE (CodigoCliente = " & .Fields("Codigo").Value & ")", pathDBMySQL)
                        Call BorrarBase(vConfigGral.vempresa & ".FacturaAutomatica WHERE (CodigoCliente = " & .Fields("Codigo").Value & ")", pathDBMySQL)
                        
                        Buscar
                    End If
                End If
            End With

        
        Case 5
            Buscar
                
        Case 6
            vVieneImpresion = Me.Name
            frmImprimir.Show
            
        Case 7
            Unload Me
            
    End Select

If Err Then GrabarLog "CommandBars_Execute", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub dgClientes_DblClick()
On Error Resume Next

    If vieneCobro = True Then
        frmCobros.esComprobanteAutomatico = True
        Dim i As Integer
        'Borro los datos anteriores de la grilla
        frmCobros.KlexDetalle.Rows = 1
       
        frmCobros.BuscarDatosOperacionesCliente rsClientes.Fields("codigo").Value, 0
        
        frmCobros.codCliente = rsClientes.Fields("codigo").Value
        
        frmCobros.txtCliente(1) = rsClientes.Fields("nombre").Value
        
        frmCobros.WindowState = vmaximizar
        frmCobros.Show
        Call frmCobros.initCobro
    
        Unload Me
    Else
        If Not (rsClientes.EOF = True) And Not (rsClientes.BOF = True) Then
            With frmClientesAlta
                
               dgClientes.Col = 0
               .ModificarCliente (dgClientes.Text)
                '.dgClientes.Text
                '.ModificarCliente (rsClientes.Fields("idClientes").Value)
                .vaccion = "Modificar"
                .Caption = "Modificar cliente"
                .Show
            End With
        End If
    End If
        
If Err Then GrabarLog "dgClientes_DblClick", Err.Number & "-" & Err.Description, Me.Name
End Sub
Private Sub dgClientes_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next

    Call OrdenarDataGrid(ColIndex, rsClientes, dgClientes)

    If Err Then GrabarLog "dgClientes_HeadClick", Err.Number & "-" & Err.Description, Me.Name
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next

If Err Then GrabarLog "Form_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Load()
    On Error Resume Next
    
    CargarBotonera
    Call CentrarFormulario(Me)
     

    
    With Me
        .Show
        '.height = 8850
        '.width = 14625
        .Top = 0
        .Left = 0
        .PicInferior.Visible = True
        .PicInferior.Top = -45
        .PicInferior.Left = 3475
    End With
    
    Buscar
    
    CentrarFormulario (Me)
    
    txtBuscar.SetFocus

    CentrarFormulario Me

    If Err Then GrabarLog "Form_Load", Err.Number & "-" & Err.Description, Me.Name
End Sub
Private Sub CambiarPassword()
On Error Resume Next
    
    Set md5PasswordCliente = New MD5

    Dim rsClientePassword As New ADODB.Recordset, sqlClientePassword As String
    
    sqlClientePassword = "SELECT * FROM Clientes ORDER BY idClientes ASC"
    
    With rsClientePassword
        .CursorLocation = adUseClient
        
        Call .Open(sqlClientePassword, ConnDDBB, adOpenDynamic, adLockOptimistic)
        
        
        If Not .EOF = True Then .MoveFirst
        
        Do Until .EOF = True
            .Fields("PasswordWeb").Value = md5PasswordCliente.DigestStrToHexStr(.Fields("PasswordWeb").Value)
            .MoveNext
        Loop

    
    End With
    
    sqlClientePassword = ""
    
    If rsClientePassword.State = 1 Then
        rsClientePassword.Close
        Set rsClientePassword = Nothing
    End If

    Set md5PasswordCliente = Nothing
    
If Err Then GrabarLog "CambiarPassword", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub ImprimirEnvases()
On Error Resume Next
   
      'drinfo_cli2.Show

If Err Then GrabarLog "ImprimirEnvases", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub ImprimirEstadoGral()
On Error Resume Next
       
       'drInfo_cli3.Show

If Err Then GrabarLog "ImprimirEstadoGral", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub ImprimirPasivos()
On Error Resume Next
     
    With Mantenimiento.rsPasivos
        If Not .State = 1 Then .Open
        .Close
        .Open
        
        .Sort = vSQLOrden
    End With
    
    If MsgBox("¿Desea imprimir los clientes suspendidos con saldo por reparto?", vbInformation + vbYesNo) = vbYes Then
        With Mantenimiento.rsPasivos
            
            '.filter = "Reparto = '" + Trim(txtNombre.Text) + "'"
            .Sort = "CodigoNum"
        End With
    Else
        With Mantenimiento.rsPasivos
            .Filter = "ID > 0"
            .Sort = "CodigoNum"
        End With

    End If
    
    'With drpasivos
    '    '.Sections(2).Controls("ereparto").caption = "SUSPENDIDOS"
    '    .Show
    'End With

If Err Then GrabarLog "ImprimirPasivos", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub PusExcel_Click()
    Call grillaToExcel3(Me.dgClientes, rsClientes.RecordCount)
End Sub

Private Sub txtBuscar_Change()
    Buscar
End Sub
Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then dgClientes_DblClick

If Err Then GrabarLog "txtBuscar_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub vlocalidad_Change()
Call Buscar
End Sub
