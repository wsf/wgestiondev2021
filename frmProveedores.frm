VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#13.0#0"; "Codejock.CommandBars.v13.0.0.Demo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmProveedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Proveedores"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   270
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   11490
   Begin VB.PictureBox PicInferior 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   2760
      Picture         =   "frmProveedores.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   8565
      TabIndex        =   2
      Top             =   0
      Width           =   8565
      Begin WGestion.AlphaIcon IconoFormulario 
         Height          =   600
         Left            =   5800
         Top             =   0
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   1058
         bvData          =   "frmProveedores.frx":50B3
         bData           =   -1  'True
         IconHeight      =   48
         IconWidth       =   48
         Stretch         =   -1  'True
      End
   End
   Begin XtremeSuiteControls.TabControl TabProveedores 
      Height          =   6015
      Left            =   60
      TabIndex        =   0
      Top             =   525
      Width           =   11280
      _Version        =   851968
      _ExtentX        =   19897
      _ExtentY        =   10610
      _StockProps     =   68
      Color           =   8
      ItemCount       =   3
      Item(0).Caption =   "Todos"
      Item(0).ControlCount=   4
      Item(0).Control(0)=   "dgProveedores"
      Item(0).Control(1)=   "txtBuscar"
      Item(0).Control(2)=   "lblBuscar"
      Item(0).Control(3)=   "Label1"
      Item(1).Caption =   ""
      Item(1).ControlCount=   0
      Item(2).Caption =   ""
      Item(2).ControlCount=   0
      Begin MSDataGridLib.DataGrid dgProveedores 
         Height          =   5055
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   11070
         _ExtentX        =   19526
         _ExtentY        =   8916
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
         Width           =   9675
         _Version        =   851968
         _ExtentX        =   17066
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   225
         Left            =   1260
         TabIndex        =   5
         Top             =   60
         Width           =   3615
         _Version        =   851968
         _ExtentX        =   6376
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "<F1>  para agregar una nueva persona"
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
      Begin XtremeSuiteControls.Label lblBuscar 
         Height          =   255
         Left            =   480
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
      Left            =   360
      Top             =   0
      _Version        =   851968
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmProveedores.frx":65BD
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
Attribute VB_Name = "frmProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsProveedores As ADODB.Recordset
Dim vsql As String, vSQLOrden As String
Dim sqlProveedores As String
Public vienePago As Boolean
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
        sqlProveedores = ""
        vSQLOrden = "Codigo ASC"
        
        
            If Val(txtBuscar.Text) > 0 Then
                vsql = vsql + " AND ((codigo = '" + Trim(txtBuscar.Text) + "'))"
            Else
                vsql = vsql + " AND ((nombre LIKE '%" + Trim(txtBuscar.Text) + "%') OR (codigo LIKE '%" + Trim(txtBuscar.Text) + "%')) "
            End If
        
        
        'If Not Trim(txtCodigo.Text) = "" Then vSQL = vSQL & " AND (codigo = '" + Trim(txtCodigo.Text) + "')"
        'If Not Trim(txtNombre.Text) = "" Then vSQL = vSQL & " AND (nombre LIKE '%" + Trim(txtNombre.Text) + "%')"
        'If Not Trim(txtLocalidad.Text) = "" Then vSQL = vSQL & " AND (Localidad LIKE '%" + Trim(txtLocalidad) + "%')"
        'If Not chkActivos.Value = 0 Then vSQL = vSQL & "AND (pasivo = 'NO')"
        
        Set rsProveedores = New ADODB.Recordset
        
        sqlProveedores = "SELECT * FROM " & vConfigGral.vEmpresa & ".Proveedores WHERE (1=1 " & vsql & ") ORDER BY " & Trim(vSQLOrden)
        
        With rsProveedores
            .CursorLocation = adUseClient
            Call .Open(sqlProveedores, ConnDDBB, adOpenStatic, adLockPessimistic)
            
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
    
    With dgProveedores
        Set .DataSource = rsProveedores
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
    
    End With

If Err Then GrabarLog "", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub CommandBars_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error Resume Next

    Select Case control.Index
                
        Case 1
            With frmProveedoresAlta
                .Show
                .vaccion = "Nuevo"
            End With
            
        Case 2
            If Not (rsProveedores.EOF = True) And Not (rsProveedores.BOF = True) Then
                With frmProveedoresAlta
                    .Show
                    .ModificarProveedor (rsProveedores.Fields("idProveedores").Value)
                    .vaccion = "Modificar"
                End With
            End If
            
        Case 3
            'Duplicar
        
        Case 4
            With rsProveedores
                If Not (.EOF = True) And Not (.BOF = True) Then
                    Call BorrarBase(vConfigGral.vEmpresa & ".Proveedores WHERE (idProveedores = " & .Fields("idProveedores").Value & ")", pathDBMySQL)
                    Buscar
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
Private Sub dgProveedores_DblClick()
On Error Resume Next

    Dim i As Integer

    If vienePago = True Then
        
        With frmCobros ' Ale: tuve que cambiar las referencias a los nombres de los met y variables para unificar cobro pago
            .esComprobanteAutomatico = True
            .KlexDetalle.Rows = 1
            frmCobros.BuscarDatosOperacionesCliente rsProveedores.Fields("codigo").Value, 0 ' Alfredo: Ale: cambiarle l
            .codCliente = EsNulo(rsProveedores.Fields("codigo").Value)
            .txtCliente(1).Text = rsProveedores.Fields("nombre").Value
            frmCobros.WindowState = vmaximizar
            .Show
        End With
        
        Unload Me
    Else
        With rsProveedores
            If Not (.EOF = True) And Not (.BOF = True) Then
                With frmProveedoresAlta
                    .Show
                    .ModificarProveedor (rsProveedores.Fields("idProveedores").Value)
                    .vaccion = "Modificar"
                End With
            End If
        End With
    
    End If
    
If Err Then GrabarLog "dgProveedores_DblClick", Err.Number & "-" & Err.Description, Me.Name
End Sub
Private Sub dgProveedores_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next

    Call OrdenarDataGrid(ColIndex, rsProveedores, dgProveedores)

    If Err Then GrabarLog "DgClientes_HeadClick", Err.Number & "-" & Err.Description, Me.Name
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF1 Then
            With frmProveedoresAlta
                .Show
                .vaccion = "Nuevo"
            End With
End If

End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    CargarBotonera
    
    With Me
        .Show
        .Top = 0
        .Left = 0
        .PicInferior.Top = -45
        .PicInferior.Left = 3475
    End With
    
    Call CentrarFormulario(Me)
    
    Call Buscar
    
    Me.txtBuscar.SetFocus
    
    If Err Then GrabarLog "Form_Load", Err.Number & "-" & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub txtBuscar_Change()
    Buscar
End Sub
Private Sub txtBuscar_KeyPress(Keyascii As Integer)
On Error Resume Next

    If Keyascii = 13 Then dgProveedores_DblClick
    
   
If Err Then GrabarLog "txtBuscar_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtBuscar_KeyUp(KeyCode As Integer, Shift As Integer)
 
 If KeyCode = vbKeyF1 Then
                With frmProveedoresAlta
                .Show
                .vaccion = "Nuevo"
            End With
    End If

End Sub
