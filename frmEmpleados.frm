VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#13.0#0"; "Codejock.CommandBars.v13.0.0.Demo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmEmpleados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Empleados"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   270
   ClientWidth     =   14535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   14535
   Begin VB.PictureBox PicInferior 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   3480
      Picture         =   "frmEmpleados.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   11325
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   11325
      Begin WGestion.AlphaIcon IconoFormulario 
         Height          =   555
         Left            =   10200
         Top             =   0
         Width           =   750
         _extentx        =   1323
         _extenty        =   979
         bdata           =   -1  'True
         bvdata          =   "frmEmpleados.frx":50B3
         iconwidth       =   48
         iconheight      =   48
         stretch         =   -1  'True
      End
   End
   Begin XtremeSuiteControls.TabControl TabEmpleados 
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   525
      Width           =   14535
      _Version        =   851968
      _ExtentX        =   25638
      _ExtentY        =   13785
      _StockProps     =   68
      Color           =   8
      ItemCount       =   3
      Item(0).Caption =   "Todos"
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "txtBuscar"
      Item(0).Control(1)=   "lblBuscar"
      Item(0).Control(2)=   "dgEmpleados"
      Item(1).Caption =   ""
      Item(1).ControlCount=   0
      Item(2).Caption =   ""
      Item(2).ControlCount=   0
      Begin MSDataGridLib.DataGrid dgEmpleados 
         Height          =   6735
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   14295
         _ExtentX        =   25215
         _ExtentY        =   11880
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   4
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
         Width           =   12495
         _Version        =   851968
         _ExtentX        =   22049
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
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
      Left            =   480
      Top             =   0
      _Version        =   851968
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmEmpleados.frx":6187
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
Attribute VB_Name = "frmEmpleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsempleados As ADODB.Recordset
Dim vsql As String, vSQLOrden As String
Dim sqlEmpleados As String
Public vieneCobro As Boolean
Private Sub CargarBotonera()
On Error Resume Next
    
    CommandBarsGlobalSettings.App = App
     
    Dim control As CommandBarControl
    Dim ToolBar As CommandBar
    Set ToolBar = CommandBars.Add("Standard", xtpBarTop)
    
    AddControl ToolBar.Controls, xtpControlButton, 2, "&Nuevo", False, "Crea un Nuevo Empleado"
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
        sqlEmpleados = ""
        vSQLOrden = "Codigo ASC"

        If Not txtBuscar.Text = "" Then
            vsql = vsql + " AND ((nombre LIKE '%" + Trim(txtBuscar.Text) + "%') OR (codigo LIKE '%" + Trim(txtBuscar.Text) + "%')) "
        End If
        
        Set rsempleados = New ADODB.Recordset
        
        If vConfigGral.vUsarEstaEmpresa = True Then
            sqlEmpleados = "SELECT * FROM " & vConfigGral.vEmpresa & ".Empleados E LEFT JOIN " & vConfigGral.vEmpresa & ".TipoIva Ti ON E.idTipoIva=Ti.idTipoIva WHERE (1=1 " & vsql & ") ORDER BY " & Trim(vSQLOrden)
        Else
            sqlEmpleados = "SELECT * FROM " & vConfigGral.vEmpresa & ".Empleados E LEFT JOIN " & vConfigGral.vEmpresa & ".TipoIva Ti ON E.idTipoIva=Ti.idTipoIva WHERE (1=1 " & vsql & ") "
        End If
        
        With rsempleados
            If .State = 1 Then .Close
            .CursorLocation = adUseClient
            Call .Open(sqlEmpleados, ConnDDBB, adOpenStatic, adLockPessimistic)
            
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
    
    With dgEmpleados
        Set .DataSource = rsempleados
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
        '.Columns(13).width = 1000
    
        .Columns(32).Caption = "Cond. Iva"
        .Columns(32).Width = 2000
    End With

If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub CommandBars_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error Resume Next

    Select Case control.Index
                
        Case 1
            With frmEmpleadosAlta
                .Show
                .vaccion = "Nuevo"
                .Caption = "Agregar Empleado"
                .vVieneEmpleadosAlta = Me.Name
            End With
            
        Case 2
            If Not (rsempleados.EOF = True) And Not (rsempleados.BOF = True) Then
                With frmEmpleadosAlta
                    .Show
                    .ModificarEmpleado (rsempleados.Fields("idEmpleados").Value)
                    .vaccion = "Modificar"
                    .Caption = "Modificar Empleado"
                    .vVieneEmpleadosAlta = Me.Name
                End With
            End If
            
        Case 3
            'Duplicar
           ' .vVieneEmpleadosAlta = Me.Name
        
        Case 4
            With rsempleados
                If Not (.EOF = True) And Not (.BOF = True) Then
                    If MsgBox("Esta seguro que desea borrar este registro?", vbInformation + vbYesNo, "Mensaje ...") = vbYes Then
                        Call BorrarBase(vConfigGral.vEmpresa & ".Empleados WHERE (idEmpleados = " & .Fields("idEmpleados").Value & ")", pathDBMySQL)
                        'Call BorrarBase(vConfigGral.vEmpresa & ".Articulosclientes WHERE (CodigoCliente = " & .Fields("Codigo").Value & ")", pathDBMySQL)
                        
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
Private Sub dgEmpleados_DblClick()
On Error Resume Next

    If vieneCobro = True Then
    '    frmCobros.BuscarDatosOperacionesCliente rsEmpleados.Fields("codigo").Value, 0
    '    frmCobros.esComprobanteAutomatico = True
    '    frmCobros.codCliente = rsClientes.Fields("codigo").Value
    '    frmCobros.Show
    '    Unload Me
    Else
        If Not (rsempleados.EOF = True) And Not (rsempleados.BOF = True) Then
            With frmEmpleadosAlta
                .Show
                .ModificarEmpleado (rsempleados.Fields("idEmpleados").Value)
                .vaccion = "Modificar"
                .Caption = "Modificar Empleado"
                .vVieneEmpleadosAlta = Me.Name
            End With
        End If
    End If
        
If Err Then GrabarLog "dgEmpleados_DblClick", Err.Number & "-" & Err.Description, Me.Name
End Sub
Private Sub dgEmpleados_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next

    Call OrdenarDataGrid(ColIndex, rsempleados, dgEmpleados)

    If Err Then GrabarLog "dgEmpleados_HeadClick", Err.Number & "-" & Err.Description, Me.Name
End Sub
Private Sub Form_KeyPress(Keyascii As Integer)
On Error Resume Next

If Err Then GrabarLog "Form_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Load()
    On Error Resume Next
    
    CargarBotonera
    
    With Me
        .Show
        .Height = 8850
        .Width = 14625
        .Top = 0
        .Left = 0
        .PicInferior.Visible = True
        .PicInferior.Top = -45
        .PicInferior.Left = 3475
    End With
    
    Buscar
    
    If Err Then GrabarLog "Form_Load", Err.Number & "-" & Err.Description, Me.Name
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
Private Sub txtBuscar_Change()
    Buscar
End Sub
Private Sub txtBuscar_KeyPress(Keyascii As Integer)
On Error Resume Next

    If Keyascii = 13 Then dgEmpleados_DblClick

If Err Then GrabarLog "txtBuscar_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
