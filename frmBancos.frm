VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#13.0#0"; "Codejock.CommandBars.v13.0.0.Demo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmBancos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Bancos y Cajas"
   ClientHeight    =   8340
   ClientLeft      =   2040
   ClientTop       =   1230
   ClientWidth     =   14535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   14535
   Begin XtremeSuiteControls.TabControl TabArticulos 
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
      Item(0).Control(0)=   "GroBancosY"
      Item(0).Control(1)=   "GroDetalles"
      Item(0).Control(2)=   "GroupBox1"
      Item(1).Caption =   ""
      Item(1).ControlCount=   0
      Item(2).Caption =   ""
      Item(2).ControlCount=   0
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   525
         Left            =   60
         TabIndex        =   5
         Top             =   360
         Width           =   14295
         _Version        =   851968
         _ExtentX        =   25215
         _ExtentY        =   926
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         BorderStyle     =   2
         Begin XtremeSuiteControls.FlatEdit txtBuscar 
            Height          =   285
            Left            =   1950
            TabIndex        =   6
            Top             =   180
            Width           =   12285
            _Version        =   851968
            _ExtentX        =   21669
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.Label lblBuscar 
            Height          =   255
            Left            =   1200
            TabIndex        =   7
            Top             =   180
            Width           =   735
            _Version        =   851968
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Buscar :"
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GroDetalles 
         Height          =   2355
         Left            =   60
         TabIndex        =   3
         Top             =   5370
         Width           =   14325
         _Version        =   851968
         _ExtentX        =   25268
         _ExtentY        =   4154
         _StockProps     =   79
         Caption         =   "Detalles"
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
         Begin MSDataGridLib.DataGrid dgCuentas 
            Height          =   1995
            Left            =   90
            TabIndex        =   4
            Top             =   270
            Width           =   14145
            _ExtentX        =   24950
            _ExtentY        =   3519
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
      End
      Begin XtremeSuiteControls.GroupBox GroBancosY 
         Height          =   4455
         Left            =   60
         TabIndex        =   1
         Top             =   900
         Width           =   14295
         _Version        =   851968
         _ExtentX        =   25215
         _ExtentY        =   7858
         _StockProps     =   79
         Caption         =   "Bancos y Caja"
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
         Begin MSDataGridLib.DataGrid dgBancos 
            Height          =   4125
            Left            =   90
            TabIndex        =   2
            Top             =   240
            Width           =   14115
            _ExtentX        =   24897
            _ExtentY        =   7276
            _Version        =   393216
            HeadLines       =   1
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
      End
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
      Left            =   360
      Top             =   0
      _Version        =   851968
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmBancos.frx":0000
   End
End
Attribute VB_Name = "frmBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsBancos As New ADODB.Recordset
Dim rsBancosCuentas As New ADODB.Recordset
Public Sub Buscar()
    On Error Resume Next
    
    Dim sqlBancos As String
    Dim vsql As String, vorden As String, i As Integer
    
    MousePointer = vbHourglass
    
    'For i = 0 To opOrden.Count - 1
    '    If opOrden(i).Value = True Then
    '        vorden = opOrden(i).Tag
    '    End If
    'Next
    
    vsql = ""

    'If chkFaltantes.Value = 1 Then
    '    vSQL = vSQL + " and (Stock <= 0)"
    'End If

    If Not txtBuscar.Text = "" Then
        vsql = vsql + " AND ((Descripcion LIKE '%" + Trim(txtBuscar.Text) + "%') OR (idBancos LIKE '%" + Trim(txtBuscar.Text) + "%'))"
    End If

    'If Not txtBusqueda(1).Text = "" Then
    '    vSQL = vSQL + " AND (proveedor LIKE '%" + Trim(txtBusqueda(1).Text) + "%')"
    'End If

    'If Not Trim(txtBusqueda(2).Text) = "" Then
    '    vSQL = vSQL + " AND (Rubro_Descrip LIKE '%" + Trim(txtBusqueda(2).Text) + "%')"
    'End If

    'If chkInvertir.Value = 0 Then
        'sqlBancos = "SELECT * FROM " & vConfigGral.vEmpresa & ".Bancos WHERE 1=1" + vSQL + " ORDER BY 1" & vorden
    'Else
    '    sqlArticulos = "SELECT * FROM Articulos WHERE NOT (1=1" + vSQL + ") ORDER BY " & vorden
    'End If

    sqlBancos = "SELECT B.idBancos, B.Descripcion, B.EsCaja, B.CuentaContableAsociada, C.Cuenta, b.tipodisponibilidad FROM Bancos B LEFT JOIN Cuentas C ON CuentaContableAsociada=CodigoCuenta WHERE 1=1 " & vsql & ""
    
    With rsBancos
        If .State = 1 Then .Close
        Call .Open(sqlBancos, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not .EOF = True Then
            .MoveFirst
            FormatoGrilla (0)
        Else

        End If

    End With

    MousePointer = vbDefault

    If Err Then GrabarLog "cmdFiltrar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub FormatoGrilla(vIndex As Byte)
On Error Resume Next
    
    Dim i As Integer
    
    If vIndex = 0 Then
    
        With dgBancos
            Set .DataSource = rsBancos
            .HeadLines = 2
        
            .ScrollBars = dbgVertical
        
            For i = 0 To .Columns.Count - 1
                .Columns(i).Width = 0
            Next
        
            .Columns(0).Width = 750
            .Columns(0).Caption = "Codigo"
            
            .Columns(1).Width = 5000
            .Columns(1).Caption = "Nombre del Banco"
            
            .Columns(2).Width = 1000
            .Columns(2).Caption = "¿ Es Caja ?"
            
            .Columns(3).Width = 2000
            .Columns(3).Caption = "Codigo Cuenta Contable"
            
            .Columns(4).Width = 3500
            .Columns(4).Caption = "Nombre de la Cuenta"
        
            .Columns(5).Width = 1500
            .Columns(5).Caption = "Disponibilidad"
        
        
        End With
    
    Else
        With dgCuentas
            Set .DataSource = rsBancosCuentas
            .HeadLines = 2
        
            .ScrollBars = dbgVertical
        
            For i = 0 To .Columns.Count - 1
                .Columns(i).Width = 0
            Next
        
            .Columns(1).Width = 750
            .Columns(1).Caption = "Codigo"
            
            .Columns(2).Width = 5000
            .Columns(2).Caption = "Descripcion de la Cuenta"
            
            .Columns(3).Width = 1000
            .Columns(3).Caption = "Cuenta"
            
            .Columns(5).Width = 1000
            .Columns(5).Caption = "Tipo de Cuenta"
            
            .Columns(6).Width = 2000
            .Columns(6).Caption = "Codigo Cuenta Contable"
            
            .Columns(7).Width = 3500
            .Columns(7).Caption = "Nombre de la Cuenta"
        
        End With
    
    End If
    
If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub dgBancos_Click()
On Error Resume Next

    With rsBancos
        If Not (.EOF = True) And Not (.BOF = True) Then
            CargarCuentaBancaria (.Fields("idBancos").Value)
        End If
    
    End With

If Err Then GrabarLog "dgBancos_Click", Err.Number & "-" & Err.Description, Me.Name
End Sub
Private Sub CargarCuentaBancaria(vidBancos As String)
On Error Resume Next

    Dim sqlBancosCuentas As String
    
    sqlBancosCuentas = "SELECT idBancosCuentas, idBancos, Bc.Descripcion, BC.Cuenta,  T.idTipoCuentaBanco, T.TipoCuentaBanco, C.CodigoCuenta, C.Cuenta FROM BancosCuentas BC LEFT JOIN tipocuentabanco t ON BC.idTipoCuentaBanco=T.idTipoCuentaBanco LEFT JOIN Cuentas C ON BC.CuentaContableAsociada=C.CodigoCuenta WHERE (idBancos = '" & vidBancos & "')"
    
    With rsBancosCuentas
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        
        Call .Open(sqlBancosCuentas, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then
            .MoveFirst
            FormatoGrilla (1)
        End If
        
    End With

If Err Then GrabarLog "CargarCuentaBancaria", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub dgBancos_DblClick()
On Error Resume Next

        If Not (rsBancos.EOF = True) And Not (rsBancos.BOF = True) Then
            With frmBancosAlta
                .Show
                .ModificarBanco (rsBancos.Fields("idBancos").Value)
                .vAccion = "Modificar"
            End With
        End If

If Err Then GrabarLog "dgBancos_DblClick", Err.Number & "-" & Err.Description, Me.Name
End Sub
Private Sub dgCuentas_DblClick()
On Error Resume Next

        If Not (rsBancosCuentas.EOF = True) And Not (rsBancosCuentas.BOF = True) Then
            With frmBancosCuentaAlta
                .Show
                .ModificarBancosCuentas (rsBancosCuentas.Fields("idBancosCuentas").Value)
                .vAccion = "Modificar"
            End With
        End If

If Err Then GrabarLog "dgCuentas_DblClick", Err.Number & "-" & Err.Description, Me.Name
End Sub
Private Sub dgBancos_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next

    Call OrdenarDataGrid(ColIndex, rsBancos, dgBancos)

    If Err Then GrabarLog "dgBancos_HeadClick", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CommandBars_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error Resume Next
Dim vvsql, valor, vmensaje  As String


    Select Case control.Index
                
        Case 1
            With frmBancosAlta
                .Show
                .vAccion = "Nuevo"
            End With
            
        Case 2
            If Not (rsBancos.EOF = True) And Not (rsBancos.BOF = True) Then
                With frmBancosAlta
                    .Show
                    .ModificarBanco (rsBancos.Fields("idBancos").Value)
                    .vAccion = "Modificar"
                End With
            End If
            
        Case 3
            'Duplicar
        
        'frmArticulos
        Case 4
            With rsBancos
                If Not (.EOF = True) And Not (.BOF = True) Then
                    If MsgBox("Esta seguro que desea borrar este registro?", vbInformation + vbYesNo, "Mensaje ...") = vbYes Then
                    
                    
                        valor = ""
                        
                        vvsql = "select * from cheques where idbancos like '%" + Trim(.Fields("idBancos").Value) + "'"
                        valor = traerDatos2(vvsql, "idbancos", pathDBMySQL)
                    
                        vvsql = "select * from bancosmovimientos where idbancos like '%" + Trim(.Fields("idBancos").Value) + "'"
                        valor = valor + traerDatos2(vvsql, "idbancos", pathDBMySQL)
                         
                         If Len(valor) > 0 Then
                            vmensaje = "Si borra este banco desaparen las referencias del módulo de cheques y los movimientos de bancos correspondientes"
                         
                            If MsgBox(vmensaje + Chr(13) + "Lo quiere borrar de todas manera ?", vbInformation + vbYesNo, "Mensaje ...") = vbNo Then
                                Exit Sub
                            End If
                         
                         End If
                    
                        Call BorrarBase(vConfigGral.vEmpresa & ".Bancos WHERE (idBancos = '" & .Fields("idBancos").Value & "')", pathDBMySQL)
                        Call BorrarBase(vConfigGral.vEmpresa & ".BancosCuentas WHERE (idBancos = '" & .Fields("idBancos").Value & "')", pathDBMySQL)
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
Private Sub Form_Load()
    On Error Resume Next

    CargarBotonera

    With Me
        .Show
        .KeyPreview = True
        '.PicInferior.Top = -45
        '.PicInferior.Left = 3425
        .Left = 0
        .Top = 0
    End With
    
    TabArticulos.Selected = 0
    
    Call CentrarFormulario(Me)
    
    Buscar
    
    If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CargarBotonera()
On Error Resume Next
    
    CommandBarsGlobalSettings.App = App
     
    Dim control As CommandBarControl
    Dim ToolBar As CommandBar
    Set ToolBar = CommandBars.Add("Standard", xtpBarTop)
    
    AddControl ToolBar.Controls, xtpControlButton, 2, "&Nuevo", False, "Crea un Nuevo Cliente"
    AddControl ToolBar.Controls, xtpControlButton, 11, "&Modificar", False, ""
    AddControl ToolBar.Controls, xtpControlButton, 5, "&Duplicar", False, ""
    AddControl ToolBar.Controls, xtpControlButton, 6, "&Borrar", False, ""
    ToolBar.Closeable = True
    AddControl ToolBar.Controls, xtpControlButton, 14, "Bu&scar", True, ""
    AddControl ToolBar.Controls, xtpControlButton, 27, "&Imprimir", False, ""
    ToolBar.Closeable = True
    AddControl ToolBar.Controls, xtpControlButton, 16, "&Salir", False, ""
    'AddControl ToolBar.Controls, xtpControlButton, 7, "&", True, ""
    'AddControl ToolBar.Controls, xtpControlButton, 8, "", False
    'AddControl ToolBar.Controls, xtpControlButton, 9, "Salir", False
      
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
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

    If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtBuscar_Change()
On Error Resume Next

    Buscar
    
If Err Then GrabarLog "txtBuscar_Change", Err.Number & " " & Err.Description, Me.Caption
End Sub
