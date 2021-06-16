VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#13.0#0"; "Codejock.CommandBars.v13.0.0.Demo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#13.0#0"; "Codejock.ReportControl.v13.0.0.Demo.ocx"
Begin VB.Form frmArticulos 
   Caption         =   "Mantenimiento de Artículos"
   ClientHeight    =   7785
   ClientLeft      =   2055
   ClientTop       =   1245
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7785
   ScaleWidth      =   11130
   Begin XtremeSuiteControls.TabControl TabArticulos 
      Height          =   7155
      Left            =   0
      TabIndex        =   0
      Top             =   525
      Width           =   19470
      _Version        =   851968
      _ExtentX        =   34343
      _ExtentY        =   12621
      _StockProps     =   68
      ItemCount       =   3
      Item(0).Caption =   "Todos"
      Item(0).ControlCount=   6
      Item(0).Control(0)=   "GroupBox1"
      Item(0).Control(1)=   "Frame1"
      Item(0).Control(2)=   "PusConsultarTarifaria"
      Item(0).Control(3)=   "PushButton1"
      Item(0).Control(4)=   "lblInsertPara"
      Item(0).Control(5)=   "PusExcel"
      Item(1).Caption =   "Relaciones entre artículos, ventas, compras, pagos , cobros"
      Item(1).ControlCount=   0
      Item(2).Caption =   "Agrupar artículos"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "report"
      Begin XtremeReportControl.ReportControl report 
         Height          =   7875
         Left            =   -69880
         TabIndex        =   17
         Top             =   450
         Visible         =   0   'False
         Width           =   16695
         _Version        =   851968
         _ExtentX        =   29448
         _ExtentY        =   13891
         _StockProps     =   64
      End
      Begin XtremeSuiteControls.PushButton PusExcel 
         Height          =   315
         Left            =   16410
         TabIndex        =   23
         Top             =   780
         Width           =   1425
         _Version        =   851968
         _ExtentX        =   2514
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Excel"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   345
         Left            =   16410
         TabIndex        =   18
         Top             =   390
         Width           =   1395
         _Version        =   851968
         _ExtentX        =   2461
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Recalcular Stock"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   795
         Left            =   60
         TabIndex        =   3
         Top             =   330
         Width           =   13815
         Begin XtremeSuiteControls.CheckBox chkChkTodos 
            Height          =   240
            Left            =   12195
            TabIndex        =   25
            Top             =   90
            Width           =   1365
            _Version        =   851968
            _ExtentX        =   2408
            _ExtentY        =   423
            _StockProps     =   79
            Caption         =   "Mostar Todos"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PushButton3 
            Height          =   345
            Left            =   12210
            TabIndex        =   24
            Top             =   315
            Width           =   915
            _Version        =   851968
            _ExtentX        =   1614
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "Buscar"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.CommandButton cmdCommand1 
            Caption         =   "Command1"
            Height          =   225
            Left            =   6660
            TabIndex        =   20
            Top             =   510
            Visible         =   0   'False
            Width           =   1515
         End
         Begin XtremeSuiteControls.PushButton PushButton2 
            Height          =   195
            Left            =   8430
            TabIndex        =   19
            Top             =   210
            Visible         =   0   'False
            Width           =   435
            _Version        =   851968
            _ExtentX        =   767
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "PushButton2"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RBTipoCarga 
            Height          =   255
            Index           =   0
            Left            =   1440
            TabIndex        =   4
            Tag             =   "Articulos"
            Top             =   510
            Width           =   1815
            _Version        =   851968
            _ExtentX        =   3201
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Desde Articulos"
            Appearance      =   6
         End
         Begin XtremeSuiteControls.FlatEdit txtBuscar 
            Height          =   285
            Left            =   840
            TabIndex        =   5
            Top             =   180
            Width           =   2985
            _Version        =   851968
            _ExtentX        =   5265
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.RadioButton RBTipoCarga 
            Height          =   255
            Index           =   1
            Left            =   3330
            TabIndex        =   6
            Tag             =   "CargadoPorRemito"
            Top             =   510
            Width           =   1815
            _Version        =   851968
            _ExtentX        =   3201
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Desde Remito"
            Appearance      =   6
         End
         Begin XtremeSuiteControls.RadioButton RBTipoCarga 
            Height          =   255
            Index           =   2
            Left            =   5250
            TabIndex        =   7
            Top             =   510
            Width           =   1815
            _Version        =   851968
            _ExtentX        =   3201
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Ambos"
            Appearance      =   6
         End
         Begin XtremeSuiteControls.FlatEdit txtproveedor 
            Height          =   285
            Left            =   6120
            TabIndex        =   10
            Top             =   150
            Width           =   2265
            _Version        =   851968
            _ExtentX        =   3995
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit txtrubro 
            Height          =   285
            Left            =   9990
            TabIndex        =   13
            Top             =   120
            Width           =   1965
            _Version        =   851968
            _ExtentX        =   3466
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit txtsubrubro 
            Height          =   285
            Left            =   9990
            TabIndex        =   15
            Top             =   420
            Width           =   1965
            _Version        =   851968
            _ExtentX        =   3466
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.Label lblBuscar 
            Height          =   255
            Index           =   4
            Left            =   9120
            TabIndex        =   16
            Top             =   420
            Width           =   885
            _Version        =   851968
            _ExtentX        =   1561
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Sub-Rubro:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblBuscar 
            Height          =   255
            Index           =   3
            Left            =   9450
            TabIndex        =   14
            Top             =   150
            Width           =   735
            _Version        =   851968
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Rubro:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblBuscar 
            Height          =   255
            Index           =   2
            Left            =   3870
            TabIndex        =   11
            Top             =   180
            Width           =   2205
            _Version        =   851968
            _ExtentX        =   3889
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Parte del nombre deproveedor:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblBuscar 
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   9
            Top             =   180
            Width           =   1755
            _Version        =   851968
            _ExtentX        =   3087
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Descrip:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblBuscar 
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   8
            Top             =   510
            Width           =   1755
            _Version        =   851968
            _ExtentX        =   3087
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Modo de Carga :"
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   6015
         Left            =   30
         TabIndex        =   1
         Top             =   1050
         Width           =   19365
         _Version        =   851968
         _ExtentX        =   34158
         _ExtentY        =   10610
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid g 
            Height          =   5685
            Left            =   90
            TabIndex        =   22
            Top             =   135
            Width           =   19215
            _ExtentX        =   33893
            _ExtentY        =   10028
            _Version        =   393216
            BackColorSel    =   1375373
            ForeColorSel    =   0
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSDataGridLib.DataGrid dgArticulos 
            Height          =   5805
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   19215
            _ExtentX        =   33893
            _ExtentY        =   10239
            _Version        =   393216
            AllowUpdate     =   0   'False
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
      Begin XtremeSuiteControls.PushButton PusConsultarTarifaria 
         Height          =   345
         Left            =   13920
         TabIndex        =   12
         Top             =   360
         Width           =   2295
         _Version        =   851968
         _ExtentX        =   4048
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Consultar Tarifaria del artículo"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label lblInsertPara 
         Caption         =   "<+> para modificar precios"
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
         Left            =   13950
         TabIndex        =   21
         Top             =   750
         Width           =   2295
      End
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   45
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
      Icons           =   "frmArticulo.frx":0000
   End
End
Attribute VB_Name = "frmArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsArticulos As New ADODB.Recordset
Dim cArticulos As String
Dim vcodigo, vformatoCampos, vgralfiltro As String
Dim vidArticulos As Long
Dim vulinea As Integer

Public Sub init()
'vformatoCampos = "|<Codigo | <Descrip| <Rubro| <SubRubro| <Proveedor| < Fabricante  > PCosto| > PVenta1| > PVenta2 | < Fecha_Modificacion | >Stock"
'cArticulos = "Codigo,Descrip,Rubro,SubRubro,Proveedor,Fabricante,Porcentaje,format(PCosto,'###,###,##0.000') as pcosto,format(PVenta1,'###,###,##0.000') as pventa1 ,format(PVenta2,'###,###,##0.000') as pventa2,Fecha_Modificacion,Stock, idArticulos"
cArticulos = "Codigo,Descrip,Rubro,SubRubro,Proveedor,Fabricante,Porcentaje,format(PCosto,4) as pcosto,format(PVenta1,4) , format(PVenta2,4),format(PVenta1*1.21,4),Fecha_Modificacion,Stock, idArticulos"


If UCase(LeerXml("Cliente")) = "HERNAN" Then
    dgArticulos.Visible = True
    g.Visible = False
Else
    dgArticulos.Visible = False
    g.Visible = True
End If

End Sub


Public Sub Buscar(vFiltrar As String, Optional vvtipo As String)
    On Error Resume Next
    
    Dim sqlArticulos, vnot  As String
    Dim vsql, vsql2 As String, vorden As String, i As Integer
    
    MousePointer = vbHourglass

    vsql = ""
    vsql2 = ""
    
    If Not Me.txtrubro.Text = "" Then vsql2 = vsql2 + " and (idrubros = '" + Trim(Me.txtrubro.Tag) + "') "
    
    If Not Me.txtsubrubro.Text = "" Then vsql2 = vsql2 + " and (idsubrubros = '" + Trim(Me.txtsubrubro.Tag) + "') "
        
        
    If vFiltrar = "" Then

        If 1 = 1 Or txtBuscar.Text = "" Then
        
        
                                If UCase(LeerXml("Cliente")) = "HERNAN" Then
                                                vsql = vsql + " AND ((descrip LIKE '%" + Trim(txtBuscar.Text) + "%') OR (codigo LIKE '%" + Trim(txtBuscar.Text) + "%'))   "
                                End If
                                
        
                                If Not UCase(LeerXml("Cliente")) = "HERNAN" Then
                                
                                                If Val(txtBuscar.Text) > 0 Then
                                                
                                                    vsql = vsql + " AND ((codigo = '" + Trim(txtBuscar.Text) + "')) "
                                                
                                                Else
                                                     'vsql = vsql + " AND ((descrip LIKE '%" + Trim(txtBuscar.Text) + "%') OR (codigo LIKE '%" + Trim(txtBuscar.Text) + "%'))   "
                                                
                                                     vsql = vsql + " AND ((descrip LIKE '%" + Trim(txtBuscar.Text) + "%')) " + vsql2
                                               
                                                End If
                                               'If Not Trim(Me.txtProveedor = "") Then vsql = vsql + " and proveedor like '%" + Trim(txtProveedor.Text) + "%' "
                                    
                                End If
                                
        End If
        
     Else
        
        vsql = vFiltrar
        
        Me.Caption = vFiltrar
     
     End If
    
'    If frmActualizacionPrecio.chkNot = xtpChecked And Not vsql = "" Then
'        vnot = " not "
'            sqlArticulos = "SELECT * FROM " & vConfigGral.vEmpresa & ".VistaArticulos WHERE not (" + "1=1" + vsql + ")" + " ORDER BY 1"
'
'    Else
'
'            sqlArticulos = "SELECT * FROM " & vConfigGral.vEmpresa & ".VistaArticulos WHERE 1=1" + vsql + " ORDER BY 1"
'
'    End If
'
    ' sqlArticulos = "SELECT * FROM " & vConfigGral.vEmpresa & ".VistaArticulos WHERE 1=1" + vsql + " ORDER BY 1"
     
     'sqlArticulos = "SELECT " + cArticulos + " FROM " & vConfigGral.vEmpresa & ".VistaArticulos WHERE 1=1" + vsql + " ORDER BY 1 limit 100"
     
    If UCase(vvtipo) = UCase("todos") Then
        vvtipo = "  "
    Else
        vvtipo = " limit 100"
    End If
     
     
     If UCase(LeerXml("p1")) = "CARLETI" Then
        sqlArticulos = "SELECT " + cArticulos + " FROM VistaArticulos WHERE 1=1" + vsql + " ORDER BY 1 " + vvtipo
    Else
        sqlArticulos = "SELECT " + cArticulos + " FROM VistaArticulos WHERE 1=1" + vsql + " ORDER BY 2 " + vvtipo
    End If
    
    'sqlArticulos = "SELECT * FROM " & vConfigGral.vEmpresa & ".VistaArticulos WHERE 1=1" + vnot + "(" + vsql + ")" + " ORDER BY 1"
    'sqlArticulos = "SELECT * FROM " & vConfigGral.vEmpresa & ".VistaArticulos WHERE 1=1" + vsql + " ORDER BY 1"

    With rsArticulos
        
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        Call .Open(sqlArticulos, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not .EOF = True Then
            .MoveLast
            FormatoGrilla
        Else
           ' MsgBox "No hay datos", vbExclamation
        End If

    End With


    MousePointer = vbDefault

    If Err Then GrabarLog "cmdFiltrar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub FormatoGrilla()
On Error Resume Next
    
    Dim i As Integer
    


    With dgArticulos


        Set .DataSource = rsArticulos
        .HeadLines = 2

        '.ScrollBars = dbgAutomatic

        For i = 0 To .Columns.Count - 1
            .Columns(i).Width = 0
        Next

        .Columns(0).Width = 1500
        .Columns(0).Caption = "Codigo"

        .Columns(1).Width = 6000
        .Columns(1).Caption = "Descripcion del Articulo"

        .Columns(2).Width = 800
        .Columns(2).Caption = "Rubro"

        .Columns(3).Width = 700
        .Columns(3).Caption = "Sub-Rubro"

        .Columns(4).Width = 750
        .Columns(4).Caption = "% Iva"

        .Columns(5).Width = 600
        .Columns(5).Caption = "Proveedor"

        .Columns(6).Width = 500
        .Columns(6).Caption = "Fabricante"


        '.Columns(13).Width = 750
        '.Columns(13).Caption = "% Iva"

        .Columns(7).Width = 1000
        .Columns(7).Caption = "P. Costo"

        .Columns(8).Width = 1000
        .Columns(8).Caption = "P. Venta 1"

        .Columns(9).Width = 1000
        .Columns(9).Caption = "P. Venta 2"

        '.Columns(18).width = 1000
        '.Columns(18).Caption = "P. Venta 3"

        .Columns(10).Width = 1000
        .Columns(10).Caption = "Final"

        .Columns(11).Width = 1200
        .Columns(11).Caption = "Ult. Modif"

        .Columns(12).Width = 750
        .Columns(12).Caption = "Stock Actual"



    End With

    
    
      With g
      
      
   '  .FormatString = vformatoCampos

      
        Set .DataSource = rsArticulos
        
        g.Cols = 15
        '.ScrollBars = dbgAutomatic
        
        For i = 0 To .Cols - 1
            .ColWidth(i) = 0
        Next
        
        
        .ColWidth(1) = 1500
        .TextMatrix(0, 1) = "Codigo"
        
        .ColWidth(2) = 6000
        .TextMatrix(0, 2) = "Descrip"
        
        
        .ColWidth(3) = 1500
        .TextMatrix(0, 3) = "Rubro"
        
        .ColWidth(4) = 1000
        .TextMatrix(0, 4) = "Sub Rubro"
    
        
        .ColWidth(5) = 750
        .TextMatrix(0, 5) = "Iva"
        .ColAlignment(5) = 7
        
        
        .ColWidth(6) = 750
        .TextMatrix(0, 6) = "Proveedor"
        .ColAlignment(6) = 1
        
        .ColWidth(7) = 750
        .TextMatrix(0, 7) = "Fabricante"
        .ColAlignment(7) = 1
        
        
        .ColWidth(8) = 1000
        .TextMatrix(0, 8) = "P. Costo"
        .ColAlignment(8) = 7
        
        
        .ColWidth(9) = 1000
        .TextMatrix(0, 9) = "P.Venta 1"
        .ColAlignment(9) = 7
         
        
        .ColWidth(10) = 1000
        .TextMatrix(0, 10) = "P.Venta 2"
        .ColAlignment(10) = 7
    
        .ColWidth(11) = 1000
        .TextMatrix(0, 11) = "Final"
        
        
        .ColWidth(12) = 1000
        .TextMatrix(0, 12) = "Modif"
        
    
         .ColWidth(13) = 750
        .TextMatrix(0, 13) = "Stock"
        
              
        .ColWidth(14) = 1000
        .TextMatrix(0, 14) = "Id"
 
    End With
    
    
    
    
    
If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub GenerarReportes(Index As Integer)
    On Error Resume Next
    
    MousePointer = vbHourglass

    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox "Prepare la Impresora !!!", vbInformation, "Mensaje ..."
        
    Select Case Index
    
        Case 0
            With Mantenimiento.rsc1
                If .State = 1 Then .Close
                
                .Source = rsArticulos.Source
                
                If .State = 0 Then .Open
                .Close
                .Open

            End With
            
            With drListadoArticulosSP
                .Show
            End With
            
        Case 1
            'If Val(cbonLista.Text) = 0 Then
                MousePointer = vbDefault
                MsgBox "Debe Seleccionar una Lista de Precio", vbOKOnly, "Informacion"
                Exit Sub
            'End If
        
            With Mantenimiento.rsc1
                If .State = 1 Then .Close
                
                .Source = rsArticulos.Source
                
                If .State = 0 Then .Open
                .Close
                .Open

            End With
            
            With drListadoArticulosCP
                .Sections("TituloEmpresa").Controls("lblTitulo").Caption = "Listado de Articulos - Lista Nº: " & Val(1)
                .Sections(3).Controls("Pventa").DataField = "Pventa" & Val(1)
                .Show
            End With
            
        Case 2
            With Mantenimiento.rslistas
                If .State = 1 Then .Close
                
                .Source = rsArticulos.Source
                
                If .State = 0 Then .Open
                .Close
                .Open
            
            End With
            
            With drListasDePrecio
                .Show
            End With
                
    
    End Select
        
        
    '
    '   Setear el DataField de un campo para Imprimir la Lista que quiero
    '

    MousePointer = vbDefault

    If Err Then GrabarLog "cmdImprimir_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub ImprimirEspeciales()
    On Error Resume Next
    
    If MsgBox("¿ Desea Imprimir Solo los Clientes Filtrados Previamente ?", vbInformation + vbYesNo) = vbYes Then
    
        ImprimirPrecioEspecial "uno_solo"
    
    Else
    
        ImprimirPrecioEspecial "Todos"

    End If
    
    If Err Then GrabarLog "cmdACImprimirTodos_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub ImprimirListas()
On Error Resume Next
        
    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox "   Prepare la Impresora   ", vbInformation, "Mensaje ..."
    
    With Mantenimiento.rslistas
        If Not .State = 1 Then .Open
        .Close
        .Open
        
        .Filter = ("Codigo >= 0")
        .Sort = "Codigo ASC, Descrip ASC"
    End With
    
    With drListasDePrecio
        .Show
    End With

If Err Then GrabarLog "ImprimirListas", Err.Number & " " & Err.Description, Me.Name
End Sub


Private Sub cmdCommand1_Click()
 
' Set dgArticulos.DataSource = rsArticulos
' dgArticulos.Refresh
 Set dgArticulos.DataSource = rsArticulos
 
'Set dgArticulos.DataSource = Nothing
End Sub

Private Sub dgArticulos_AfterColUpdate(ByVal ColIndex As Integer)
'Dim vsql As String

'Dim vp1, vp2 As Double
'Dim vcodigo As String
'dgArticulos.Col = 1
'vcodigo = dgArticulos.Text

'dgArticulos.Col = 16
'vp1 = dgArticulos.Text

'dgArticulos.Col = 17
'vp2 = dgArticulos.Text


'If (ColIndex = 17 Or ColIndex = 16) And Me.chkPrecio.Value = 1 Then
'    vsql = "update articulos set pventa1=" + Str(vp1) + ", pventa2=" + Str(vp2) + " where codigo = '" + Trim(vcodigo) + "'"
'    Call EjecutarScript(vsql, pathDBMySQL)
'End If

End Sub

Private Sub dgArticulos_DblClick()
On Error Resume Next

        If Not (rsArticulos.EOF = True) And Not (rsArticulos.BOF = True) Then
            With frmArticulosAlta
                .Show
                .ModificarArticulo (rsArticulos.Fields("idArticulos").Value)
                .vaccion = "Modificar"
            End With
        End If

If Err Then GrabarLog "dgArticulos_DblClick", Err.Number & "-" & Err.Description, Me.Name
End Sub
Private Sub dgArticulos_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next

    Call OrdenarDataGrid(ColIndex, rsArticulos, dgArticulos)

    If Err Then GrabarLog "dgArticulos_HeadClick", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CommandBars_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error Resume Next

    Select Case control.Index
                
        Case 1
            With frmArticulosAlta
                .Show
                .vaccion = "Nuevo"
                .Caption = "Agregar artículo"
            End With
            
        Case 2
            If Not (rsArticulos.EOF = True) And Not (rsArticulos.BOF = True) Then
                With frmArticulosAlta
                    .Show
                    '.ModificarArticulo (rsArticulos.Fields("idArticulos").Value)
                    .ModificarArticulo (vidArticulos)

                    .vaccion = "Modificar"
                End With
            End If
            
        Case 3
            'Duplicar
        
        'frmArticulos
        Case 4
            Dim vRespuestaBorrado As Integer
            
            With rsArticulos
                If Not (.EOF = True) And Not (.BOF = True) Then
                    
                    vRespuestaBorrado = MsgBox("Esta Seguro de borrar este registro?", vbInformation + vbYesNoCancel, "Mensaje ...")
                    
                    Select Case vRespuestaBorrado
                        Case 2
                            'No hago nada puso Cancel
                        
                        Case 6
                            Call BorrarBase(vConfigGral.vempresa & ".Articulos WHERE (idArticulos = " & vidArticulos & ")", pathDBMySQL)
                            'Call BorrarBase(vConfigGral.vempresa & ".Articulosclientes WHERE (CodigoArticulo = " & .Fields("Codigo").Value & ")", pathDBMySQL)
                            'Call BorrarBase(vConfigGral.vempresa & ".ArticulosProveedorPrecio WHERE (CodigoArticulo = " & .Fields("Codigo").Value & ")", pathDBMySQL)
                            'Call BorrarBase(vConfigGral.vempresa & ".FacturaAutomaticaDetalle WHERE (CodigoArticulo = " & .Fields("Codigo").Value & ")", pathDBMySQL)
                            Buscar ("")
                        
                        Case 7
                            'No hago nada, puso NO
                    End Select
                End If
            
            End With
        
        Case 5
            
            If MsgBox("Esta Seguro de borrar todos estos registros?", vbInformation + vbYesNo, "Mensaje ...") = vbNo Then
                Exit Sub
            End If
            
            With rsArticulos
                
                If Not .EOF = True Then .MoveFirst
                
                Do Until .EOF = True
                    'Call BorrarBase(vConfigGral.vEmpresa & ".Articulosclientes WHERE (CodigoArticulo = " & .Fields("Codigo").Value & ")", pathDBMySQL)
                    'Call BorrarBase(vConfigGral.vEmpresa & ".ArticulosProveedorPrecio WHERE (CodigoArticulo = " & .Fields("Codigo").Value & ")", pathDBMySQL)
                    'Call BorrarBase(vConfigGral.vEmpresa & ".FacturaAutomaticaDetalle WHERE (CodigoArticulo = " & .Fields("Codigo").Value & ")", pathDBMySQL)
                    Call BorrarBase(vConfigGral.vempresa & ".Articulos WHERE (idArticulos = " & .Fields("idArticulos").Value & ")", pathDBMySQL)
                    
                    .MoveNext
                Loop
                
            
            End With
            RBTipoCarga(0).Value = True
            Buscar ("")
        
        Case 6
            Buscar ("")
                
        Case 7
            vVieneImpresion = Me.Name
            frmImprimir.Show
            
        Case 8
            frmActualizacionPrecio.txtFiltro(0).Text = txtBuscar.Text
            frmActualizacionPrecio.Show

        Case 9
            Unload Me

            
    End Select

If Err Then GrabarLog "CommandBars_Execute", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub dgArticulos_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyInsert Then

                        Dim vValor, vsql, vcodigo   As String
                        Dim i As Integer
                        
                        
                        i = Me.dgArticulos.Col
                        vValor = InputBox("Nuevo valor:")
                        
                        If i = 17 Then
                        
                        Me.dgArticulos.Col = 1
                        vcodigo = Me.dgArticulos.Text
                        
                            vsql = "update articulos set pventa2=" + Str(vValor) + " where codigo = '" + Trim(vcodigo) + "'"
                            Call EjecutarScript(vsql, pathDBMySQL)
                        '    Me.dgArticulos.Col = 17
                        '    Me.dgArticulos.Text = vValor
                        
                        End If
                        
                        If i = 16 Then
                        
                        Me.dgArticulos.Col = 1
                        vcodigo = Me.dgArticulos.Text
                        
                          vsql = "update articulos set pventa1=" + Str(vValor) + " where codigo = '" + Trim(vcodigo) + "'"
                          Call EjecutarScript(vsql, pathDBMySQL)
                        '    Me.dgArticulos.Col = 16
                        '    Me.dgArticulos.Text = vValor
                        
                        End If
                        
                        
                        
                            'Me.dgArticulos.Col = 1
                            'Me.dgArticulos.Text = vcodigo
                            
                            'Me.dgArticulos.Col = i
                        
                        Call txtBuscar_Change
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim vValor As String


End Sub

Private Sub Form_Load()
    On Error Resume Next

    Call init

    CargarBotonera
    'CargarPorcentajes

    With Me
        .Show
        .Top = 0
        .Left = 0
        .KeyPreview = True
        .Width = 19710
        .Height = 8190
        
        '.PicInferior.Top = -45
        '.PicInferior.Left = 4000 + 285
    End With
    
    TabArticulos.Selected = 0
    
    
    'Call Buscar("")
    
    txtBuscar.SetFocus
    
    'RBTipoCarga(2).Value = True
    
    CentrarFormulario Me
    
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
    AddControl ToolBar.Controls, xtpControlButton, 13, "&Borrar Seleccion", False, ""
    ToolBar.Closeable = True
    AddControl ToolBar.Controls, xtpControlButton, 14, "Bu&scar", True, ""
    AddControl ToolBar.Controls, xtpControlButton, 27, "&Imprimir", False, ""
    AddControl ToolBar.Controls, xtpControlButton, 12, "&Actualizar", False, "Actualiza Precios de Articulos de Manera Masiva"
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
Private Sub ImprimirPrecioEspecial(sql_tipo As String)
    On Error Resume Next

    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox "   Prepare la Impresora   ", vbInformation, "Mensaje ..."

    With Mantenimiento.rslart_cliente
    
        If Not .State = 1 Then .Open
        .Close
        .Open
    
        'If Not sql = "Todos" Then
            '.filter = ("ID >0") + sql
        '    .Sort = "Cliente ASC"
        'Else
            .Filter = ("id > 0")
            .Sort = "Articulo ASC, Cliente ASC"
        'End If
    
    End With

    With drArticuloCliente
        .Show
    End With
    
    If Err Then GrabarLog "ImprimirPrecioEspecial", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub g_Click()
On Error Resume Next
'If UCase(LeerXml("")) = "PONS" Then
            vidArticulos = g.TextMatrix(g.Row, 14)
            
            Call pintar(g.Row)
            'vulinea = g.Row
            Call despintar(vulinea)
            vulinea = g.Row
            g.CellBackColor = vbRed
'End If

'vidArticulos = g.TextMatrix(g.Row, 14)
'vulinea = g.Row

If Err Then Exit Sub
End Sub



Private Sub despintar(i As Integer)
On Error Resume Next

Dim j, k, kk As Integer
k = g.Row
kk = g.Col
If i = 0 Then Exit Sub
g.Row = i

For j = 1 To g.Cols - 1
    g.Col = j
    g.CellBackColor = vbWhite
Next

g.Row = k
g.Col = kk

If Err Then Exit Sub
End Sub
Private Sub pintar(i As Integer)
On Error Resume Next
Dim j, k, kk As Integer

k = g.Row
kk = g.Col

g.Row = i

For j = 1 To g.Cols - 1
    g.Col = j
    g.CellBackColor = vbGreen
Next

g.Row = k
g.Col = kk
If Err Then Exit Sub
End Sub



Private Sub g_EnterCell()
Call g_Click
End Sub

Private Sub g_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 43 Then


With g



                        Dim vValor, vsql, vcodigo   As String
                        Dim i As Integer
                        
                        
                        i = .Col
                         
                        If (i - 8 = 1 Or i - 8 = 2) Then
                         vValor = InputBox("Nuevo valor:")
                         
                         
                         'Me.dgArticulos.Col = 1
                         vcodigo = .TextMatrix(.Row, 1)
                         
                             vsql = "update articulos set pventa" + Trim(Str(i - 8)) + "=" + Str(vValor) + " where codigo = '" + Trim(vcodigo) + "'"
                        
                             Call EjecutarScript(vsql, pathDBMySQL)
                         '    Me.dgArticulos.Col = 17
                             .TextMatrix(.Row, i) = vValor
                             .CellBackColor = vbGreen
                        End If
                        
                        
                        
                        If i = 12 Then
                             vcodigo = .TextMatrix(.Row, 1)
                             vValor = InputBox("Nuevo valor:")
                         
                             vsql = "update articulos set stock=" + Str(vValor) + " where codigo = '" + Trim(vcodigo) + "'"
                        
                             Call EjecutarScript(vsql, pathDBMySQL)
                         '    Me.dgArticulos.Col = 17
                             .TextMatrix(.Row, i) = vValor
                             .CellBackColor = vbGreen
                        
                        End If
                        
                        
                      '  Call txtBuscar_Change
                        
End With

End If

Me.g.SetFocus


If Err Then Exit Sub
End Sub

Private Sub g_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If this is not row 0, do nothing.
    If g.MouseRow <> 0 Then Exit Sub

    ' Sort by the clicked column.
    Call SortByColumn(g.MouseCol, g)
End Sub

Private Sub PusConsultarTarifaria_Click()
    MsgBox consultarTarifaria(rsArticulos.Fields("idArticulos").Value), vbInformation, "Tarifarias:"
End Sub

Private Sub PusCSV_Click()

End Sub

Private Sub PusExcel_Click()
On Error Resume Next
    
Call Buscar("", "todos")
  
  
Call grillaToExcel(Me.g)

If Err Then Exit Sub
End Sub

Private Sub PushButton1_Click()
On Error Resume Next

Dim vsaldo As Double
Dim vcodarticulo As String
                        
Dim rsArticulos As New ADODB.Recordset
Dim sqlArticulos As String
    
    sqlArticulos = "SELECT codigo FROM articulos"
    
    With rsArticulos
        Call .Open(sqlArticulos, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not .EOF = True Then .MoveFirst
        
        Do Until .EOF = True
           
           vcodarticulo = .Fields("codigo")
           
           vsaldo = Val(Format(GenerarDato("SELECT Sum(Entrada), Sum(Salida), Sum(Entrada-Salida) AS SaldoActual FROM Stock WHERE  CodigoArticulo = '" & vcodarticulo & "'", "SaldoActual"), "#####0.00"))
           
           Call actualizastockEnArticulo(EsNulo(vcodarticulo), vsaldo)

            .MoveNext
        Loop
    
    End With
        
    sqlArticulos = ""
    
    rsArticulos.Close
    Set rsArticulos = Nothing

If Err Then GrabarLog "ActualizarRubros", Err.Number & " " & Err.Description, "Procedimientos"

End Sub

Private Sub PushButton2_Click()
frmMonitor.Show
End Sub

Private Sub PushButton3_Click()

If Me.chkChkTodos.Value = xtpChecked Then
    Call Buscar("", "Todos")
Else
    Call Buscar("", "Todos")
End If


End Sub

Private Sub RBTipoCarga_Click(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 0
            Call Buscar(" AND NOT (Observaciones = 'CargadoPorRemito')")
        
        Case 1
            Call Buscar(" AND (Observaciones = 'CargadoPorRemito')")
            
        Case 2
            Call Buscar("")
    
    End Select
    

If Err Then GrabarLog "RBTipoCarga_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub TabArticulos_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error Resume Next

    Select Case Item.Index
    
        Case 0
        
        Case 1
        
        Case 2
        
        Case 3
            'Por ahora no se usa
    End Select

If Err Then GrabarLog "TabArticulos_SelectedChanged", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtBuscar_Change()
    
    If UCase(LeerXml("Cliente")) = "HERNAN" Or True Then Buscar ("")

End Sub

Private Sub txtrubro_Click()
Dim vsql As String

vsql = "(select * from rubros) as t"

Call fbuscarGrilla(vsql, "Rubro", "idRubros", txtrubro.Name, Me)  ' ema:


End Sub

Private Sub txtsubrubro_click()
Dim vsql As String

vsql = "(select * from rubroes) as t"

Call fbuscarGrilla(vsql, "Subrubro", "idSubRubros", txtsubrubro.Name, Me)   ' ema:

End Sub
