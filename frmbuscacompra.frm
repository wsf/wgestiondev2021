VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "Copia de KlexGrid.ocx"
Begin VB.Form frmBuscarCompra 
   Caption         =   "Listado de Documentos de Compras"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6870
   ScaleWidth      =   11595
   Begin MSAdodcLib.Adodc bFactura 
      Height          =   330
      Left            =   5400
      Top             =   5160
      Visible         =   0   'False
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   582
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
      Caption         =   "bPFactura"
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
   Begin VB.Frame Frame2 
      Height          =   3675
      Left            =   90
      TabIndex        =   40
      Top             =   2520
      Width           =   11295
      Begin MSDataGridLib.DataGrid dgPFactura 
         Bindings        =   "frmbuscacompra.frx":0000
         Height          =   765
         Left            =   30
         TabIndex        =   41
         Top             =   120
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   1349
         _Version        =   393216
         HeadLines       =   1.2
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
      Begin Grid.KlexGrid KlexDocumentos 
         Height          =   2175
         Left            =   90
         TabIndex        =   46
         Top             =   1260
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   3836
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
         MouseIcon       =   "frmbuscacompra.frx":0018
         Rows            =   10
      End
   End
   Begin VB.PictureBox PicInferior 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      Picture         =   "frmbuscacompra.frx":0034
      ScaleHeight     =   555
      ScaleWidth      =   11550
      TabIndex        =   25
      Top             =   6240
      Width           =   11550
      Begin XtremeSuiteControls.PushButton cmdCerrar 
         Height          =   375
         Left            =   10080
         TabIndex        =   26
         Top             =   90
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cerrar"
         Appearance      =   6
         Picture         =   "frmbuscacompra.frx":50E7
      End
      Begin XtremeSuiteControls.PushButton CmdVerDetalle 
         Height          =   375
         Left            =   2160
         TabIndex        =   27
         Top             =   90
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Ver detalle"
         Appearance      =   6
         Picture         =   "frmbuscacompra.frx":54E7
      End
      Begin XtremeSuiteControls.PushButton CmdBorrar 
         Height          =   375
         Left            =   3480
         TabIndex        =   28
         Top             =   90
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Borrar"
         Enabled         =   0   'False
         Appearance      =   6
         Picture         =   "frmbuscacompra.frx":58EE
      End
      Begin XtremeSuiteControls.PushButton cmdImprimir 
         Height          =   375
         Index           =   0
         Left            =   6120
         TabIndex        =   29
         Top             =   90
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Listado"
         Appearance      =   6
         Picture         =   "frmbuscacompra.frx":5D0C
      End
      Begin XtremeSuiteControls.PushButton cmdAnularFactura 
         Height          =   375
         Left            =   7440
         TabIndex        =   30
         Top             =   90
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Anular Doc."
         Enabled         =   0   'False
         Appearance      =   6
         Picture         =   "frmbuscacompra.frx":60ED
      End
      Begin XtremeSuiteControls.PushButton cmdEjecutarPago 
         Height          =   375
         Left            =   8760
         TabIndex        =   31
         Top             =   90
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Ejec. Pago"
         Enabled         =   0   'False
         Appearance      =   6
         Picture         =   "frmbuscacompra.frx":64CA
      End
      Begin XtremeSuiteControls.PushButton cmdImprimir 
         Height          =   375
         Index           =   1
         Left            =   4800
         TabIndex        =   32
         Top             =   90
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Factura"
         Transparent     =   -1  'True
         Appearance      =   6
         Picture         =   "frmbuscacompra.frx":68D4
      End
      Begin VB.Label lblWGESTION2010 
         BackStyle       =   0  'Transparent
         Caption         =   "WGESTION 2010"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Index           =   0
         Left            =   50
         TabIndex        =   33
         Top             =   150
         Width           =   1770
      End
      Begin VB.Label lblWGESTION2010 
         BackStyle       =   0  'Transparent
         Caption         =   "WGESTION 2010"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Index           =   1
         Left            =   75
         TabIndex        =   34
         Top             =   170
         Width           =   1770
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2505
      Left            =   30
      ScaleHeight     =   2505
      ScaleWidth      =   11505
      TabIndex        =   0
      Top             =   0
      Width           =   11505
      Begin VB.Frame Frame1 
         Height          =   585
         Left            =   90
         TabIndex        =   35
         Top             =   -30
         Width           =   6345
         Begin XtremeSuiteControls.FlatEdit txtProveedor 
            Height          =   315
            Index           =   0
            Left            =   1170
            TabIndex        =   36
            Top             =   210
            Width           =   1215
            _Version        =   851968
            _ExtentX        =   2143
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit txtProveedor 
            Height          =   315
            Index           =   1
            Left            =   2790
            TabIndex        =   37
            Top             =   210
            Width           =   3495
            _Version        =   851968
            _ExtentX        =   6174
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Locked          =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton pbCarga 
            Height          =   315
            Index           =   0
            Left            =   2430
            TabIndex        =   38
            Tag             =   "Proveedor"
            Top             =   210
            Width           =   315
            _Version        =   851968
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label lblIngEl 
            Caption         =   "> Proveedor:"
            Height          =   195
            Left            =   60
            TabIndex        =   39
            Top             =   270
            Width           =   945
         End
      End
      Begin VB.Frame fraNComprobante 
         Caption         =   "Nº de Comprobante"
         ForeColor       =   &H00808080&
         Height          =   585
         Left            =   90
         TabIndex        =   20
         Top             =   1350
         Width           =   5265
         Begin VB.TextBox nhasta 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3450
            TabIndex        =   22
            Top             =   180
            Width           =   1755
         End
         Begin VB.TextBox ndesde 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   870
            TabIndex        =   21
            Top             =   180
            Width           =   1845
         End
         Begin VB.Label Label10 
            Caption         =   "Hasta :"
            Height          =   285
            Left            =   2850
            TabIndex        =   24
            Top             =   210
            Width           =   615
         End
         Begin VB.Label Label11 
            Caption         =   "Desde :"
            Height          =   165
            Left            =   210
            TabIndex        =   23
            Top             =   210
            Width           =   555
         End
      End
      Begin VB.Frame fraFechas 
         ForeColor       =   &H00808080&
         Height          =   615
         Left            =   90
         TabIndex        =   14
         Top             =   690
         Width           =   6345
         Begin VB.CheckBox chkFecha 
            Caption         =   "Todas"
            Height          =   255
            Left            =   5430
            TabIndex        =   15
            Top             =   240
            Value           =   1  'Checked
            Width           =   765
         End
         Begin MSComCtl2.DTPicker dtpDesde 
            Height          =   285
            Left            =   810
            TabIndex        =   16
            Top             =   210
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   75890689
            CurrentDate     =   39234
         End
         Begin MSComCtl2.DTPicker dtpHasta 
            Height          =   285
            Left            =   3420
            TabIndex        =   17
            Top             =   210
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   75890689
            CurrentDate     =   39234
         End
         Begin VB.Label Label2 
            Caption         =   "Desde :"
            Height          =   225
            Left            =   90
            TabIndex        =   19
            Top             =   240
            Width           =   675
         End
         Begin VB.Label Label3 
            Caption         =   "Hasta :"
            Height          =   255
            Left            =   2790
            TabIndex        =   18
            Top             =   270
            Width           =   735
         End
      End
      Begin VB.Frame fraTipo 
         Caption         =   "Tipos de documento"
         ForeColor       =   &H00404040&
         Height          =   1335
         Left            =   6450
         TabIndex        =   4
         Top             =   -30
         Width           =   4875
         Begin VB.CheckBox chkNotaCA 
            Caption         =   "Notas de Crédito A"
            Height          =   255
            Left            =   3120
            TabIndex        =   6
            Top             =   300
            Value           =   1  'Checked
            Width           =   1725
         End
         Begin VB.CheckBox chkNotaCB 
            Caption         =   "Notas de Crédito B"
            Height          =   255
            Left            =   3120
            TabIndex        =   5
            Top             =   690
            Width           =   1725
         End
         Begin VB.CheckBox chkPresupuesto 
            Caption         =   "Presupuesto"
            Height          =   255
            Left            =   1530
            TabIndex        =   12
            Top             =   630
            Width           =   1755
         End
         Begin VB.CheckBox chkFacturaC 
            Caption         =   "Factura C"
            Height          =   255
            Left            =   90
            TabIndex        =   11
            Top             =   990
            Width           =   1395
         End
         Begin VB.CheckBox chkRemito 
            Caption         =   "Remito"
            Height          =   255
            Left            =   1530
            TabIndex        =   10
            Top             =   960
            Width           =   1245
         End
         Begin VB.CheckBox chkNotasDe 
            Caption         =   "Notas de Débito"
            Height          =   255
            Left            =   1530
            TabIndex        =   9
            Top             =   300
            Width           =   1455
         End
         Begin VB.CheckBox chkFacturaB 
            Caption         =   "Factura B"
            Height          =   255
            Left            =   90
            TabIndex        =   8
            Top             =   690
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkFacturaA 
            Caption         =   "Factura A"
            Height          =   255
            Left            =   90
            TabIndex        =   7
            Top             =   330
            Value           =   1  'Checked
            Width           =   1245
         End
      End
      Begin VB.Frame fraImpresion 
         Height          =   495
         Left            =   5430
         TabIndex        =   1
         Top             =   1380
         Width           =   2115
         Begin VB.CheckBox chkNoImpreso 
            Caption         =   "No Impreso"
            Height          =   225
            Left            =   30
            TabIndex        =   3
            Top             =   180
            Value           =   1  'Checked
            Width           =   1125
         End
         Begin VB.CheckBox chkImpreso 
            Caption         =   "Impreso"
            Height          =   225
            Left            =   1170
            TabIndex        =   2
            Top             =   180
            Value           =   1  'Checked
            Width           =   915
         End
      End
      Begin XtremeSuiteControls.PushButton cmdBuscaryCalcular 
         Height          =   405
         Left            =   120
         TabIndex        =   13
         Top             =   1980
         Width           =   11205
         _Version        =   851968
         _ExtentX        =   19764
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Filtrar Datos"
         Appearance      =   6
         Picture         =   "frmbuscacompra.frx":6CEF
      End
      Begin XtremeSuiteControls.GroupBox GroAgrupadoPor 
         Height          =   525
         Left            =   7590
         TabIndex        =   42
         Top             =   1350
         Width           =   3765
         _Version        =   851968
         _ExtentX        =   6641
         _ExtentY        =   926
         _StockProps     =   79
         Caption         =   "Agrupado por:"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton RadNoAgrupado 
            Height          =   225
            Left            =   60
            TabIndex        =   43
            Top             =   240
            Width           =   1275
            _Version        =   851968
            _ExtentX        =   2249
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "No agrupado"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RadProvincias 
            Height          =   225
            Left            =   1380
            TabIndex        =   44
            Top             =   240
            Width           =   1005
            _Version        =   851968
            _ExtentX        =   1773
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "Provincias"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RadPersona 
            Height          =   225
            Left            =   2400
            TabIndex        =   45
            Top             =   240
            Width           =   1245
            _Version        =   851968
            _ExtentX        =   2196
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "Razón Social"
            UseVisualStyle  =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frmBuscarCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vsql As String
Public vremito, vViene  As String
Public vienePago As Boolean
Private Sub cmdBorrar_Click()
On Error Resume Next

    Dim vRemitoABorrar As Integer, vDocumentoABorrar As String, vnroasiento As Long

    With bfactura
        
        If (.Recordset.EOF = True) Or (.Recordset.BOF = True) Then
            MsgBox "Debe seleccionar un Documento/Factura para borrar!!", vbExclamation, "Mensaje ..."
            Exit Sub
        End If
        
        If MsgBox("Esta seguro que desea borrar la factura Nº: " & EsNulo(.Recordset("NComp").Value), vbInformation + vbYesNo, "Mensaje ...") = vbNo Then Exit Sub
        vnroasiento = Val(.Recordset("NroAsiento").Value)
        vRemitoABorrar = Val(.Recordset("Remito").Value)
        vDocumentoABorrar = EsNulo(.Recordset("Tipo").Value)
        
        Call BorrarBase("PFactura WHERE (Remito = " & Val(vRemitoABorrar) & ")", pathDBMySQL)
        
        
        .Refresh
        If Not .Recordset.EOF = True Then .Recordset.MoveLast
        
        FormatoGrilla
        
    End With
    
    If Val(vRemitoABorrar) = 0 Then Exit Sub
    
    If (vDocumentoABorrar = "Fact A") Or (vDocumentoABorrar = "Fact B") Or (vDocumentoABorrar = "Fact C") Or (vDocumentoABorrar = "Nota C") Then
        Call BorrarBase("IvaFacturaCompra WHERE (remito = " & vRemitoABorrar & ")", pathDBMySQL)
    End If
    
    Call BorrarBase("PCuentascorrientes WHERE (remito = " & vRemitoABorrar & ")", pathDBMySQL)
    Call BorrarBase("Caja WHERE (remito = " & vRemitoABorrar & ")", pathDBMySQL)
    
    '---------------- Borro el Asiento ---------------------------------------
    Call BorrarBase(" Asientos WHERE (Numero = " & Val(vnroasiento) & ")", pathDBMySQL)

    '---------------- Borro el Asiento ---------------------------------------
    Call BorrarBase(" AsientosDetalle WHERE (Numero = " & Val(vnroasiento) & ")", pathDBMySQL)
    
    Dim rsPFDetalle As New ADODB.Recordset, sqlPFDetalle As String
    
    sqlPFDetalle = "SELECT * FROM pfdetalle WHERE (remito = " & vRemitoABorrar & ")"
    
    With rsPFDetalle
        .CursorLocation = adUseClient
        Call .Open(sqlPFDetalle, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        Do Until .EOF = True
            If vDocumentoABorrar = "Nota C" Then
                'Call ModificarStock(1, .Fields("cantidad").Value, .Fields("codigo").Value)
            Else
                If Not vDocumentoABorrar = "Presupuesto" Then
                    'Call ModificarStock(-1, .Fields("cantidad").Value, .Fields("codigo").Value)
                End If
            End If
            
            .MoveNext
        Loop
        
        Call BorrarBase("pfdetalle WHERE (remito = " & vRemitoABorrar & ")", pathDBMySQL)
    
    End With
    
    sqlPFDetalle = ""
    
    rsPFDetalle.Close
    Set rsPFDetalle = Nothing
    
    
    
If Err Then GrabarLog "cmdBorrar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub cmdBuscaryCalcular_Click()
On Error Resume Next
    
    vsql = ""
    
    If Not Trim(txtProveedor(0).Text) = "" Then vsql = vsql + " AND (Codigo = '" & Trim(txtProveedor(0).Text) & "')"

    If Not (Trim(ndesde.Text) = "") And Not (Trim(nhasta.Text) = "") Then
        vsql = vsql + " AND (NComp >= '" & ndesde.Text & "' AND NComp <= '" & Trim(nhasta.Text) & "')"
    End If

    vsql = vsql + " and (1<1 "
    If chkFacturaA.Value = 1 Then vsql = vsql + " OR (tipo = 'Fact A')"
    If chkFacturaB.Value = 1 Then vsql = vsql + " OR (tipo = 'Fact B')"
    If chkFacturaC.Value = 1 Then vsql = vsql + " OR (tipo = 'Fact C')"
    If chkNotasDe.Value = 1 Then vsql = vsql + " OR (Tipo = 'Nota D')"
    If chkRemito.Value = 1 Then vsql = vsql + " OR (Tipo = 'Remito')"
    'If chkdocno.value = 1 Then vSQL = vSQL + " OR (Tipo = 'Documento')"
    If chkPresupuesto.Value = 1 Then vsql = vsql + " OR (Tipo = 'Presupuesto')"
    If chkNotaCA.Value = 1 Then vsql = vsql + " OR (tipo = 'Nota C' AND (Letra = 'A'))"
    If chkNotaCB.Value = 1 Then vsql = vsql + " OR (tipo = 'Nota C' AND (Letra = 'B'))"
    vsql = vsql + ")"
    
    If dtpDesde.Enabled = True And dtpHasta.Enabled = True Then
        vsql = vsql + " and Fecha >= '" & strfechaMySQL(dtpDesde.Value) + "' and fecha <= '" & strfechaMySQL(dtpHasta.Value) + "'"
    End If

    With bfactura
        .ConnectionString = pathDBMySQL
        
        If Me.RadProvincias Then
        
            .RecordSource = "SELECT  proveedores.Provincia, sum(pfactura.Total) AS TotalFacturado From  ivafacturacompra t INNER JOIN pfactura ON (t.Remito=pfactura.Remito) INNER JOIN proveedores ON (proveedores.Codigo=pfactura.Codigo) where 1=1 " + vsql + " Group By  proveedores.Provincia"
            .Refresh
            Set dgPFactura.DataSource = .Recordset
            
            Set KlexDocumentos.Recordset = .Recordset
        
        Else
        
                .RecordSource = "SELECT * FROM DocCompras WHERE (1=1" & vsql & ") ORDER BY idPFactura ASC"
                .Refresh
                If Not .Recordset.EOF = True Then .Recordset.MoveLast
                FormatoGrilla
                Set KlexDocumentos.Recordset = .Recordset
    
        End If
    
    End With

    

If Err Then GrabarLog "cmdBuscar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub ImprimirListadoDeFactura()
On Error Resume Next
    
    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox "Prepare la Impresora!!!", vbInformation, "Mensaje ..."
  
    With Mantenimiento.rsldoccompras
        If .State = 1 Then .Close
        
        .Source = bfactura.RecordSource
        
        If .State = 0 Then .Open
        .Close
        .Open
    
  
    End With
  
    With drListadoFacturaCompra
        .Show
    End With
    
If Err Then GrabarLog "ImprimirListadoDeFactura", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Seleccionar()
    On Error Resume Next
    
    With bfactura
        If (.Recordset.EOF = True) Or (.Recordset.BOF = True) Then
            MsgBox "Debe elegir un Documento para poder modificarlo!!!", vbInformation, "Mensaje ..."
            Exit Sub
        End If
    End With
    
    MousePointer = vbHourglass
    
    With frmCompras
        .vGrabaModo = 1
        .CargarRemito (bfactura.Recordset("remito").Value)
        
        If vPFDetalle = True Then .f(0).SetFocus
            
        MousePointer = vbDefault
        
    End With
    
    Unload Me

    If Err Then GrabarLog "Seleccionar", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CargarDocumentos()
    On Error Resume Next
    
    Dim i As Integer
    
    With bfactura
        If Not .Recordset.EOF = True Then .Recordset.MoveFirst
            
       ' FormatoGrilla (.Recordset.RecordCount)
        
        i = 1
        Do Until .Recordset.EOF = True
        
            KlexDocumentos.TextMatrix(i, 0) = ""
            KlexDocumentos.TextMatrix(i, 1) = EsNulo(.Recordset("idFactura").Value)
            KlexDocumentos.TextMatrix(i, 2) = EsNulo(.Recordset("TipoMovimiento").Value)
            KlexDocumentos.TextMatrix(i, 3) = EsNulo(.Recordset("Letra").Value)
            KlexDocumentos.TextMatrix(i, 4) = EsNulo(.Recordset("PuntoDeVenta").Value)
            KlexDocumentos.TextMatrix(i, 5) = EsNulo(.Recordset("NComprobante").Value)
            KlexDocumentos.TextMatrix(i, 6) = EsNulo(.Recordset("Fecha").Value)
            KlexDocumentos.TextMatrix(i, 7) = EsNulo(.Recordset("Codigo").Value)
            KlexDocumentos.TextMatrix(i, 8) = EsNulo(.Recordset("Nombre").Value)
            KlexDocumentos.TextMatrix(i, 9) = EsNulo(.Recordset("Cuit").Value)
            KlexDocumentos.TextMatrix(i, 10) = Format(.Recordset("Total").Value, "######0.00")
            KlexDocumentos.TextMatrix(i, 11) = EsNulo(.Recordset("NroInterno").Value)
            KlexDocumentos.TextMatrix(i, 12) = EsNulo(.Recordset("Comentario").Value)
            KlexDocumentos.TextMatrix(i, 13) = EsNulo(.Recordset("Remito").Value)
            KlexDocumentos.TextMatrix(i, 14) = EsNulo(.Recordset("NroAsiento").Value)
            
            'klexDocumentos.TextMatrix(i, 13) = EsNulo(.Recordset("Endoso").Value)
            
            'klexDocumentos.TextMatrix(i, 15) = EsNulo(.Recordset("FechaAcreditacion").Value)
            
            'klexDocumentos.TextMatrix(i, 17) = EsNulo(.Recordset("NroInterno").Value)
            'klexDocumentos.TextMatrix(i, 18) = EsNulo(.Recordset("Observaciones").Value)
            .Recordset.MoveNext
        
            i = i + 1
        Loop
        
        KlexDocumentos.TopRow = i - 1
        .Refresh
        
        
        'If .Recordset.EOF = True Then .Recordset.MoveLast
    End With
    
If Err Then GrabarLog "CargarDocumentos", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub chkFecha_Click()
    On Error Resume Next
    
    dtpDesde.Enabled = CBool(chkFecha.Value - 1)
    dtpHasta.Enabled = CBool(chkFecha.Value - 1)

    If Err Then GrabarLog "chkFecha_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdCerrar_Click()
On Error Resume Next

    Unload Me
    
If Err Then GrabarLog "cmdCerrar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdEjecutarPago_Click()
On Error Resume Next

    If (bfactura.Recordset.EOF > True) Then
        With frmPagos
            .NroComprobante = EsNulo(bfactura.Recordset("NComp").Value)
            .tipoComprobante = EsNulo(bfactura.Recordset("Tipo").Value)
            .fechaDocumento = EsNulo(bfactura.Recordset("Fecha").Value)
            .remito = EsNulo(bfactura.Recordset("remito").Value)
            .codProveedor = EsNulo(bfactura.Recordset("codigo").Value)
           
            'si el cliente no coincide se debe alertar y no continuar el proceso
            If (.codProveedor <> EsNulo(bfactura.Recordset("codigo").Value)) And (.codProveedor <> "") Then
                MsgBox "El documento seleccionado corresponde a un Proveedor distinto del que se está cobrando, revise su operación."
            Else
        
                Call .BuscarDatosOperacionesProveedor(EsNulo(bfactura.Recordset("codigo").Value), EsNulo(bfactura.Recordset("remito").Value))
                .esComprobanteAutomatico = False
        
                .txtNroComprobante = EsNulo(bfactura.Recordset("NComp").Value)
                .txtTipoComp = EsNulo(bfactura.Recordset("Tipo").Value)
                            
                .Show
            End If
            
        End With
        
        Unload Me
    Else
        MsgBox "Debe seleccionar un documento de Compra", vbInformation, "WGestion"
    End If
    
If Err Then GrabarLog "cmdEjecutarPago_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 0
            ImprimirListadoDeFactura
        Case 1
            'No se
    
    End Select

If Err Then GrabarLog "", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub cmdVerDetalle_Click()
On Error Resume Next

    Seleccionar

If Err Then GrabarLog "CmdVerDetalle_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub dgPFactura_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    
    Call OrdenarDataGrid(ColIndex, bfactura.Recordset, dgPFactura)

    If Err Then GrabarLog "dgPFactura_HeadClick", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    If KeyCode = vbKeyF3 Then
        pbCarga_Click (0)
    End If

If Err Then GrabarLog "Form_KeyUp", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub Form_Load()
On Error Resume Next

    dtpDesde.Value = Date
    dtpHasta.Value = Date
    
    With Me
        .Top = 0
        .Left = 0
        .Width = 11900
        .Height = 7250
        .dtpDesde.Enabled = False
        .dtpHasta.Enabled = False
        .KeyPreview = True
    End With

    With bfactura
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        If vViene = "pctacte" Then
            .RecordSource = "SELECT idPFactura, Remito, (NComprobante) as NroComp, Fecha, Codigo, Nombre, CVenta, Tipo, SubTotal, Total FROM PFactura WHERE (remito = " & vremito & ")"
            vViene = ""
        Else
            .RecordSource = "SELECT * FROM DocCompras"
        End If
        .Refresh
    
    End With
        
    FormatoGrilla
    
    'If Not vIdUsuarioNivel = 1 Then ControlarPermisos
    
    If Err Then GrabarLog "Form_load", Err.Number & "  " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub dgPFactura_DblClick()
    On Error Resume Next

    If vienePago = True Then
        cmdEjecutarPago_Click
    Else
        Seleccionar
    End If
    
    If Err Then GrabarLog "dgPFactura_DblClick", Err.Number & "  " & Err.Description, Me.Name
End Sub
Private Sub ControlarPermisos()
On Error Resume Next

    Dim rsUsuariosPermisos As New ADODB.Recordset, sqlUsuariosPermisos As String
    
    'sqlUsuariosPermisos = "SELECT * FROM UsuariosPermisos WHERE (idUsuarios = " & vIdUsuario & ") AND (Formulario = 'frmBuscaCompra')  AND (NOT Accion IS NULL OR NOT Accion = '')"
   
    With rsUsuariosPermisos
        Call .Open(sqlUsuariosPermisos, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then .MoveFirst
        
        Do Until .EOF = True
    
            Select Case .Fields("Accion").Value
                
                Case "Nuevo"
                    
                    
                
                Case "Borrar"
                    cmdBorrar.Enabled = CBool(.Fields("Habilitado").Value)
                    'cmdAnular.Enabled = CBool(.Fields("Habilitado").value)
                    
                Case "Guardar"
                    
                
                Case "Imprimir"
                    cmdImprimir(0).Enabled = CBool(.Fields("Habilitado").Value)
                    cmdImprimir(1).Enabled = CBool(.Fields("Habilitado").Value)
                    
                Case "Buscar"
                    cmdBuscaryCalcular.Enabled = CBool(.Fields("Habilitado").Value)
                
                Case "Modificar"
                    cmdVerDetalle.Enabled = CBool(.Fields("Habilitado").Value)
                    dgPFactura.Enabled = CBool(.Fields("Habilitado").Value)
                
                    
            End Select

            .MoveNext
        Loop
    
    End With
    
    sqlUsuariosPermisos = ""
    
    If rsUsuariosPermisos.State = 1 Then
        rsUsuariosPermisos.Close
        Set rsUsuariosPermisos = Nothing
    End If

If Err Then GrabarLog "ControlarPermisos", Left(Err.Number & " " & Err.Description, 99), Me.Name
End Sub
Private Sub FormatoGrilla()
On Error Resume Next

    With dgPFactura
         .HeadLines = 2
         
         .Columns(0).Width = 0
         .Columns(1).Width = 0
         
         .Columns(2).Caption = "Nro. Comp."
         .Columns(2).Width = 1500
         
         .Columns(3).Caption = "Tipo"
         .Columns(3).Width = 1000
         
         .Columns(4).Caption = "Letra"
         .Columns(4).Width = 750
         
         .Columns(5).Caption = "Fecha"
         .Columns(5).Width = 1250
         
         .Columns(6).Width = 0
         .Columns(7).Caption = "Proveedor"
         .Columns(7).Width = 3500
         
        .Columns(8).Width = 0

        .Columns(9).Caption = "SubTotal"
        .Columns(9).Width = 1000
        .Columns(9).Alignment = dbgRight
        .Columns(9).DataFormat.Format = "$ #######0.00"
        
        .Columns(10).Caption = "Total"
        .Columns(10).Width = 1100
        .Columns(10).Alignment = dbgRight
        .Columns(10).DataFormat.Format = "$ #######0.00"
    
        .Columns(11).Width = 0

    End With

If Err Then GrabarLog "FormatearGrilla", Left(Err.Number & " " & Err.Description, 99), Me.Name
End Sub

Private Sub KlexDocumentos_DblClick()
On Error Resume Next

    If vienePago = True Then
        cmdEjecutarPago_Click
    Else
        Seleccionar
    End If
    
    If Err Then GrabarLog "dgPFactura_DblClick", Err.Number & "  " & Err.Description, Me.Name
End Sub

Private Sub pbCarga_Click(Index As Integer)
    On Error Resume Next

    vVuelveBusqueda = Me.Name
    vVieneBusqueda = pbCarga(Index).Tag

    Select Case Index

        Case 0 To 10
            frmBusqueda.Show
    
    End Select
    
    If Err Then GrabarLog "pbCarga_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub PushButton2_Click()
End Sub
