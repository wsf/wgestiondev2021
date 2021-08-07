VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "Copia de KlexGrid.ocx"
Object = "{9746E3DA-06E1-4D26-9CE4-D9F6411A9C70}#1.0#0"; "SMGA_OcxTxt2008.ocx"
Begin VB.Form frmIvaVenta 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Libro IVA Venta"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   180
   ClientWidth     =   13560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   13560
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   495
      Left            =   30
      TabIndex        =   30
      Top             =   420
      Width           =   11835
      _Version        =   851968
      _ExtentX        =   20876
      _ExtentY        =   873
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton PushButton6 
         Height          =   375
         Left            =   3210
         TabIndex        =   31
         Top             =   90
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "F6"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.FlatEdit vcodEmpresa 
         Height          =   285
         Left            =   1380
         TabIndex        =   32
         Top             =   150
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3096
         _ExtentY        =   503
         _StockProps     =   77
         ForeColor       =   4210752
         BackColor       =   -2147483633
         Appearance      =   3
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit vdescEmpresa 
         Height          =   285
         Left            =   3600
         TabIndex        =   33
         Top             =   150
         Width           =   8115
         _Version        =   851968
         _ExtentX        =   14314
         _ExtentY        =   503
         _StockProps     =   77
         ForeColor       =   4210752
         BackColor       =   -2147483633
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   285
         Left            =   120
         TabIndex        =   34
         Top             =   150
         Width           =   2295
         _Version        =   851968
         _ExtentX        =   4048
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Empresa:"
         ForeColor       =   255
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
   Begin XtremeSuiteControls.ProgressBar barra 
      Height          =   135
      Left            =   4710
      TabIndex        =   29
      Top             =   420
      Width           =   3615
      _Version        =   851968
      _ExtentX        =   6376
      _ExtentY        =   238
      _StockProps     =   93
   End
   Begin XtremeSuiteControls.ProgressBar b1 
      Height          =   135
      Left            =   60
      TabIndex        =   28
      Top             =   420
      Width           =   4575
      _Version        =   851968
      _ExtentX        =   8070
      _ExtentY        =   238
      _StockProps     =   93
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cambiar"
      Height          =   315
      Left            =   10230
      TabIndex        =   19
      Top             =   6030
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.TextBox v27 
      Height          =   315
      Left            =   8220
      TabIndex        =   18
      Text            =   "V27"
      Top             =   6030
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.TextBox v21 
      Height          =   315
      Left            =   6270
      TabIndex        =   17
      Text            =   "V21"
      Top             =   6030
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.TextBox v105 
      Height          =   315
      Left            =   4350
      TabIndex        =   16
      Text            =   "V105"
      Top             =   6030
      Visible         =   0   'False
      Width           =   1875
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   525
      Left            =   0
      TabIndex        =   9
      Top             =   -90
      Width           =   11895
      _Version        =   851968
      _ExtentX        =   20981
      _ExtentY        =   926
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.ProgressBar ProgressBar2 
         Height          =   105
         Left            =   60
         TabIndex        =   23
         Top             =   480
         Width           =   11895
         _Version        =   851968
         _ExtentX        =   20981
         _ExtentY        =   185
         _StockProps     =   93
         Text            =   "Barra"
      End
      Begin XtremeSuiteControls.ProgressBar ProgressBar1 
         Height          =   105
         Left            =   30
         TabIndex        =   22
         Top             =   480
         Width           =   11835
         _Version        =   851968
         _ExtentX        =   20876
         _ExtentY        =   185
         _StockProps     =   93
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   315
         Left            =   7770
         TabIndex        =   20
         Top             =   150
         Width           =   915
         _Version        =   851968
         _ExtentX        =   1614
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Sacar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   375
         Left            =   6120
         TabIndex        =   15
         Top             =   120
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Excel"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   375
         Index           =   3
         Left            =   10410
         TabIndex        =   10
         Top             =   120
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cerrar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmIvaVenta.frx":0000
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   11
         Top             =   135
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Generar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmIvaVenta.frx":0400
         BorderGap       =   10
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   375
         Index           =   1
         Left            =   1470
         TabIndex        =   12
         Top             =   120
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Definitivo"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmIvaVenta.frx":083A
         BorderGap       =   10
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   375
         Index           =   2
         Left            =   2910
         TabIndex        =   13
         Top             =   120
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmIvaVenta.frx":0C4E
         BorderGap       =   10
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   315
         Left            =   8730
         TabIndex        =   21
         Top             =   150
         Width           =   915
         _Version        =   851968
         _ExtentX        =   1614
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Poner"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.TabControl TabIva 
      Height          =   4365
      Left            =   60
      TabIndex        =   0
      Top             =   1620
      Width           =   11955
      _Version        =   851968
      _ExtentX        =   21087
      _ExtentY        =   7699
      _StockProps     =   68
      PaintManager.Layout=   4
      PaintManager.BoldSelected=   -1  'True
      ItemCount       =   3
      Item(0).Caption =   "Listado"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "KlexFacturas"
      Item(1).Caption =   "Totales"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "klexTotales"
      Item(1).Control(1)=   "PushButton1"
      Item(2).Caption =   "Totales por Localidad"
      Item(2).ControlCount=   3
      Item(2).Control(0)=   "gridTLocalidad"
      Item(2).Control(1)=   "gridTotalesLocalidad"
      Item(2).Control(2)=   "PushButton5"
      Begin XtremeSuiteControls.PushButton PushButton5 
         Height          =   315
         Left            =   -69850
         TabIndex        =   27
         Top             =   2490
         Visible         =   0   'False
         Width           =   3195
         _Version        =   851968
         _ExtentX        =   5636
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Recalcular"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   345
         Left            =   -60820
         TabIndex        =   14
         Top             =   390
         Visible         =   0   'False
         Width           =   2685
         _Version        =   851968
         _ExtentX        =   4736
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Imprimir Totales"
         UseVisualStyle  =   -1  'True
      End
      Begin Grid.KlexGrid klexTotales 
         Height          =   3705
         Left            =   -69910
         TabIndex        =   1
         Top             =   810
         Visible         =   0   'False
         Width           =   11805
         _ExtentX        =   20823
         _ExtentY        =   6535
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
         MouseIcon       =   "frmIvaVenta.frx":105C
         Rows            =   10
      End
      Begin Grid.KlexGrid KlexFacturas 
         Height          =   3795
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   6694
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
         MouseIcon       =   "frmIvaVenta.frx":1078
         Rows            =   10
      End
      Begin Grid.KlexGrid gridTLocalidad 
         Height          =   1905
         Left            =   -69880
         TabIndex        =   25
         Top             =   480
         Visible         =   0   'False
         Width           =   11685
         _ExtentX        =   20611
         _ExtentY        =   3360
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
         MouseIcon       =   "frmIvaVenta.frx":1094
         Rows            =   10
      End
      Begin Grid.KlexGrid gridTotalesLocalidad 
         Height          =   1065
         Left            =   -69880
         TabIndex        =   26
         Top             =   2940
         Visible         =   0   'False
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   1879
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
         MouseIcon       =   "frmIvaVenta.frx":10B0
         Rows            =   10
      End
   End
   Begin MSAdodcLib.Adodc bFacturas 
      Height          =   330
      Left            =   240
      Top             =   6960
      Visible         =   0   'False
      Width           =   10005
      _ExtentX        =   17648
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
      Caption         =   "bFacturas"
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
   Begin MSAdodcLib.Adodc bTemp_Iva 
      Height          =   330
      Left            =   240
      Top             =   6960
      Visible         =   0   'False
      Width           =   10005
      _ExtentX        =   17648
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
      Caption         =   "bTemp_Iva"
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
   Begin MSAdodcLib.Adodc bIvaVenta 
      Height          =   330
      Left            =   240
      Top             =   6960
      Visible         =   0   'False
      Width           =   10005
      _ExtentX        =   17648
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
      Caption         =   "bIvaVenta"
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
   Begin XtremeSuiteControls.GroupBox GBParametros 
      Height          =   555
      Left            =   30
      TabIndex        =   3
      Top             =   1050
      Width           =   11865
      _Version        =   851968
      _ExtentX        =   20929
      _ExtentY        =   979
      _StockProps     =   79
      Caption         =   "Parametros"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.CheckBox chkcae 
         Height          =   225
         Left            =   6390
         TabIndex        =   24
         Top             =   240
         Width           =   2385
         _Version        =   851968
         _ExtentX        =   4207
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Solo Doc. con C.A.E."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkTotales 
         Height          =   255
         Left            =   9060
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   2625
         _Version        =   851968
         _ExtentX        =   4630
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Solo Calcular Totales del Periodo"
         UseVisualStyle  =   -1  'True
      End
      Begin Aplisoft_CajasDeTexto.TxF dtpFecha 
         Height          =   285
         Index           =   0
         Left            =   1260
         TabIndex        =   5
         Top             =   180
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   503
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
      Begin Aplisoft_CajasDeTexto.TxF dtpFecha 
         Height          =   285
         Index           =   1
         Left            =   3960
         TabIndex        =   6
         Top             =   180
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   503
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
      Begin VB.Label lblDatos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo Inicial :"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label lblDatos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo Final : "
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   7
         Top             =   240
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmIvaVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vcont As Integer

Dim vcuit, vIdTipoIva As String

Dim vcoddoc As Integer

Dim vtipo As String
Dim vtitulo As String
Dim fdesde  As String
Dim fhasta As String
Dim vNumeroPagina As Integer, vNumeroInicial As Integer

Dim vNetoFactA As Double, vI105FactA As Double, vI210FactA As Double, vI270FactA As Double, vTotalFactA As Double
Dim vNetoMonotributo As Double, vI105Monotributo As Double, vI210Monotributo As Double, vI270Monotributo As Double, vTotalMonotributo As Double
Dim vNetoNotaC As Double, vI105NotaC As Double, vI210NotaC As Double, vI270NotaC As Double, vTotalNotaC As Double
Dim vNetoNotaD As Double, vI105NotaD As Double, vI210NotaD As Double, vI270NotaD As Double, vTotalNotaD As Double
Dim vali As Double
Dim vgline As Integer
Dim vTipoIva As Double

Dim Canal1%, Canal2%

Dim vgline2 As Integer

Private Sub CopiarFacturas()
On Error Resume Next

Dim mesano As String

mesano = Format(Me.dtpFecha(0), "MM") + Format(Me.dtpFecha(0), "YYYY")


Canal1 = FreeFile
Open App.Path + "\CITI\REGINFO_CV_VENTAS_CBTE_" + mesano + ".TXT" For Output As Canal1

Canal2 = FreeFile
Open App.Path + "\CITI\REGINFO_CV_VENTAS_ALICUOTAS_" + mesano + ".TXT" For Output As Canal2

    vgline = 0
    
    With bFacturas
    
        Do Until .Recordset.EOF = True

            Call CopiarTemp(0)
            
            .Recordset.MoveNext
            b1.Value = b1.Value + 1
        Loop
    End With

Close Canal1

Close Canal2


If Err Then GrabarLog "CopiarFacturas", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CopiarTemp(vtipo As Byte)
On Error Resume Next

Dim vsql, vIdTipoIva, vtipodoc As String
Dim vlb, vla, vla2 As String
'Dim vIdTipoIva As String
Dim b1 As Integer
Dim ponerenciti As Boolean


    vsql = "select * from clientes where codigo='" + bFacturas.Recordset("Codigo").Value + "'"
    vIdTipoIva = traerDatos2(vsql, "idTipoIva", pathDBMySQL)
    
    vtipodoc = bFacturas.Recordset("Tipo").Value


    If vtipodoc = "Documento" Then
        'logform (" ...... Documentos no válidos como Factura : " + bFacturas.Recordset("nombre").Value)
        Exit Sub
    End If
    
    
    vlb = ""
    vla = ""
    
    With bTemp_Iva
        .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM Temp_Iva group by PuntoDeVenta,NroComprobante,TipoMovimiento,Codigo ORDER BY Fecha ASC, TipoMovimiento ASC, NroComprobante"
        .Refresh
        
        .Recordset.AddNew
        
        .Recordset("remito").Value = bFacturas.Recordset("remito")
        .Recordset("Fecha").Value = strfechaMySQL(bFacturas.Recordset("Fecha").Value)
        '.Recordset("TipoMovimiento").Value = EsNulo(bfacturas.Recordset("TipoMovimiento").Value) ' panic: ver si dejo esta configuración con el tipo
        .Recordset("TipoMovimiento").Value = EsNulo(bFacturas.Recordset("Tipo").Value)
         
        Dim vtipoDocu As String
        
        If .Recordset("TipoMovimiento") = "Nota C" And bFacturas.Recordset("Letra").Value = "B" Then
           vtipoDocu = getTipoDoc("NotaCB")
        Else
            vtipoDocu = getTipoDoc(.Recordset("TipoMovimiento").Value)
        End If
        

    
        .Recordset("PuntoDeVenta").Value = FormatoFactura(4, bFacturas.Recordset("PuntoDeVenta").Value)
        
        .Recordset("NroComprobante").Value = FormatoFactura(8, bFacturas.Recordset("NComprobante").Value)
        
        .Recordset("Cuit").Value = bFacturas.Recordset("CUIT").Value

        .Recordset("Nombre").Value = bFacturas.Recordset("Nombre").Value ' denominación

        .Recordset("Total").Value = bFacturas.Recordset("Total").Value
   
        If Abs(.Recordset("iva270").Value) Then vali = vali + 1
       
        .Recordset("Letra").Value = EsNulo(bFacturas.Recordset("Letra").Value)

        
        If Trim(bFacturas.Recordset("Codigo").Value) = "0006" Then
         '  MsgBox ">" + bfacturas.Recordset("Codigo").Value
        End If
        
       .Recordset("Codigo").Value = bFacturas.Recordset("Codigo").Value

        If EsNulo(bFacturas.Recordset("estadodocumento").Value) = "Anulado" Then
            .Recordset("Nombre").Value = bFacturas.Recordset("Nombre").Value + "(Anulado)"
        Else
            .Recordset("Neto").Value = Val(bFacturas.Recordset("SubTotal").Value)
            

        If UCase(LeerXml("Puesto")) = "PONS" Or UCase(LeerXml("PorRemito")) = "TRUE" Then   ' arreglarlo
            Call GuardarIvaFactura(bTemp_Iva.Recordset, bFacturas.Recordset("Remito").Value)  ' paso 2
        Else
            Call GuardarIvaFactura(bTemp_Iva.Recordset, bFacturas.Recordset("nrointerno").Value)  ' paso 2
        End If
        
        End If
        
        Select Case Trim(bFacturas.Recordset("TipoMovimiento").Value)
        
            Case "FC"
                .Recordset("Total").Value = bFacturas.Recordset("Total").Value
            Case "LV", "SL", "RG", "IB", "SU", "RI", "RV", "NC"
                .Recordset("Total").Value = IIf(bFacturas.Recordset("Total").Value > 0, bFacturas.Recordset("Total").Value * (-1), bFacturas.Recordset("Total").Value)
                .Recordset("Neto").Value = IIf(bFacturas.Recordset("SubTotal").Value > 0, bFacturas.Recordset("SubTotal").Value * (-1), bFacturas.Recordset("SubTotal").Value)
            Case Else
                'MsgBox Trim(bFacturas.Recordset("TipoMovimiento").Value)
        End Select
               
                .Recordset("Total").Value = Abs(bFacturas.Recordset("Total").Value)
                .Recordset("Neto").Value = Abs(bFacturas.Recordset("SubTotal").Value)
                .Recordset("iva105").Value = Abs(.Recordset("iva105").Value)
                .Recordset("iva210").Value = Abs(.Recordset("iva210").Value)
                .Recordset("iva270").Value = Abs(.Recordset("iva270").Value)
               
               
        If bFacturas.Recordset("Tipo").Value = "Nota C" Then
                .Recordset("Total").Value = IIf(bFacturas.Recordset("Total").Value > 0, bFacturas.Recordset("Total").Value * (-1), bFacturas.Recordset("Total").Value)
                .Recordset("Neto").Value = IIf(bFacturas.Recordset("SubTotal").Value > 0, bFacturas.Recordset("SubTotal").Value * (-1), bFacturas.Recordset("SubTotal").Value)
                .Recordset("iva105").Value = -1 * .Recordset("iva105").Value
                .Recordset("iva210").Value = -1 * .Recordset("iva210").Value
                .Recordset("iva270").Value = -1 * .Recordset("iva270").Value
        End If
                
                
        If bFacturas.Recordset("Tipo").Value = "Nota C" And Not .Recordset("Letra").Value = "A" Then  ' panic 2310
          
          
               .Recordset("Total").Value = IIf(bFacturas.Recordset("Total").Value > 0, bFacturas.Recordset("Total").Value * (-1), bFacturas.Recordset("Total").Value)
    
                .Recordset("Neto").Value = .Recordset("Total").Value / 1.21
               .Recordset("iva210").Value = .Recordset("total").Value - .Recordset("neto").Value
        
        
        End If
                
                
                
        
               
        If bFacturas.Recordset("Tipo").Value = "Fact B" Then
                .Recordset("Neto").Value = bFacturas.Recordset("Total").Value / (1.21)
                .Recordset("Iva210").Value = bFacturas.Recordset("Total").Value - .Recordset("Neto").Value
        End If
                
                
        '.Recordset("NroComprobante").Value = FormatoNC(bFacturas.Recordset("NComprobante").Value)
        '.Recordset("Nombre").Value = "ANULADA"
        
        .Recordset("Remito").Value = bFacturas.Recordset("Remito").Value
        
        vNumeroPagina = vNumeroInicial + SeleccionarNumero15(.Recordset.RecordCount)
        .Recordset("NumHoja").Value = vNumeroPagina
        
        
    
        ponerenciti = True
        Debug.Print "-----------> " + Str(bFacturas.Recordset("remito").Value)
       
       If bFacturas.Recordset("repartidor").Value = "Anulado" Then
                Debug.Print "=========== > " + Str(bFacturas.Recordset("remito").Value)
                
                ponerenciti = False
                KlexFacturas.Col = 0
                KlexFacturas.Row = .Recordset.AbsolutePosition
                KlexFacturas.CellBackColor = vbRed
                
                '.Recordset("remito") = 0
                
                .Recordset("Codigo").Value = "Anulado"
                
        Else
                Debug.Print "++++++++++++ > " + Str(bFacturas.Recordset("remito").Value)
       End If
        
        .Recordset("remito") = bFacturas.Recordset("remito").Value

        .Recordset.Update
        
        KlexFacturas.TextMatrix(bFacturas.Recordset.AbsolutePosition, 0) = .Recordset("remito")
        KlexFacturas.TextMatrix(bFacturas.Recordset.AbsolutePosition, 1) = strfechaMySQL(.Recordset("Fecha").Value)
        KlexFacturas.TextMatrix(bFacturas.Recordset.AbsolutePosition, 2) = EsNulo(.Recordset("TipoMovimiento").Value)
        KlexFacturas.TextMatrix(bFacturas.Recordset.AbsolutePosition, 3) = EsNulo(.Recordset("Letra").Value)
        KlexFacturas.TextMatrix(bFacturas.Recordset.AbsolutePosition, 4) = EsNulo(.Recordset("PuntoDeVenta").Value)
        KlexFacturas.TextMatrix(bFacturas.Recordset.AbsolutePosition, 5) = EsNulo(.Recordset("NroComprobante").Value)
        KlexFacturas.TextMatrix(bFacturas.Recordset.AbsolutePosition, 6) = EsNulo(.Recordset("Codigo").Value)
        KlexFacturas.TextMatrix(bFacturas.Recordset.AbsolutePosition, 7) = EsNulo(.Recordset("Nombre").Value)
        KlexFacturas.TextMatrix(bFacturas.Recordset.AbsolutePosition, 8) = EsNulo(.Recordset("Cuit").Value)
        KlexFacturas.TextMatrix(bFacturas.Recordset.AbsolutePosition, 9) = formatNumero(EsNulo(.Recordset("Neto").Value))
        
        KlexFacturas.TextMatrix(bFacturas.Recordset.AbsolutePosition, 10) = formatNumero(EsNulo(.Recordset("Iva105").Value))
        KlexFacturas.TextMatrix(bFacturas.Recordset.AbsolutePosition, 11) = formatNumero(EsNulo(.Recordset("Iva210").Value))
        KlexFacturas.TextMatrix(bFacturas.Recordset.AbsolutePosition, 12) = formatNumero(EsNulo(.Recordset("Iva270").Value))
        
        KlexFacturas.TextMatrix(bFacturas.Recordset.AbsolutePosition, 13) = formatNumero(EsNulo(.Recordset("Retenciones").Value))
        KlexFacturas.TextMatrix(bFacturas.Recordset.AbsolutePosition, 14) = formatNumero(EsNulo(.Recordset("Total").Value))
       
       
       KlexFacturas.TopRow = .Recordset.AbsolutePosition

           ' ------------------------citi cbte ----------------------------------
       
       ' Dim vtipoDocu As String
       ' Dim vali As Double
        
       ' vtipoDocu = getTipoDoc(.Recordset("TipoMovimiento").Value)
       
        vgline = vgline + 1
        vlb = ""
        
        If vgline > 1 Then vlb = vbCrLf
        
        vlb = vlb + FF(.Recordset("Fecha").Value) ' 1

        vlb = vlb + fn(Val(vtipoDocu), 3, "N") ' 2

        vlb = vlb + fn(.Recordset("PuntoDeVenta").Value, 5, "N") ' 3'

        vlb = vlb + fn(.Recordset("NroComprobante").Value, 20, "N") ' nro de comprobante 4

        vlb = vlb + fn(.Recordset("NroComprobante").Value, 20, "N") ' nro de comprobante hasta 5
        
        Debug.Print (Val(.Recordset("NroComprobante").Value))
        
        
        If 11605 = Val(.Recordset("NroComprobante").Value) Then
           ' MsgBox " parada"
        End If
    
        
        vIdTipoIva = traerDatos2("select idTipoIva as c from clientes where codigo = '" + .Recordset("codigo") + "'", "c", pathDBMySQL)
        
        If vIdTipoIva = "005" Then
            vcoddoc = 99
            vcuit = "00000000000"
        Else
            vcoddoc = 80
            vcuit = .Recordset("Cuit").Value
        
        
        End If
        
        
        
        vlb = vlb + fn(vcoddoc, 2, "N") ' Código de documento del comprador 6
        
        'vlb = vlb + fn(Trim(.Recordset("Cuit").Value), 20, "N") ' Código de documento del comp ' 7'
        vlb = vlb + fn(Trim(vcuit), 20, "N") ' Código de documento del comp ' 7'
        
       
        If Not validarCUIT2(.Recordset("Cuit").Value) = 1 And vcoddoc = 80 Then
                ponerenciti = False
                vcont = vcont + 1
                logform ("Error #" + Str(vcont) + " Cuit mal formado: " + .Recordset("Cuit").Value + " Nro. Comp. " + .Recordset("NroComprobante").Value + "  Persona: " + bFacturas.Recordset("Nombre").Value)
        End If
        
        
        
        
        vlb = vlb + fc(.Recordset("Nombre").Value, 30) ' denominación   '8'
        
         vlb = vlb + fn(.Recordset("Total").Value, 15, "S") ' total '9'
       
         vlb = vlb + fn(.Recordset("NoGravado").Value, 15, "S") ' NoGravado '10'

         vlb = vlb + fn(0, 15, "S") ' Percepción a no categorizados '11'

         vlb = vlb + fn(.Recordset("ImpExento").Value, 15, "S") ' ImpExento (12)

         vlb = vlb + fn(.Recordset("Percepciones").Value, 15, "S") ' Percepciones (13)

         vlb = vlb + fn(0, 15, "S") ' Importe de percepciones de Ingresos Brutos (14)
        
         vlb = vlb + fn(0, 15, "S") ' Importe de percepciones impuestos Municipales (15)

         vlb = vlb + fn(0, 15, "S") ' Importe impuesto interno (16)
        
         vlb = vlb + fc("PES", 3) ' Código de la moneda (17)

         vlb = vlb + fn(10000, 10, "S") ' Importe impuesto interno (18)


        vali = 0
        
        If Abs(.Recordset("iva105").Value) > 0 Then vali = vali + 1
        If Abs(.Recordset("iva210").Value) > 0 Then vali = vali + 1
        If Abs(.Recordset("iva270").Value) > 0 Then vali = vali + 1
        
        If vali = 0 Then vali = 1


        vlb = vlb + fn(vali, 1, "N") ' Importe impuesto interno (19)
        
        If vali = 0 Then
             vlb = vlb + fc("N", 1) ' Importe impuesto interno (20)
        Else
             vlb = vlb + fc(" ", 1)
        End If
        
        vlb = vlb + fn(0, 15, "S") ' otros tributos (21)
        
        vlb = vlb + FF(Date + 30) ' fecha vencimiento (22)

        If ponerenciti Then Print #Canal1, vlb;  ' graba la linea en el archivo REGINFO_CV_VENTAS_CBTE.TXT'
    
        '-----------
            
        vla = ""
        vgline2 = vgline2 + 1
        
        If vgline2 > 1 Then vla = vbCrLf
        
        ' -------- acá arranca alicuota
        
        
        vla = vla + fn(Val(vtipoDocu), 3, "N") ' 1
        vla = vla + fn(.Recordset("PuntoDeVenta").Value, 5, "N") ' 2
        vla = vla + fn(.Recordset("NroComprobante").Value, 20, "N") ' nro de comprobante '3'
        
        
        Dim neto_para21, neto_para105, neto_para27, vtotal, vneto   As Double
        
        vtotal = .Recordset("total")
        vneto = .Recordset("neto")
        
        neto_para21 = fcalculaNeto(21, .Recordset("iva210").Value, vtotal, vneto)
        
        neto_para105 = fcalculaNeto(10.5, .Recordset("iva105").Value, vtotal, vneto)
        
        neto_para27 = fcalculaNeto(27, .Recordset("iva270").Value, vtotal, vneto)
        
        
        
        neto_para105 = Round((vneto - Round(neto_para21, 2)), 3)
        
       ' vla = vla + fn(.Recordset("Neto").Value, 15, "S")  ' Importe neto gravado '4'

       ' vla2 = vla
       
       ' parada todo ale

        b1 = 0
        If Abs(.Recordset("iva105").Value) > 0.01 Then ' iva 10.5'
        
                vla2 = vla + fn(neto_para105, 15, "S")
                vla2 = vla2 + fn(4, 4, "N")
                vla2 = vla2 + fn(Abs(.Recordset("iva105").Value), 15, "S")
                
               If ponerenciti Then Print #Canal2, vla2;  '\REGINFO_CV_VENTAS_ALICUOTAS.TXT'
                
                b1 = b1 + 1
        End If

        vla2 = vla


        If Abs(.Recordset("iva210").Value) > 0 Then  'iva 21'
            vla2 = vla + fn(neto_para21, 15, "S")
            vla2 = vla2 + fn(5, 4, "N")  ' Importe neto gravado '4'
            vla2 = vla2 + fn(Abs(.Recordset("iva210").Value), 15, "S")
            
           If ponerenciti Then Print #Canal2, vla2; '\REGINFO_CV_VENTAS_ALICUOTAS.TXT'
            b1 = b1 + 1
        End If

        vla2 = vla

        If Abs(.Recordset("iva270").Value) > 0 Then  ' iva 27'
        
                vla2 = vla + fn(neto_para27, 15, "S")
                vla2 = vla2 + fn(6, 4, "N")  ' Importe neto gravado '4'
                vla2 = vla2 + fn(Abs(.Recordset("iva270").Value), 15, "S")
                
                
                If ponerenciti Then Print #Canal2, vla2;     '\REGINFO_CV_VENTAS_ALICUOTAS.TXT'
                b1 = b1 + 1
        
        End If
        
        If b1 = 0 Then ' no hubo ivas
            vgline2 = vgline2 - 1
        End If

       '-----------------------------------------------------------------------

       
    End With
    
If Err Then GrabarLog "CopiarTemp", Err.Number & " " & Err.Description, Me.Name
End Sub

Function fcalculaNeto(piva As Double, vIva As Double, ByVal vtotal As Double, ByVal vneto As Double) As Double
    Dim p, s, d As Double
    
    d = vtotal - vneto
    
    p = vIva * 100 / d
    
    s = vneto * p / 100


    'fcalculaNeto = (vIva * 100 / piva)
    
    fcalculaNeto = s


End Function


Private Sub GuardarIvaFactura(rsTemp_Iva As Recordset, vnroremito As Long)
On Error Resume Next

    Dim rsIvaFactVenta As New ADODB.Recordset, sqlIvaFactVenta As String
    Dim vPercepciones As Double, vRetenciones As Double, vITC As Double, vNoGravado As Double, vImpExento As Double
    
    ' corregir el problema de no tener que controlar que el nro interno sea = 0
    'sqlIvaFactVenta = "SELECT * FROM IvaFacturaVenta WHERE (Remito = " & Val(vnroremito) & ")"
    If vnroremito = 5940 Then
    
       ' MsgBox "parar"
    
    End If
    
    If UCase(LeerXml("Puesto")) = "PONS" Or UCase(LeerXml("PorRemito")) = "TRUE" Then
        sqlIvaFactVenta = "SELECT * FROM IvaFacturaVenta WHERE (remito = " & Str(vnroremito) & ") order by idIvaFacturaVenta desc"
    Else
          sqlIvaFactVenta = "SELECT * FROM IvaFacturaVenta WHERE (nrointerno = " & Str(vnroremito) & ") order by idIvaFacturaVenta desc"
    End If
    
    
    With rsIvaFactVenta
        Call .Open(sqlIvaFactVenta, ConnDDBB, adOpenDynamic, adLockBatchOptimistic)
        
        If Not .EOF = True Then
        
            vPercepciones = 0
            vRetenciones = 0
            vITC = 0
            vNoGravado = 0
            vImpExento = 0
            
' ver si pongo abs

            rsTemp_Iva.Fields("IVA105").Value = Val(Format(.Fields("Iva105").Value, "######0.00"))
            rsTemp_Iva.Fields("IVA210").Value = Val(Format(.Fields("Iva210").Value, "######0.00"))
            rsTemp_Iva.Fields("IVA270").Value = Val(Format(.Fields("Iva270").Value, "######0.00"))

            vPercepciones = Val(Format(.Fields("Percepciones").Value, "######0.00"))
            vRetenciones = Val(Format(.Fields("Retenciones").Value, "######0.00"))
            vITC = Val(Format(.Fields("ITC").Value, "######0.00"))
            vNoGravado = Val(Format(.Fields("NoGravado").Value, "######0.00"))
            vImpExento = Val(Format(.Fields("ImpExento").Value, "######0.00"))

            rsTemp_Iva.Fields("Percepciones").Value = vPercepciones 'IIf(vPercepciones > 0, vPercepciones * -1, vPercepciones)
            rsTemp_Iva.Fields("Retenciones").Value = vRetenciones 'IIf(vRetenciones > 0, vRetenciones * -1, vRetenciones)
            rsTemp_Iva.Fields("NoGravado").Value = IIf(vNoGravado > 0, vNoGravado * -1, vNoGravado)
            rsTemp_Iva.Fields("ITC").Value = IIf(vITC > 0, vITC * -1, vITC)
            rsTemp_Iva.Fields("ImpExento").Value = IIf(vImpExento > 0, vImpExento * -1, vImpExento)
            
                
             If traerDatos2("select * from factura where remito=" + Trim(.Fields("remito").Value), "TipoMovimiento", pathDBMySQL) = "NC" Then
                            rsTemp_Iva.Fields("Percepciones").Value = IIf(vPercepciones > 0, vPercepciones * -1, vPercepciones)
                            rsTemp_Iva.Fields("Retenciones").Value = IIf(vRetenciones > 0, vRetenciones * -1, vRetenciones)
                            rsTemp_Iva.Fields("NoGravado").Value = IIf(vNoGravado > 0, vNoGravado * -1, vNoGravado)
                            rsTemp_Iva.Fields("ITC").Value = IIf(vITC > 0, vITC * -1, vITC)
                            rsTemp_Iva.Fields("ImpExento").Value = IIf(vImpExento > 0, vImpExento * -1, vImpExento)
                            
                            rsTemp_Iva.Fields("IVA105").Value = -1 * Val(Format(.Fields("Iva105").Value, "######0.00"))
                            rsTemp_Iva.Fields("IVA210").Value = -1 * Val(Format(.Fields("Iva210").Value, "######0.00"))
                            rsTemp_Iva.Fields("IVA270").Value = -1 * Val(Format(.Fields("Iva270").Value, "######0.00"))
             End If
            
            End If
        
    End With
    
    sqlIvaFactVenta = ""
    
    If rsIvaFactVenta.State = 1 Then
        rsIvaFactVenta.Close
        Set rsIvaFactVenta = Nothing
    End If
        
If Err Then GrabarLog "GuardarIvaFactura", Err.Number & " " & Err.Description, Me.Name
End Sub
Function FormatoFactura(vcant As Integer, vNcomp As Long) As String
On Error Resume Next
    
    Dim i As Integer

    FormatoFactura = String(vcant - Len(Trim(Str(vNcomp))), "0") & vNcomp

If Err Then GrabarLog "FormatoFactura", Err.Number & " " & Err.Description, Me.Name
End Function
Function FormatoNC(vNcomp As Long) As String
On Error Resume Next
    
    Dim i, vsucursal As Integer

    FormatoNC = "0001-" & String(8 - Len(Trim(Str(vNcomp))), "0") & vNcomp

If Err Then GrabarLog "FormatoNC", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub DocumentosAnulados()
On Error Resume Next
    
    Dim vDocAnulado As Long

    With bIvaVenta
        .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT Tipo,NroComprobante  FROM IvaVenta WHERE (tipo = '" & vtipo & "') ORDER BY nrocomprobante ASC"
        .Refresh
        
        If Not .Recordset.EOF = True Then
            .Recordset.MoveLast
            vDocAnulado = .Recordset("nrocomprobante").Value + 1
        Else
            vDocAnulado = 1
        End If
        
    End With

If Err Then GrabarLog "DocumentosAnulados", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub GenerarListado()
On Error Resume Next

b1.Value = 0
    If Trim(dtpFecha(0).Text) = "" Or Trim(dtpFecha(1).Text) = "" Then
        MsgBox "Debe Seleccionar un Mes y/o un Año para poder ejecutar un listado!!!", vbInformation, "Mensaje ..."
        Exit Sub
    End If
    
Dim vcondEmpresas As String

vcondEmpresas = ""


If Not Val(Me.vcodEmpresa.Text) = 0 Then
    vcondEmpresas = " and (nrointerno in select nrointerno from t_rel where where idEmpresa =  " + Me.vcodEmpresa.Tag + ")"
End If


    If Not chkTotales.Value = xtpChecked Then
        With bFacturas
            .ConnectionString = pathDBMySQL
            '.RecordSource = "SELECT * FROM factura WHERE (month(fecha) = '" & AjustarMes(Month(dtpFecha(0).Value)) & "' AND year(fecha) = '" & Year(dtpFecha(1).Value) & "') AND (tipo <> 'Documento' OR Tipo Is NULL) AND (TipoMovimiento <> 'RC')  ORDER BY Fecha ASC, tipo ASC, Ncomprobante ASC"
            .RecordSource = "SELECT * FROM factura WHERE (month(fecha) = '" & AjustarMes(Month(dtpFecha(0).Value)) & "' AND year(fecha) = '" & Year(dtpFecha(1).Value) & "') AND (tipo <> 'Documento' OR Tipo Is NULL) and (Letra = 'A' or Letra = 'B')   ORDER BY Fecha ASC, tipo ASC, Ncomprobante ASC"
            
            .RecordSource = "SELECT * FROM factura WHERE fecha >= '" + strfechaMySQL(dtpFecha(0)) + "' and fecha <= '" + strfechaMySQL(dtpFecha(1)) + "' AND (tipo <> 'Documento' OR Tipo Is NULL) and (Letra = 'A' or Letra = 'B')   " + _
            " " + vcondEmpresas + _
            "  group by PuntoDeVenta,Ncomprobante,TipoMovimiento,Codigo,tipo " + _
            " ORDER BY Fecha ASC, tipo ASC, Ncomprobante ASC"
            
            If chkcae.Value Then
                .RecordSource = "SELECT * FROM factura WHERE (not cae = '' and not cae is null) and fecha >= '" + strfechaMySQL(dtpFecha(0)) + "' and fecha <= '" + strfechaMySQL(dtpFecha(1)) + "' AND (tipo <> 'Documento' OR Tipo Is NULL) and (Letra = 'A' or Letra = 'B')  " + _
                "  group by PuntoDeVenta,Ncomprobante,TipoMovimiento,Codigo,tipo " + _
                " ORDER BY Fecha ASC, tipo ASC, Ncomprobante ASC"
            End If
            
            
          '  If UCase(LeerXml("Cliente")) = "PONS" Then
          '      .RecordSource = "SELECT * FROM factura WHERE (month(fecha) = '" & AjustarMes(Month(dtpFecha(0).Value)) & "' AND year(fecha) = '" & Year(dtpFecha(1).Value) & "') AND (tipo <> 'Documento' OR Tipo Is NULL) and " + _
          '      " (Letra = 'A' or Letra = 'B')   " + _
          '      " and (not cae='Anulado') " + _
          '      " ORDER BY Fecha ASC, tipo ASC, Ncomprobante ASC"
          '  End If
            
            .Refresh
            
            If Not .Recordset.EOF = True Then
                Barra.Value = 0
                Barra.Max = .Recordset.RecordCount
                FormatoGrilla (.Recordset.RecordCount)
            Else
                MsgBox "No existen movimientos de este mes!!!", vbExclamation, "Mensaje ..."
                FormatoGrilla (1)
                Exit Sub
            End If
            
            
        End With
    
        Call IniciarVariables
        
        Call CopiarFacturas  ' paso1
        
       ' Call DocumentosAnulados ' Panic ! ver si lo tengo que volver a poner
        
        Call CargarGrillaTotales
    
        TabIva.Item(0).Selected = True
    
    Else
        
        Call CargarGrillaTotales
        TabIva.Item(1).Selected = True
    
    End If
If Err Then GrabarLog "GenerarListado", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub FormatoGrilla(vCantidadRenglones As Long)
On Error Resume Next

    Dim i As Integer
   
    With KlexFacturas
        .FixedRows = 1
        .FixedCols = 1

        .Cols = 20
        .Rows = vCantidadRenglones + 1
    
        If vCantidadRenglones = 1 Then
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
                .ColWidth(i) = 0
            Next
        End If
    
        
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 150
    
        .TextMatrix(0, 1) = "Fecha"
        .ColWidth(1) = 1000
    
        .TextMatrix(0, 2) = "Tipo"
        .ColWidth(2) = 500
        
        .TextMatrix(0, 3) = "L."
        .ColWidth(3) = 350
           
        .TextMatrix(0, 4) = "Punto V"
        .ColWidth(4) = 500
        
        .TextMatrix(0, 5) = "Nro Comp."
        .ColWidth(5) = 500
        
        .TextMatrix(0, 6) = "Codigo"
        .ColWidth(6) = 0
            
        .TextMatrix(0, 7) = "Razon Social"
        .ColWidth(7) = 3500
            
        .TextMatrix(0, 8) = "Cuit"
        .ColWidth(8) = 0
            
        .TextMatrix(0, 9) = "Bruto"
        .ColWidth(9) = 1000
            
        .TextMatrix(0, 10) = "% 10.5"
        .ColWidth(10) = 800
            
        .TextMatrix(0, 11) = "% 21"
        .ColWidth(11) = 800
            
        .TextMatrix(0, 12) = "% 27"
        .ColWidth(12) = 0
        
        .TextMatrix(0, 13) = "$ Retenciones"
        .ColWidth(13) = 800
        
        .TextMatrix(0, 14) = "$ Neto"
        .ColWidth(14) = 800
        
        .BackColorAlternate = &HC0C0C0
    End With

If Err Then GrabarLog "FormatoGriila", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub GuardarIva()
Dim i As Integer
On Error Resume Next

    With bIvaVenta
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM IvaVenta"
        .Refresh
    End With
        
    With bTemp_Iva
        .Refresh
        If Not .Recordset.EOF = True Then
            .Recordset.MoveFirst
            Barra.Value = 0
            Barra.Max = .Recordset.RecordCount
        Else
            MsgBox "NO tiene datos pre-cargados para guardar!!", vbExclamation, "Mensaje ..."
            Exit Sub
        End If
        
        Do Until .Recordset.EOF
            bIvaVenta.Recordset.AddNew
            For i = 1 To (.Recordset.Fields.Count - 1)
                If Not IsNull(.Recordset(i).Value) = True Then
                    bIvaVenta.Recordset(i).Value = .Recordset(i).Value
                End If
            Next
            Barra.Value = Barra.Value + 1
            bIvaVenta.Recordset.Update
            .Recordset.MoveNext
        Loop
        
    End With
    
    Call EjecutarScript("INSERT INTO IvaVentaCerrado (Periodo) VALUES ('" & AjustarMes(Month(dtpFecha(0).Value)) & Year(dtpFecha(1).Value) & "')")

If Err Then GrabarLog "GuardarIva", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub Form_Load()
On Error Resume Next

    With Me
        .KeyPreview = True
        .Top = 0
        .Left = 0
        .width = 12000
        .height = 6550
        .Show
    End With

    FormatoGrilla (1)
    FormatoGrillaTotales (1)
    
    dtpFecha(0).SetFocus
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    If KeyCode = vbKeyF6 Then
        If chkTotales.Value = xtpChecked Then
            chkTotales.Value = xtpUnchecked
        Else
            chkTotales.Value = xtpChecked
        End If
    End If

If Err Then GrabarLog "Form_KeyUp", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub IniciarVariables()
  On Error Resume Next

    'fdesde = "01" & "/" & AjustarMes(txtFecha(0).Value) & "/" & Year(txtFecha(1).Value)
    'fhasta = DiasDelMes(fdesde) & "/" & AjustarMes(txtFecha(0).Value) & "/" & Year(txtFecha(1).Value)

    fdesde = strfechaMySQL(dtpFecha(0).Value)
    fhasta = strfechaMySQL(dtpFecha(1).Value)

    Call BorrarBase("Temp_Iva", pathDBMySQL)
    Call BorrarBase("Temp", pathDBMySQL)
    
    vNumeroPagina = 0
    vNumeroInicial = UltimaHoja(False, "NroHojaIV")
    
    vNetoFactA = 0
    vI105FactA = 0
    vI210FactA = 0
    vI270FactA = 0
    vTotalFactA = 0
    
    vNetoMonotributo = 0
    vI105Monotributo = 0
    vI210Monotributo = 0
    vI270Monotributo = 0
    vTotalMonotributo = 0
    
    vNetoNotaC = 0
    vI105NotaC = 0
    vI210NotaC = 0
    vI270NotaC = 0
    vTotalNotaC = 0
    
    vNetoNotaD = 0
    vI105NotaD = 0
    vI210NotaD = 0
    vI270NotaD = 0
    vTotalNotaD = 0
    
If Err Then GrabarLog "IniciarVariables", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Reporte()
On Error Resume Next

    Unload Mantenimiento
    Load Mantenimiento

    MsgBox "Prepare la Impresora!!!!", vbInformation, "Mensaje ..."
    
    If Not LeerConfig(21) = "PorFecha" Then
    
        With Mantenimiento.rsIva
            If Not .State = 0 Then .Close
        
            .Source = "SHAPE {SELECT * FROM Temp} AS Iva APPEND ({SELECT *, (t.`Iva105` + t.`Iva210` + t.`Iva270`)as ivas FROM Temp_Iva t where  not Codigo = 'Anulado'} AS Temp_Iva RELATE 'NumHoja' TO 'NumHoja') AS Temp_Iva"
        
            If Not .State = 1 Then .Open
            .Close
            .Open
        
            .Filter = "1=1"
            .Sort = "id_temp ASC"
        
            If .RecordCount = 0 Then
                MsgBox "No existen datos para visualizar", vbExclamation, "Mensaje ..."
                Exit Sub
            End If
        End With

        'vI105FactA = CalcularTotal("Fact A", "I105")
        'vI210FactA = CalcularTotal("Fact A", "I210")
        'vI270FactA = CalcularTotal("Fact A", "I270")
        'vNetoFactA = CalcularTotal("Fact A", "Neto")
        'vTotalFactA = CalcularTotal("Fact A", "Total")
        
        'vI105Monotributo = CalcularTotal("Fact B", "I105")
        'vI210Monotributo = CalcularTotal("Fact B", "I210")
        'vI270Monotributo = CalcularTotal("Fact B", "I270")
        'vNetoMonotributo = CalcularTotal("Fact B", "Neto") - vI105Monotributo - vI210Monotributo - vI270Monotributo
        'vTotalMonotributo = CalcularTotal("Fact B", "Total") - vI105Monotributo - vI210Monotributo - vI270Monotributo
    
        'vI105NotaC = Val(CalcularTotal("Nota C", "I105"))
        'vI210NotaC = Val(CalcularTotal("Nota C", "I210"))
        'vI270NotaC = Val(CalcularTotal("Nota C", "I270"))
        'vNetoNotaC = Val(CalcularTotal("Nota C", "Neto"))
        'vTotalNotaC = Val(CalcularTotal("Nota C", "Total"))
    
        'vNetoNotaD = -Val(CalcularTotal("Nota D", "Neto"))
        'vI105NotaD = -Val(CalcularTotal("Nota D", "I105"))
        'vI210NotaD = -Val(CalcularTotal("Nota D", "I210"))
        'vI270NotaD = -Val(CalcularTotal("Nota D", "I270"))
        'vTotalNotaD = -Val(CalcularTotal("Nota D", "Total"))
        
        With drIvaVenta.Sections("TituloEmpresa")
            .Controls("vmes").Caption = AjustarMes(Month(dtpFecha(0).Value))
            .Controls("vano").Caption = Year(dtpFecha(0).Value)
            
            .Controls("lblNombre").Caption = vDatosEmpresa.Nombre
            .Controls("lblDueno").Caption = "DE " & vDatosEmpresa.Responsable
            .Controls("lblCuit").Caption = "CUIT " & vDatosEmpresa.cuit
        End With
        
        With drIvaVenta.Sections("ReportFooter")
            .Controls("nfacturaa").Caption = Format(vNetoFactA, "$########0.00")
            .Controls("nfacturab").Caption = Format(vNetoMonotributo, "$########0.00")
            .Controls("nncredito").Caption = Format(vNetoNotaC, "$########0.00")
            .Controls("nndebito").Caption = Format(vI105NotaD, "$########0.00")
            '.Controls("nfacturae").Caption = Format(CalcularTotal("Fact E", "Neto"), "$#######0.00")
    
            .Controls("ifacturaa105").Caption = Format(vI105FactA, "$#######0.00")
            .Controls("ifacturab105").Caption = Format(vI105Monotributo, "$########0.00")
            .Controls("incredito105").Caption = Format(vI105NotaC, "$########0.00")
            .Controls("indebito105").Caption = Format(vI105NotaD, "$########0.00")
            '.Controls("ifacturae105").Caption = Format(CalcularTotal("Fact E", "Neto"), "$#######0.00")
    
            .Controls("ifacturaa210").Caption = Format(vI210FactA, "$#######0.00")
            .Controls("ifacturab210").Caption = Format(vI210Monotributo, "$########0.00")
            .Controls("incredito210").Caption = Format(vI210NotaC, "$########0.00")
            .Controls("indebito210").Caption = Format(vI210NotaD, "$########0.00")
            '.Controls("ifacturae210").Caption = Format(CalcularTotal("Fact E", "Neto"), "$#######0.00")
            
            .Controls("ifacturaa270").Caption = Format(vI270FactA, "$#######0.00")
            .Controls("ifacturab270").Caption = Format(vI270Monotributo, "$########0.00")
            .Controls("incredito270").Caption = Format(vI270NotaC, "$########0.00")
            .Controls("indebito270").Caption = Format(vI270NotaD, "$########0.00")
            '.Controls("ifacturae270").Caption = Format(CalcularTotal("Fact E", "Iva270"), "$#######0.00")
    
            .Controls("tfacturaa").Caption = Format(vTotalFactA, "$#######0.00")
            .Controls("tfacturab").Caption = Format(vTotalMonotributo, "$########0.00")
            .Controls("tncredito").Caption = Format(vTotalNotaC, "$########0.00")
            .Controls("tndebito").Caption = Format(vTotalNotaD, "$########0.00")
            '.Controls("tfacturae").Caption = Format(CalcularTotal("Fact E", "Total"), "$#######0.00")
        
            .Controls("ntotal").Caption = Format(vNetoFactA + vNetoMonotributo + vI105NotaC + vNetoNotaD, "$#######0.00")
            .Controls("I105total").Caption = Format(vI105FactA + vI105Monotributo + vI105NotaC + vI105NotaD, "$########0.00")
            .Controls("I210total").Caption = Format(vI210FactA + vI210Monotributo + vI210NotaC + vI210NotaD, "$########0.00")
            .Controls("I270total").Caption = Format(vI270FactA + vI270Monotributo + vI270NotaC + vI270NotaD, "$########0.00")
            .Controls("ttotal").Caption = Format(vTotalFactA + vTotalMonotributo + vTotalNotaC + vTotalNotaD, "$#######0.00")
        End With
    
        With drIvaVenta
            .Orientation = rptOrientLandscape
            
            .Refresh
            .Show
        End With

    Else
        With Mantenimiento.rsIvaPorFecha
            If Not .State = 0 Then .Close
        
            .Source = " SHAPE {SELECT Fecha FROM Temp_Iva GROUP BY Fecha} AS IvaPorFecha APPEND ({SELECT *, (t.`Iva105` + t.`Iva210` + t.`Iva270`) as ivas FROM Temp_Iva t where not codigo = 'Anulado'} AS IvaPorFechaDetalle RELATE 'Fecha' TO 'Fecha') AS IvaPorFechaDetalle"
        
            If Not .State = 1 Then .Open
            .Close
            .Open
        
            '.Sort = "id_temp ASC"
        
            If .RecordCount = 0 Then
                MsgBox "No existen datos para visualizar", vbExclamation, "Mensaje ..."
                Exit Sub
            End If
        End With
    
            
    
        With drIvaVentaPorFecha.Sections("TituloEmpresa")
        
        
               .Controls("vmes").Caption = AjustarMes(Month(dtpFecha(0).Value))
               .Controls("vano").Caption = Year(dtpFecha(0).Value)
               .Controls("lblResponsable").Caption = vDatosEmpresa.Direccion
                
        
            If vConfigGral.vempresa = "wgestionPoli" Then
            
                .Controls("lblNombre").Caption = "Poliwheel SRL"
                .Controls("lblDueno").Caption = ""
                .Controls("lblCuit").Caption = "30-70738431-6"
                            
            
            Else
                
                .Controls("lblNombre").Caption = vDatosEmpresa.Nombre
                .Controls("lblDueno").Caption = "DE " & vDatosEmpresa.Responsable
                .Controls("lblCuit").Caption = "CUIT " & vDatosEmpresa.cuit
            
            End If
        
        
        End With
        
        With drIvaVentaPorFecha
            '.Orientation = rptOrientLandscape
            
            .Refresh
            .Show
        End With
    End If
    

If Err Then GrabarLog "Reporte", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Function SeleccionarNumero15(vCantidad As Integer) As Integer
On Error Resume Next
    
    Select Case vCantidad
    
        Case 1 To 15
            SeleccionarNumero15 = 1
        
        Case 16 To 30
            SeleccionarNumero15 = 2
            
        Case 31 To 45
            SeleccionarNumero15 = 3
        
        Case 46 To 60
            SeleccionarNumero15 = 4
    
        Case 97 To 120
            SeleccionarNumero15 = 5
        
        Case 121 To 144
            SeleccionarNumero15 = 6
        
        Case 145 To 168
            SeleccionarNumero15 = 7

        Case 169 To 192
            SeleccionarNumero15 = 8

        Case 193 To 216
            SeleccionarNumero15 = 9

        Case 217 To 240
            SeleccionarNumero15 = 10
    
        Case 241 To 264
            SeleccionarNumero15 = 11
        
        Case 265 To 289
            SeleccionarNumero15 = 12
        
        Case 217 To 360
            SeleccionarNumero15 = 13
        
        Case 217 To 400
            SeleccionarNumero15 = 14
    
        Case 271 To 300
            SeleccionarNumero15 = 15
    
        Case 301 To 330
            SeleccionarNumero15 = 16
    
        Case 331 To 360
            SeleccionarNumero15 = 17
        
        Case 361 To 390
            SeleccionarNumero15 = 18
    
        Case 391 To 420
            SeleccionarNumero15 = 19
    
        Case 421 To 450
            SeleccionarNumero15 = 20
    
    End Select
    
    Call temp(0, 0, 0, 0, (vNumeroInicial + SeleccionarNumero15))

If Err Then GrabarLog "SeleccionarNumero", Err.Number & " " & Err.Description, Me.Name
End Function
Private Function SeleccionarNumero24(vCantidad As Integer) As Integer
On Error Resume Next
    
    Select Case vCantidad
    
        Case 1 To 24
            SeleccionarNumero24 = 1
        
        Case 25 To 48
            SeleccionarNumero24 = 2
            
        Case 49 To 72
            SeleccionarNumero24 = 3
        
        Case 73 To 96
            SeleccionarNumero24 = 4
    
        Case 97 To 120
            SeleccionarNumero24 = 5
        
        Case 121 To 144
            SeleccionarNumero24 = 6
        
        Case 145 To 168
            SeleccionarNumero24 = 7

        Case 169 To 192
            SeleccionarNumero24 = 8

        Case 193 To 216
            SeleccionarNumero24 = 9

        Case 217 To 240
            SeleccionarNumero24 = 10
    
        Case 241 To 264
            SeleccionarNumero24 = 11
        
        Case 265 To 289
            SeleccionarNumero24 = 12
        
        Case 217 To 360
            SeleccionarNumero24 = 13
        
        Case 217 To 400
            SeleccionarNumero24 = 14
    
        Case 271 To 300
            SeleccionarNumero24 = 15
    
        Case 301 To 330
            SeleccionarNumero24 = 16
    
        Case 331 To 360
            SeleccionarNumero24 = 17
        
        Case 361 To 390
            SeleccionarNumero24 = 18
    
        Case 391 To 420
            SeleccionarNumero24 = 19
    
        Case 421 To 450
            SeleccionarNumero24 = 20
    
    End Select
    
    'Call Temp(0, 0, 0, 0, (vNumeroInicial + SeleccionarNumero24))

If Err Then GrabarLog "SeleccionarNumero24", Err.Number & " " & Err.Description, Me.Name
End Function
Private Function CalcularTotal(vtipo, vValor) As Double
On Error Resume Next

    Select Case vtipo
    
        Case ""
            CalcularTotal = Val(GenerarDato("SELECT Sum(Factura.Subtotal) AS Neto, Sum(IvaFacturaVenta.Iva105) AS I105, Sum(IvaFacturaVenta.Iva210) AS I210, Sum(IvaFacturaVenta.Iva270) AS I270, Neto+I105+I210+I270 AS Total FROM Factura INNER JOIN IvaFacturaVenta ON Factura.Remito = IvaFacturaVenta.Remito GROUP BY Month(Factura.Fecha), Year(Factura.Fecha) HAVING (Month(Factura.Fecha) = '" & AjustarMes(Month(dtpFecha(0).Value)) & "') AND (Year(Factura.Fecha) = '" & Year(dtpFecha(1).Value) & "')", vValor))
        
        Case "Fact A"
            CalcularTotal = Val(GenerarDato("SELECT Fecha, Factura.tipo, Sum(Factura.Subtotal) AS Neto, Sum(IvaFacturaVenta.Iva105) AS I105, Sum(IvaFacturaVenta.Iva210) AS I210, Sum(IvaFacturaVenta.Iva270) AS I270, Sum(Factura.Subtotal)+Sum(IvaFacturaVenta.Iva105)+Sum(IvaFacturaVenta.Iva210)+Sum(IvaFacturaVenta.Iva270) AS Total FROM Factura INNER JOIN IvaFacturaVenta ON Factura.Remito = IvaFacturaVenta.Remito GROUP BY Factura.tipo, Month(Factura.Fecha), Year(Factura.Fecha) HAVING (Factura.tipo = '" & vtipo & "') AND (Month(Factura.Fecha) = " & AjustarMes(Month(dtpFecha(0).Value)) & ") AND (Year(Factura.Fecha) = " & Year(dtpFecha(1).Value) & ")", vValor))
        
        Case "Fact B"
            Select Case vValor
            
                Case "Neto"
                    'CalcularTotal = Val(GenerarDato("SELECT Fecha, Factura.tipo, Sum(Factura.Subtotal) AS Neto, Sum(IvaFacturaVenta.Iva105) AS I105, Sum(IvaFacturaVenta.Iva210) AS I210, Sum(IvaFacturaVenta.Iva270) AS I270, Sum(Factura.Subtotal)+Sum(IvaFacturaVenta.Iva105)+Sum(IvaFacturaVenta.Iva210)+Sum(IvaFacturaVenta.Iva270) AS Total FROM Factura INNER JOIN IvaFacturaVenta ON Factura.Remito = IvaFacturaVenta.Remito GROUP BY Factura.tipo, Month(Factura.Fecha), Year(Factura.Fecha) HAVING (Factura.tipo = '" & vTipo & "') AND (Month(Factura.Fecha) = " & Val(cboMes.Text) & ") AND (Year(Factura.Fecha) = " & Val(txtAno.Text) & ")", vValor))
                
                Case "I105"
                    'CalcularTotal = Val(GenerarDato("SELECT Fecha, Tipo, Sum(Subtotal) AS Neto, Sum(IvaFacturaVenta.Iva105) AS I105, Sum(IvaFacturaVenta.Iva210) AS I210, Sum(IvaFacturaVenta.Iva270) AS I270, Sum(Factura.Subtotal)+Sum(IvaFacturaVenta.Iva105)+Sum(IvaFacturaVenta.Iva210)+Sum(IvaFacturaVenta.Iva270) AS Total FROM Factura INNER JOIN IvaFacturaVenta ON Factura.Remito = IvaFacturaVenta.Remito GROUP BY Tipo, Month(Fecha), Year(Fecha) HAVING (Tipo = '" & vTipo & "') AND (Month(Fecha) = " & Val(cboMes.Text) & ") AND (Year(Fecha) = " & Val(txtAno.Text) & ")", vValor))
                
                Case "I210"
                    'CalcularTotal = Val(GenerarDato("SELECT Fecha, Factura.tipo, Sum(Factura.Subtotal) AS Neto, Sum(IvaFacturaVenta.Iva105) AS I105, Sum(IvaFacturaVenta.Iva210) AS I210, Sum(IvaFacturaVenta.Iva270) AS I270, Sum(Factura.Subtotal)+Sum(IvaFacturaVenta.Iva105)+Sum(IvaFacturaVenta.Iva210)+Sum(IvaFacturaVenta.Iva270) AS Total FROM Factura INNER JOIN IvaFacturaVenta ON Factura.Remito = IvaFacturaVenta.Remito GROUP BY Factura.tipo, Month(Factura.Fecha), Year(Factura.Fecha) HAVING (Factura.tipo = '" & vTipo & "') AND (Month(Factura.Fecha) = " & Val(cboMes.Text) & ") AND (Year(Factura.Fecha) = " & Val(txtAno.Text) & ")", vValor))
                
                Case "I270"
                    'CalcularTotal = Val(GenerarDato("SELECT Fecha, Factura.tipo, Sum(Factura.Subtotal) AS Neto, Sum(IvaFacturaVenta.Iva105) AS I105, Sum(IvaFacturaVenta.Iva210) AS I210, Sum(IvaFacturaVenta.Iva270) AS I270, Sum(Factura.Subtotal)+Sum(IvaFacturaVenta.Iva105)+Sum(IvaFacturaVenta.Iva210)+Sum(IvaFacturaVenta.Iva270) AS Total FROM Factura INNER JOIN IvaFacturaVenta ON Factura.Remito = IvaFacturaVenta.Remito GROUP BY Factura.tipo, Month(Factura.Fecha), Year(Factura.Fecha) HAVING (Factura.tipo = '" & vTipo & "') AND (Month(Factura.Fecha) = " & Val(cboMes.Text) & ") AND (Year(Factura.Fecha) = " & Val(txtAno.Text) & ")", vValor))
                
                Case "Total"
                    'CalcularTotal = Val(GenerarDato("SELECT Fecha, Factura.tipo, Sum(Factura.Subtotal) AS Neto, Sum(IvaFacturaVenta.Iva105) AS I105, Sum(IvaFacturaVenta.Iva210) AS I210, Sum(IvaFacturaVenta.Iva270) AS I270, Sum(Factura.Subtotal)+Sum(IvaFacturaVenta.Iva105)+Sum(IvaFacturaVenta.Iva210)+Sum(IvaFacturaVenta.Iva270) AS Total FROM Factura INNER JOIN IvaFacturaVenta ON Factura.Remito = IvaFacturaVenta.Remito GROUP BY Factura.tipo, Month(Factura.Fecha), Year(Factura.Fecha) HAVING (Factura.tipo = '" & vTipo & "') AND (Month(Factura.Fecha) = " & Val(cboMes.Text) & ") AND (Year(Factura.Fecha) = " & Val(txtAno.Text) & ")", vValor))
            
            End Select
            
    
    End Select
    
If Err Then GrabarLog "CalcularTotal", Err.Number & " " & Err.Description, Me.Name
End Function

Private Sub FormatoGrillaTotales(vCantidadRenglones As Integer)
On Error Resume Next

    Dim i As Integer

    With Me.klexTotales
        .FixedRows = 1
        .FixedCols = 1
    
        .Cols = 12
        .Rows = vCantidadRenglones + 1
        
        If vCantidadRenglones = 1 Then
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
                .ColWidth(i) = 0
            Next
            .BackColorAlternate = &HE0E0E0
        Else

            .Row = .Rows - 1
            For i = 1 To .Cols - 1
                .CellBackColor = &HFFFCCC
                .Col = i
            Next
        End If
        
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 125
        
        'Aca Pego el IdFDetalle-Entonces se si modifico o NO
        .TextMatrix(0, 1) = "Tipo"
        .ColWidth(1) = 750
        
        .TextMatrix(0, 2) = "Neto"
        .ColWidth(2) = 1000
        .ColDisplayFormat(2) = "###,##0.00"
        
        .TextMatrix(0, 3) = "Iva10.5"
        .ColWidth(3) = 1000
        .ColDisplayFormat(3) = "###,##0.00"
        
        .TextMatrix(0, 4) = "Iva21"
        .ColWidth(4) = 1000
        .ColDisplayFormat(4) = "###,##0.00"
        
        .TextMatrix(0, 5) = "Iva27"
        .ColWidth(5) = 1000
        .ColDisplayFormat(5) = "###,##0.00"
        
        .TextMatrix(0, 6) = "Ret."
        .ColWidth(6) = 1000
        .ColDisplayFormat(6) = "###,##0.00"
        
        .TextMatrix(0, 7) = "Perc."
        .ColWidth(7) = 1000
        .ColDisplayFormat(7) = "###,##0.00"
        
        .TextMatrix(0, 8) = "NoGrab."
        .ColWidth(8) = 750
        .ColDisplayFormat(8) = "###,##0.00"
        
        .TextMatrix(0, 9) = "ITC"
        .ColWidth(9) = 750
        .ColDisplayFormat(9) = "###,##0.00"
                
        .TextMatrix(0, 10) = "Exento"
        .ColWidth(10) = 750
        .ColDisplayFormat(10) = "###,##0.00"
        
        .TextMatrix(0, 11) = "Total"
        .ColWidth(11) = 1000
        .ColDisplayFormat(11) = "###,##0.00"
        
        .Editable = False

        .EnterKeyBehaviour = klexEKNone


        .Row = 2
    End With
    
If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub CargarGrillaTotales()
On Error Resume Next
    
    Dim rsTotales As New ADODB.Recordset, sqlTotales As String, i As Integer
    
    'sqlTotales = "SELECT TipoMovimiento, SUM(SubTotal) as ImpBruto,SUM(Iva105) as ImpIva105,SUM(Iva210) as ImpIva210, SUM(Iva270) as ImpIva270, SUM(Retenciones) as ImpRetenciones,SUM(Percepciones) as ImpPercepciones,SUM(NoGravado) as ImpNoGravado, SUM(ITC) as ImpITC, SUM(ImpExento) as ImpExento,  SUM(Total) as ImpNeto FROM Factura Fa INNER JOIN IvaFacturaVenta Iv ON Fa.Remito=IV.Remito WHERE Month(fecha) =  '" & AjustarMes(Month(dtpFecha(0).Value)) & "' And Year(fecha) = '" & Year(dtpFecha(1).Value) & "' GROUP BY TipoMovimiento;"
    sqlTotales = "SELECT TipoMovimiento, SUM(Neto) as ImpBruto,SUM(Iva105) as ImpIva105,SUM(Iva210) as ImpIva210, SUM(Iva270) as ImpIva270, SUM(Retenciones) as ImpRetenciones,SUM(Percepciones) as ImpPercepciones,SUM(NoGravado) as ImpNoGravado, SUM(ITC) as ImpITC, SUM(ImpExento) as ImpExento,  SUM(Total) as ImpNeto FROM temp_IVA WHERE (not codigo = 'Anulado') and Month(fecha) =  '" & AjustarMes(Month(dtpFecha(0).Value)) & "' And Year(fecha) = '" & Year(dtpFecha(1).Value) & "' GROUP BY TipoMovimiento;"
      
    With rsTotales
        .CursorLocation = adUseClient
        
        Call .Open(sqlTotales, ConnDDBB, adOpenStatic, adLockBatchOptimistic)
        
        If Not .EOF = True Then .MoveFirst
        
        
        FormatoGrillaTotales (.RecordCount + 1)
        
        i = 0
        
        For i = 2 To 11
            klexTotales.ColAlignment(i) = 6
        Next
        
        
        i = 0
        
        
        Do Until .EOF = True
            klexTotales.TextMatrix(.AbsolutePosition, 1) = EsNulo(.Fields("TipoMovimiento").Value)
            klexTotales.TextMatrix(.AbsolutePosition, 2) = EsNulo(.Fields("ImpBruto").Value)
            klexTotales.TextMatrix(.AbsolutePosition, 3) = EsNulo(.Fields("ImpIva105").Value)
            klexTotales.TextMatrix(.AbsolutePosition, 4) = EsNulo(.Fields("ImpIva210").Value)
            klexTotales.TextMatrix(.AbsolutePosition, 5) = EsNulo(.Fields("ImpIva270").Value)
            klexTotales.TextMatrix(.AbsolutePosition, 6) = EsNulo(.Fields("ImpRetenciones").Value)
            klexTotales.TextMatrix(.AbsolutePosition, 7) = EsNulo(.Fields("ImpPercepciones").Value)
            klexTotales.TextMatrix(.AbsolutePosition, 8) = EsNulo(.Fields("ImpNoGravado").Value)
            klexTotales.TextMatrix(.AbsolutePosition, 9) = EsNulo(.Fields("ImpITC").Value)
            klexTotales.TextMatrix(.AbsolutePosition, 10) = EsNulo(.Fields("ImpExento").Value)
            klexTotales.TextMatrix(.AbsolutePosition, 11) = EsNulo(.Fields("ImpNeto").Value)
            
            .MoveNext
            i = i + 1
        Loop
        
        '--------------------------------------------------------------------------------------------------------------------------------------------------------
        
        'Cargo Los Totales dentro de la misma Grilla
        .Close
        
        'sqlTotales = "SELECT SUM(SubTotal) as ImpBruto,SUM(Iva105) as ImpIva105,SUM(Iva210) as ImpIva210, SUM(Iva270) as ImpIva270, SUM(Retenciones) as ImpRetenciones,SUM(Percepciones) as ImpPercepciones,SUM(NoGravado) as ImpNoGravado, SUM(ITC) as ImpITC, SUM(ImpExento) as ImpExento,  SUM(Total) as ImpNeto FROM Factura Fa INNER JOIN IvaFacturaVenta Iv ON Fa.Remito=IV.Remito WHERE Month(fecha) =  '" & AjustarMes(Month(dtpFecha(0).Value)) & "' And Year(fecha) = '" & Year(dtpFecha(1).Value) & "'"
        sqlTotales = "SELECT SUM(Neto) as ImpBruto,SUM(Iva105) as ImpIva105,SUM(Iva210) as ImpIva210, SUM(Iva270) as ImpIva270, SUM(Retenciones) as ImpRetenciones,SUM(Percepciones) as ImpPercepciones,SUM(NoGravado) as ImpNoGravado, SUM(ITC) as ImpITC, SUM(ImpExento) as ImpExento,  SUM(Total) as ImpNeto FROM temp_IVA WHERE Month(fecha) =  '" & AjustarMes(Month(dtpFecha(0).Value)) & "' And Year(fecha) = '" & Year(dtpFecha(1).Value) & "'"

        
        Call .Open(sqlTotales, ConnDDBB, adOpenStatic, adLockBatchOptimistic)
    
        klexTotales.Row = i + 1
                

        klexTotales.TextMatrix(i + 1, 1) = "Totales :"
        klexTotales.TextMatrix(i + 1, 2) = EsNulo(.Fields("ImpBruto").Value)
        klexTotales.TextMatrix(i + 1, 3) = EsNulo(.Fields("ImpIva105").Value)
        klexTotales.TextMatrix(i + 1, 4) = EsNulo(.Fields("ImpIva210").Value)
        klexTotales.TextMatrix(i + 1, 5) = EsNulo(.Fields("ImpIva270").Value)
        klexTotales.TextMatrix(i + 1, 6) = EsNulo(.Fields("ImpRetenciones").Value)
        klexTotales.TextMatrix(i + 1, 7) = EsNulo(.Fields("ImpPercepciones").Value)
        klexTotales.TextMatrix(i + 1, 8) = EsNulo(.Fields("ImpNoGravado").Value)
        klexTotales.TextMatrix(i + 1, 9) = EsNulo(.Fields("ImpITC").Value)
        klexTotales.TextMatrix(i + 1, 10) = EsNulo(.Fields("ImpExento").Value)
        klexTotales.TextMatrix(i + 1, 11) = EsNulo(.Fields("ImpNeto").Value)

    End With
    
    sqlTotales = ""
    
    If rsTotales.State = 1 Then
        rsTotales.Close
        Set rsTotales = Nothing
    End If
    
If Err Then GrabarLog "CargarGrillaTotales", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub dtpFecha_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        
        Select Case Index
        
            Case 0
                dtpFecha(1).MaxValor = DiasDelMes(dtpFecha(0).Value) & "/" & AjustarMes(Month(dtpFecha(0).Value)) & "/" & Year(dtpFecha(0).Value)
                dtpFecha(1).Value = dtpFecha(1).MaxValor
                dtpFecha(1).SetFocus
                
            Case 1
                PbAcciones(0).SetFocus
            
        End Select
        
    End If

If Err Then GrabarLog "txtFecha_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub KlexFacturas_Click()
'Dim i As Integer

'Me.v105 = Me.KlexFacturas.TextMatrix(i, 10)

End Sub

Private Sub KlexFacturas_DblClick()
Dim i As Integer

With KlexFacturas
    i = .Row
    .Col = 0
    
If .CellBackColor = vbRed Then
    .CellBackColor = vbYellow
Else
    .CellBackColor = vbRed
End If

End With


End Sub

Private Sub PbAcciones_Click(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 0
            GenerarListado
            
            GenerarListadoPorLocalidad
            
        Case 1
            GuardarIva
        Case 2
            Reporte
        
        Case 3
            Unload Me
    
    End Select
If Err Then GrabarLog "PbAcciones_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub GenerarListadoPorLocalidad()
Dim vsql
Dim r As New ADODB.Recordset

vsql = "select " + _
" f.Localidad,sum(ti.Neto) as Neto,sum(ti.Iva105) as Iva105,sum(ti.Iva210) as Iva21,sum(ti.Iva270) as Iva27,sum(ti.Total) as Total " + _
" from clientes f inner join temp_iva ti " + _
" on ti.codigo = f.codigo  where not ti.codigo = 'Anulado'" + _
" group by f.Localidad"


Call r.Open(vsql, ConnDDBB, adOpenStatic, adLockReadOnly)

Set Me.gridTLocalidad.Recordset = r

vsql = "select " + _
" 'Total                 : ',sum(ti.Neto) as Neto,sum(ti.Iva105) as Iva105,sum(ti.Iva210) as Iva21,sum(ti.Iva270) as Iva27,sum(ti.Total) as Total " + _
" from clientes f inner join temp_iva ti " + _
" on ti.codigo = f.codigo where not ti.codigo = 'Anulado'"


r.Close
Call r.Open(vsql, ConnDDBB, adOpenStatic, adLockReadOnly)

Set Me.gridTotalesLocalidad.Recordset = r

End Sub


Private Sub PushButton1_Click()
Call ImprimirFlex(Me.klexTotales, "Totales IVA Compra", "")
End Sub

Private Sub PushButton2_Click()
On Error Resume Next
    
  Call grillaToExcel2(Me.KlexFacturas)

If Err Then Exit Sub
End Sub

Private Sub PushButton3_Click()
sacarPoner ("Anulado")
End Sub


Private Sub sacarPoner(vset As String)
Dim i As Integer
Dim vsql, vsql2, vmen As String

'vsql = "update factura set cae='" + Trim(vset) + "' where remito = "
'vsql2 = vsql = "update temp_iva set Codigo='" + Trim(vset) + "' where remito = "



With Me.KlexFacturas
b1.Max = .Rows - 1
b1.Value = 0
        For i = 1 To .Rows - 1
                b1.Value = b1.Value + 1
                .Col = 0
                .Row = i
                If .CellBackColor = vbRed Or .CellBackColor = vbYellow Then
                vsql = "update factura set repartidor='" + Trim(vset) + "' where remito = "
                vsql2 = "update temp_iva set Codigo='" + Trim(vset) + "' where remito = "
                    vsql = vsql + .TextMatrix(i, 0)
                    vsql2 = vsql2 + .TextMatrix(i, 0)
                    Call EjecutarScript(vsql, pathDBMySQL)
                    Call EjecutarScript(vsql2, pathDBMySQL)
                    Debug.Print vsql2
                Else
                
                    Debug.Print "- No Pintado -> " + Str(.TextMatrix(i, 0))
                
                End If
        Next
End With


vmen = "Debe volver a generar el listado para actualizar los archivos del CITI" + Chr(13) + _
"Quire hacerlo ahora ? "

If MsgBox(vmen, vbYesNo) = vbYes Then
    Call PbAcciones_Click(0)
End If



End Sub

Private Sub PushButton4_Click()
sacarPoner ("")
End Sub

Private Sub PushButton5_Click()
    GenerarListadoPorLocalidad
End Sub

Private Sub PushButton7_Click()
Dim vsql, vc1, vc2 As String

vsql = "(Select * from proveedores where tipocliente  = 'Vendedor') t"
vc1 = "Nombre"
vc2 = "Codigo"


Call fbuscarGrilla(vsql, vc1, vc2, Me.vdescEmpresa.Name, Me)
End Sub

Private Sub PushButton6_Click()
Dim vsql, vc1, vc2 As String

vsql = "(Select * from proveedores where tipocliente  = 'Empresa') t"
vc1 = "Nombre"
vc2 = "Codigo"

Call fbuscarGrilla(vsql, vc1, vc2, Me.vdescEmpresa.Name, Me)

End Sub
