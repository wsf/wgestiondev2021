VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "Copia de KlexGrid.ocx"
Object = "{9746E3DA-06E1-4D26-9CE4-D9F6411A9C70}#1.0#0"; "SMGA_OcxTxt2008.ocx"
Begin VB.Form frmIvaCompra 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Libro IVA Compra"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   180
   ClientWidth     =   13125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   13125
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   525
      Left            =   30
      TabIndex        =   10
      Top             =   -90
      Width           =   11955
      _Version        =   851968
      _ExtentX        =   21087
      _ExtentY        =   926
      _StockProps     =   79
      Appearance      =   1
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   345
         Left            =   4380
         TabIndex        =   15
         Top             =   120
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Excel"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   3
         Left            =   10440
         TabIndex        =   11
         Top             =   120
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2558
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Cerrar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmIvaCompra.frx":0000
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   0
         Left            =   30
         TabIndex        =   12
         Top             =   120
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2558
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Generar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmIvaCompra.frx":0400
         BorderGap       =   10
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   1
         Left            =   1470
         TabIndex        =   13
         Top             =   120
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2558
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Definitivo"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmIvaCompra.frx":083A
         BorderGap       =   10
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   2
         Left            =   2910
         TabIndex        =   14
         Top             =   120
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2558
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmIvaCompra.frx":0C4E
         BorderGap       =   10
      End
   End
   Begin XtremeSuiteControls.GroupBox GBParametros 
      Height          =   675
      Left            =   0
      TabIndex        =   4
      Top             =   420
      Width           =   11985
      _Version        =   851968
      _ExtentX        =   21140
      _ExtentY        =   1191
      _StockProps     =   79
      Caption         =   "Parametros"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.CheckBox chkTotales 
         Height          =   255
         Left            =   8700
         TabIndex        =   9
         Top             =   300
         Visible         =   0   'False
         Width           =   3075
         _Version        =   851968
         _ExtentX        =   5424
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Solo Calcular Totales del Periodo"
         Appearance      =   6
      End
      Begin Aplisoft_CajasDeTexto.TxF dtpFecha 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   5
         Top             =   285
         Width           =   1335
         _ExtentX        =   2355
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
         Left            =   4770
         TabIndex        =   6
         Top             =   285
         Width           =   1305
         _ExtentX        =   2302
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
         Caption         =   "> Periodo Final : "
         Height          =   195
         Index           =   1
         Left            =   3450
         TabIndex        =   8
         Top             =   315
         Width           =   1185
      End
      Begin VB.Label lblDatos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "> Periodo Inicial :"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   315
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc bIvaCompra 
      Height          =   330
      Left            =   0
      Top             =   6600
      Visible         =   0   'False
      Width           =   10605
      _ExtentX        =   18706
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
      Caption         =   "bIvaCompra"
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
   Begin MSAdodcLib.Adodc bPFactura 
      Height          =   330
      Left            =   0
      Top             =   6600
      Visible         =   0   'False
      Width           =   10605
      _ExtentX        =   18706
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
   Begin MSAdodcLib.Adodc bTemp_Iva 
      Height          =   330
      Left            =   0
      Top             =   6600
      Visible         =   0   'False
      Width           =   10605
      _ExtentX        =   18706
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
   Begin MSAdodcLib.Adodc bLiquidoProducto 
      Height          =   330
      Left            =   0
      Top             =   6600
      Visible         =   0   'False
      Width           =   10605
      _ExtentX        =   18706
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
      Caption         =   "bLiquidoProducto"
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
   Begin MSComctlLib.ProgressBar Barra 
      Height          =   195
      Left            =   -60
      TabIndex        =   0
      Top             =   5760
      Width           =   12045
      _ExtentX        =   21246
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.TabControl TabIva 
      Height          =   4575
      Left            =   -60
      TabIndex        =   1
      Top             =   1140
      Width           =   12015
      _Version        =   851968
      _ExtentX        =   21193
      _ExtentY        =   8070
      _StockProps     =   68
      PaintManager.Layout=   4
      PaintManager.BoldSelected=   -1  'True
      ItemCount       =   2
      Item(0).Caption =   "Listado"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "KlexFacturas"
      Item(1).Caption =   "Totales"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "klexTotales"
      Item(1).Control(1)=   "PushButton2"
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   225
         Left            =   -60730
         TabIndex        =   16
         Top             =   390
         Visible         =   0   'False
         Width           =   2625
         _Version        =   851968
         _ExtentX        =   4630
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Imprimir Totales"
         UseVisualStyle  =   -1  'True
      End
      Begin Grid.KlexGrid klexTotales 
         Height          =   3915
         Left            =   -69850
         TabIndex        =   2
         Top             =   660
         Visible         =   0   'False
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   6906
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
         MouseIcon       =   "frmIvaCompra.frx":105C
         Rows            =   10
      End
      Begin Grid.KlexGrid KlexFacturas 
         Height          =   3975
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   7011
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
         MouseIcon       =   "frmIvaCompra.frx":1078
         Rows            =   10
      End
   End
End
Attribute VB_Name = "frmIvaCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vtipo As String
Dim vtitulo As String
Dim fdesde  As String
Dim fhasta As String
Dim vNumeroPagina As Integer, vNumeroInicial As Integer

Dim vNetoFactA As Double, vI105FactA As Double, vI210FactA As Double, vI270FactA As Double, vTotalFactA As Double
Dim vNetoMonotributo As Double, vI105Monotributo As Double, vI210Monotributo As Double, vI270Monotributo As Double, vTotalMonotributo As Double
Dim vNetoNotaC As Double, vI105NotaC As Double, vI210NotaC As Double, vI270NotaC As Double, vTotalNotaC As Double
Dim vNetoNotaD As Double, vI105NotaD As Double, vI210NotaD As Double, vI270NotaD As Double, vTotalNotaD As Double


Dim vgline, vgline2 As Integer


Dim Canal1%, Canal22%

Dim vTipoIva, vNGtotal As Double
Private Sub CopiarTemp(vModo As Byte)
On Error Resume Next


Dim vlb, vla, vla2 As String
Dim vIdTipoIva As String
Dim b1 As Integer

Dim vtipo As String
Dim vPercepciones As Double, vRetenciones As Double, vITC As Double, vNoGravado As Double, vImpExento As Double
Dim ponerenciti As Boolean

                

    With bTemp_Iva
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM Temp_Iva"
        .Refresh
        
        .Recordset.AddNew
        
        Select Case vModo
            
            Case 0
            
            
                .Recordset("Fecha").Value = strfechaMySQL(bPFactura.Recordset("Fecha").Value)
                
                '.Recordset("Tipo").Value = EsNulo(bPFactura.Recordset("Tipo").Value)
                vtipo = EsNulo(bPFactura.Recordset("Tipo").Value)
        
                .Recordset("TipoMovimiento").Value = vtipo 'EsNulo(bPFactura.Recordset("TipoMovimiento").Value)
                .Recordset("Letra").Value = EsNulo(bPFactura.Recordset("Letra").Value)
                .Recordset("NroComprobante").Value = FormatoFactura(8, bPFactura.Recordset("NComprobante").Value)
                .Recordset("PuntoDeVenta").Value = FormatoFactura(4, bPFactura.Recordset("PuntoDeVenta").Value)
                
                .Recordset("Codigo").Value = EsNulo(bPFactura.Recordset("Codigo").Value)
                .Recordset("Nombre").Value = EsNulo(bPFactura.Recordset("Nombre").Value)
                .Recordset("Cuit").Value = EsNulo(bPFactura.Recordset("CUIT").Value)
                
                ponerenciti = True
                
                If EsNulo(bPFactura.Recordset("estadodocumento").Value) = "Anulado" Then
                     ponerenciti = False
                    .Recordset("Nombre").Value = EsNulo(bPFactura.Recordset("Nombre").Value) + "(Anulado)"
                                   
                Else
                
                
            
                    
                    .Recordset("Neto").Value = bPFactura.Recordset("SubTotal").Value
                    
                    .Recordset("IVA105").Value = TraerDato("IvaFacturaCompra ", "remito = " & Str(bPFactura.Recordset("remito").Value) + " order by idivafacturacompra desc", "Iva105")
                    .Recordset("IVA210").Value = TraerDato("IvaFacturaCompra ", "remito = " & Str(bPFactura.Recordset("remito").Value) + " order by idivafacturacompra desc", "Iva210")
                    .Recordset("IVA270").Value = TraerDato("IvaFacturaCompra ", "remito = " & Str(bPFactura.Recordset("remito").Value) + " order by idivafacturacompra desc", "Iva270")
                    
                    
                    .Recordset("IVA105").Value = TraerDato("IvaFacturaCompra ", "remito = " & Str(bPFactura.Recordset("remito").Value) + " order by idivafacturacompra desc", "Iva105")
                    .Recordset("IVA210").Value = TraerDato("IvaFacturaCompra ", "remito = " & Str(bPFactura.Recordset("remito").Value) + " order by idivafacturacompra desc", "Iva210")
                    .Recordset("IVA270").Value = TraerDato("IvaFacturaCompra ", "remito = " & Str(bPFactura.Recordset("remito").Value) + " order by idivafacturacompra desc", "Iva270")
                    
                    
                    
                    .Recordset("Nombre").Value = EsNulo(bPFactura.Recordset("Nombre").Value)
         
                            
                
                '------------------------------
                
                
                vPercepciones = 0
                vRetenciones = 0
                vITC = 0
                vNoGravado = 0
                vImpExento = 0
                
                
                vRetenciones = TraerDato("IvaFacturaCompra", "remito = " & Str(bPFactura.Recordset("remito").Value) + " order by idivafacturacompra desc ", "Retenciones")
                vPercepciones = TraerDato("IvaFacturaCompra", "remito = " & Str(bPFactura.Recordset("remito").Value) + " order by idivafacturacompra desc ", "Percepciones")
                vNoGravado = TraerDato("IvaFacturaCompra", "remito = " & Str(bPFactura.Recordset("remito").Value) + " order by idivafacturacompra desc ", "Nogravado")
                vImpExento = TraerDato("IvaFacturaCompra", "remito = " & Str(bPFactura.Recordset("remito").Value) + " order by idivafacturacompra desc ", "ImpExento")
                    
 
                vRetenciones = TraerDato("IvaFacturaCompra", "remito = " & Str(bPFactura.Recordset("remito").Value) + " order by idivafacturacompra desc ", "Retenciones")
                vPercepciones = TraerDato("IvaFacturaCompra", "remito = " & Str(bPFactura.Recordset("remito").Value) + " order by idivafacturacompra desc ", "Percepciones")
                vNoGravado = TraerDato("IvaFacturaCompra", "remito = " & Str(bPFactura.Recordset("remito").Value) + " order by idivafacturacompra desc ", "Nogravado")
                vImpExento = TraerDato("IvaFacturaCompra", "remito = " & Str(bPFactura.Recordset("remito").Value) + " order by idivafacturacompra desc ", "ImpExento")
                  





                .Recordset("Percepciones").Value = vPercepciones 'IIf(vPercepciones > 0, vPercepciones * -1, vPercepciones)
                .Recordset("Retenciones").Value = vRetenciones 'IIf(vRetenciones > 0, vRetenciones * -1, vRetenciones)
                ' rsTemp_Iva.Fields("vITC").Value = IIf(vITC > 0, vITC * -1, vITC)
                .Recordset("Nogravado").Value = vNoGravado 'IIf(vNoGravado > 0, vNoGravado * -1, vNoGravado)
                .Recordset("ImpExento").Value = vImpExento 'IIf(vImpExento > 0, vImpExento * -1, vImpExento)
                
                
                '-------------------------------
                
                
                '.Recordset("Retenciones").Value = TraerDato("IvaFacturaCompra", "remito = " & bPFactura.Recordset("remito").Value, "Retenciones")
                '.Recordset("Percepciones").Value = TraerDato("IvaFacturaCompra", "remito = " & bPFactura.Recordset("remito").Value, "Percepciones")
                '.Recordset("Nogravado").Value = TraerDato("IvaFacturaCompra", "remito = " & bPFactura.Recordset("remito").Value & "", "Nogravado")
                '.Recordset("ImpExento").Value = TraerDato("IvaFacturaCompra", "remito = " & bPFactura.Recordset("remito").Value & "", "ImpExento")
                
                
                
                
               vtipo = Trim(.Recordset("Tipo").Value)
                
                If vtipo = "Nota C" Then
                
                        .Recordset("Total").Value = -1 * bPFactura.Recordset("Total").Value
                        .Recordset("Neto").Value = -1 * bPFactura.Recordset("SubTotal").Value
                        .Recordset("IVA105").Value = -1 * TraerDato("IvaFacturaCompra", "remito = " & Str(bPFactura.Recordset("remito").Value) + " order by remito desc ", "Iva105")
                        .Recordset("IVA210").Value = -1 * TraerDato("IvaFacturaCompra", "remito = " & Str(bPFactura.Recordset("remito").Value) + " order by remito desc ", "Iva210")
                        .Recordset("IVA270").Value = -1 * TraerDato("IvaFacturaCompra", "remito = " & Str(bPFactura.Recordset("remito").Value) + " order by remito desc ", "Iva270")
                        .Recordset("Nogravado").Value = IIf(vNoGravado > 0, vNoGravado * -1, vNoGravado)
                        
                        .Recordset("Percepciones").Value = IIf(vPercepciones > 0, vPercepciones * -1, vPercepciones)
                        .Recordset("Retenciones").Value = IIf(vRetenciones > 0, vRetenciones * -1, vRetenciones)
                        ' rsTemp_Iva.Fields("vITC").Value = IIf(vITC > 0, vITC * -1, vITC)
                        .Recordset("Nogravado").Value = IIf(vNoGravado > 0, vNoGravado * -1, vNoGravado)
                        
                        
                        
                        .Recordset("ImpExento").Value = IIf(vImpExento > 0, vImpExento * -1, vImpExento)
                        
                        
                Else
                       If Not vtipo = "NC" Then .Recordset("Total").Value = bPFactura.Recordset("Total").Value
                
                End If
                
                vNGtotal = vNGtotal + .Recordset("Nogravado").Value
                '.Recordset("Total").Value = bPFactura.Recordset("Total").Value
            End If

            Case 1
                .Recordset("NroComprobante").Value = FormatoNC(bPFactura.Recordset("Tipo").Value, bPFactura.Recordset("NComprobante").Value)
                .Recordset("Nombre").Value = "ANULADA"
        
            Case 2
                .Recordset("Fecha").Value = fhasta
                .Recordset("Tipo").Value = bPFactura.Recordset("Tipo").Value
                .Recordset("NroComprobante").Value = FormatoNC(vtipo, bLiquidoProducto.Recordset("NroLiquido").Value)
                .Recordset("Codigo").Value = bLiquidoProducto.Recordset("Codigo").Value
                .Recordset("Nombre").Value = bLiquidoProducto.Recordset("Nombre").Value
                .Recordset("Cuit").Value = bLiquidoProducto.Recordset("CUIT").Value
                .Recordset("Neto").Value = bLiquidoProducto.Recordset("SubTotal").Value
                .Recordset("IVA").Value = bLiquidoProducto.Recordset("IvaTotal").Value
                .Recordset("Total").Value = bLiquidoProducto.Recordset("Total").Value

            Case 3
        
        End Select
       
        vNumeroPagina = vNumeroInicial + SeleccionarNumero(.Recordset.RecordCount)
           
        .Recordset("NumHoja").Value = vNumeroPagina
        .Recordset.Update
        
        
        KlexFacturas.TextMatrix(bPFactura.Recordset.AbsolutePosition, 1) = strfechaMySQL(.Recordset("Fecha").Value)
        KlexFacturas.TextMatrix(bPFactura.Recordset.AbsolutePosition, 2) = EsNulo(.Recordset("TipoMovimiento").Value)
        KlexFacturas.TextMatrix(bPFactura.Recordset.AbsolutePosition, 3) = EsNulo(.Recordset("Letra").Value)
        KlexFacturas.TextMatrix(bPFactura.Recordset.AbsolutePosition, 4) = EsNulo(.Recordset("PuntoDeVenta").Value)
        KlexFacturas.TextMatrix(bPFactura.Recordset.AbsolutePosition, 5) = EsNulo(.Recordset("NroComprobante").Value)
        KlexFacturas.TextMatrix(bPFactura.Recordset.AbsolutePosition, 6) = EsNulo(.Recordset("Codigo").Value)
        KlexFacturas.TextMatrix(bPFactura.Recordset.AbsolutePosition, 7) = EsNulo(.Recordset("Nombre").Value)
        KlexFacturas.TextMatrix(bPFactura.Recordset.AbsolutePosition, 8) = EsNulo(.Recordset("Cuit").Value)
        KlexFacturas.TextMatrix(bPFactura.Recordset.AbsolutePosition, 9) = EsNulo(.Recordset("Neto").Value)
        
        KlexFacturas.TextMatrix(bPFactura.Recordset.AbsolutePosition, 10) = EsNulo(.Recordset("Iva105").Value)
        KlexFacturas.TextMatrix(bPFactura.Recordset.AbsolutePosition, 11) = EsNulo(.Recordset("Iva210").Value)
        KlexFacturas.TextMatrix(bPFactura.Recordset.AbsolutePosition, 12) = EsNulo(.Recordset("Iva270").Value)
        
        KlexFacturas.TextMatrix(bPFactura.Recordset.AbsolutePosition, 13) = EsNulo(.Recordset("Retenciones").Value)
        KlexFacturas.TextMatrix(bPFactura.Recordset.AbsolutePosition, 14) = EsNulo(.Recordset("Total").Value)
       
       ' ------------------------------------CITI ----------------------------------------------------------------
       
       
    
          
         
        ' ------------------------citi cbte ----------------------------------
        Dim vtipoDocu As String
        Dim vali As Double
        
        If .Recordset("TipoMovimiento") = "Nota C" And bPFactura.Recordset("Letra").Value = "B" Then
           vtipoDocu = getTipoDoc("NotaCB")
        Else
            vtipoDocu = getTipoDoc(.Recordset("TipoMovimiento").Value)
        End If
        
        ' vtipoDocu = getTipoDoc(.Recordset("TipoMovimiento").Value) ' modif 20-01-16
        
        vgline = vgline + 1
        
        vlb = ""
        
        If vgline > 1 Then vlb = vbCrLf
        
        
        
        vlb = vlb + FF(.Recordset("Fecha").Value) ' 11

        vlb = vlb + fn(Val(vtipoDocu), 3, "N") ' 22

        vlb = vlb + fn(.Recordset("PuntoDeVenta").Value, 5, "N")  ' 33

        vlb = vlb + fn(.Recordset("NroComprobante").Value, 20, "N") 'nro de comprobante 44

        vlb = vlb + fc(" ", 16) 'despacho de import 5
        
        vlb = vlb + fn(80, 2, "N") ' Código de documento del comprador 66
        
        
             
        If Not validarCUIT2(.Recordset("Cuit").Value) = 1 Then
                ponerenciti = False
                'vcont = vcont + 1
                'logform ("Error #" + Str(vcont) + " Cuit mal formado: " + .Recordset("Cuit").Value + " Nro. Comp. " + .Recordset("NroComprobante").Value + "  Persona: " + bFacturas.Recordset("Nombre").Value)
        End If
        
        
        vlb = vlb + fn(.Recordset("Cuit").Value, 20, "N") ' nro del doc ' 77'
        
        vlb = vlb + fc(.Recordset("Nombre").Value, 30) ' denominación   '88'
        
         vlb = vlb + fn(.Recordset("Total").Value, 15, "S") ' total '99'
       
         If Not (vtipoDocu = 11 Or vtipoDocu = 6) Then
            vlb = vlb + fn(.Recordset("NoGravado").Value, 15, "S") ' NoGravado '1010'
        Else
            vlb = vlb + fn(0, 15, "S") ' NoGravado '1010'
        End If
        
         vlb = vlb + fn(.Recordset("ImpExento").Value, 15, "S") ' ImpExento (1111)

         vlb = vlb + fn(.Recordset("Percepciones").Value, 15, "S") ' Percepciones (12)

         vlb = vlb + fn(0, 15, "S") ' Percepción a no categorizados '1313'

         vlb = vlb + fn(0, 15, "S") ' Importe de percepciones de Ingresos Brutos (1414)
        
         vlb = vlb + fn(0, 15, "S") ' Importe de percepciones impuestos Municipales (1515)

         vlb = vlb + fn(0, 15, "S") ' Importe impuesto interno (16)
        
         vlb = vlb + fc("PES", 3) ' Código de la moneda (17)

        vlb = vlb + fn(10000, 10, "S") ' Tipo de cambio (18)


        vali = 0
        
        If Abs(.Recordset("iva105").Value) > 0 Then vali = vali + 1
        If Abs(.Recordset("iva210").Value) > 0 Then vali = vali + 1
        If Abs(.Recordset("iva270").Value) > 0 Then vali = vali + 1

        vlb = vlb + fn(vali, 1, "N") ' Importe impuesto interno (19)
        
        If vali = 0 Then
             vlb = vlb + fc("N", 1) ' Importe impuesto interno (20)
        Else
             vlb = vlb + fc(" ", 1)
        End If
        
        vlb = vlb + fn(0, 15, "S") ' credito fiscal computable (21)

        vlb = vlb + fn(0, 15, "S") ' otros tributos (22)

        
        vlb = vlb + fn(0, 11, "N") ' cuit emisor/corredor - va con cero (23)

        vlb = vlb + fc(" ", 30) ' Denominacion emisor (24)
        
        vlb = vlb + fn(0, 15, "S") ' IVA comisión (25)

        If ponerenciti Then Print #Canal1, vlb;   ' graba la linea en el archivo REGINFO_CV_VENTAS_CBTE.TXT'
    

        '----------- Alicuota ---------------------------------
            
        vla = ""
       vgline2 = vgline2 + 1
        ' todo ojo que lo agregué
        'If vgline2 > 1 Then vla = vbCrLf
        
        
        vla = vla + fn(Val(vtipoDocu), 3, "N") ' 1
        vla = vla + fn(.Recordset("PuntoDeVenta").Value, 5, "N") ' 2
        vla = vla + fn(.Recordset("NroComprobante").Value, 20, "N") ' nro de comprobante '3'

         vla = vla + fn(80, 2, "N") ' Código de documento del vendedor 4

         vla = vla + fn(.Recordset("Cuit").Value, 20, "N") ' nro del doc '5'


        vla = vla + fn(.Recordset("Neto").Value, 15, "S")  ' Importe neto gravado '6

        vla2 = vla

        If Abs(.Recordset("iva105").Value) > 0 Then ' iva 10.5'
                vla2 = vla2 + fn(4, 4, "N") '7'
                vla2 = vla2 + fn(Abs(.Recordset("iva105").Value), 15, "S") ' 8'
             
                
                    If ponerenciti Then Print #Canal22, vla2   '\REGINFO_CV_VENTAS_ALICUOTAS.TXT'
                
                
                b1 = b1 + 1
        End If

        vla2 = vla


        If Abs(.Recordset("iva210").Value) > 0 Then  'iva 21'
            vla2 = vla2 + fn(5, 4, "N")  ' 7
            vla2 = vla2 + fn(Abs(.Recordset("iva210").Value), 15, "S") ' 8'
          
            
                If ponerenciti Then Print #Canal22, vla2   '\REGINFO_CV_VENTAS_ALICUOTAS.TXT'
            
       
            
            b1 = b1 + 1
        End If

        vla2 = vla

        If Abs(.Recordset("iva270").Value) > 0 Then  ' iva 27'
        
                vla2 = vla2 + fn(6, 4, "N")  ' 7
                vla2 = vla2 + fn(Abs(.Recordset("iva270").Value), 15, "S") ' 8'
                
                
         
                
                    If ponerenciti Then Print #Canal22, vla2   '\REGINFO_CV_VENTAS_ALICUOTAS.TXT'
                
             
                
                b1 = b1 + 1
        End If

       '-----------------------------------------------------------------------
       
       
    End With
    
    
    ' Wgestion
    If b1 = 0 Then vgline2 = vgline2 - 1

    
    

If Err Then GrabarLog "CopiarTemp", Err.Number & " " & Err.Description, Me.Name
End Sub
Function FormatoFactura(vcant As Integer, vNcomp As Long) As String
On Error Resume Next
    
    Dim i As Integer

    FormatoFactura = String(vcant - Len(Trim(Str(vNcomp))), "0") & vNcomp

If Err Then GrabarLog "FormatoFactura", Err.Number & " " & Err.Description, Me.Name
End Function
Private Function SeleccionarNumero(vCantidad As Integer) As Integer
On Error Resume Next
    
    Select Case vCantidad
    
        Case 1 To 24
            SeleccionarNumero = 1
        
        Case 25 To 48
            SeleccionarNumero = 2
            
        Case 49 To 72
            SeleccionarNumero = 3
        
        Case 73 To 96
            SeleccionarNumero = 4
    
        Case 97 To 120
            SeleccionarNumero = 5
        
        Case 121 To 144
            SeleccionarNumero = 6
        
        Case 145 To 168
            SeleccionarNumero = 7

        Case 169 To 192
            SeleccionarNumero = 8

        Case 193 To 216
            SeleccionarNumero = 9

        Case 217 To 240
            SeleccionarNumero = 10
    
    End Select
    
    Call temp(0, 0, 0, 0, Val(vNumeroInicial + SeleccionarNumero))

If Err Then GrabarLog "SeleccionarNumero", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub GenerarListado()
On Error Resume Next
    
    If Trim(dtpFecha(0).Text) = "" Or Trim(dtpFecha(1).Text) = "" Then
        MsgBox "Debe Seleccionar un Mes y/o un Año para poder ejecutar un listado!!!", vbInformation, "Mensaje ..."
        Exit Sub
    End If
    
    If Trim(dtpFecha(0).Text) > (dtpFecha(1).Text) Then
        MsgBox "Las Fechas que ha ingresado no presentan coherencia.", vbInformation, "Mensaje ..."
        Exit Sub
    End If
    
    If Not chkTotales.Value = xtpChecked Then
    
        With bPFactura
            .ConnectionString = pathDBMySQL
            .RecordSource = "SELECT * FROM pfactura WHERE (month(fecha) = '" & AjustarMes(Month(dtpFecha(0).Value)) & "' AND year(fecha) = '" & Year(dtpFecha(1).Value) & "') AND (tipo <> 'Documento' OR Tipo Is NULL) AND (TipoMovimiento <> 'RC')  ORDER BY Fecha ASC, tipo ASC, Ncomprobante ASC"
            .Refresh
            
            If Not .Recordset.EOF = True Then
                barra.Value = 0
                barra.Max = .Recordset.RecordCount
                FormatoGrilla (.Recordset.RecordCount)
            Else
                MsgBox "No existen movimientos de este mes!!!", vbExclamation, "Mensaje ..."
                FormatoGrilla (1)
                Exit Sub
            End If
            
        End With
        
        Call IniciarVariables
        
        Call CalcularTotales
    
        Call CargarGrillaTotales
    
        TabIva.Item(0).Selected = True
    Else
    
        Call CargarGrillaTotales
        TabIva.Item(1).Selected = True
    
    End If
    
    

If Err Then GrabarLog "GenerarListado", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CalcularTotales()

On Error Resume Next


Dim mesano As String

mesano = Format(Me.dtpFecha(0), "MM") + Format(Me.dtpFecha(0), "YYYY")



Canal1 = FreeFile
Open App.Path + "\CITI\REGINFO_CV_COMPRAS_CBTE_" + mesano + ".TXT" For Output As Canal1

Canal22 = FreeFile
Open App.Path + "\CITI\REGINFO_CV_COMPRAS_ALICUOTAS_" + mesano + ".TXT" For Output As Canal22

    
    vgline = 0
    vNGtotal = 0
    
    
    With bPFactura
        Do Until .Recordset.EOF = True
            Call CopiarTemp(0)
            .Recordset.MoveNext
            barra.Value = barra.Value + 1
        Loop

    End With
    
Close Canal1

Close Canal22

    'MsgBox vNGtotal
    
    KlexFacturas.TopRow = bTemp_Iva.Recordset.AbsolutePosition

If Err Then GrabarLog "CalcularTotales", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Reporte()
On Error Resume Next
    
    Unload Mantenimiento
    Load Mantenimiento

    MsgBox "Prepare la Impresora!!!!", vbInformation, "Mensaje ..."
    
    If Not LeerConfig(22) = "PorFecha" Then
        
        With Mantenimiento.rsIva
            If Not .State = 1 Then .Open
            .Close
            .Open
            
            '.Filter = "id_temp >= 0"
            '.Sort = "id_temp ASC"
        End With
       
        vI105FactA = CalcularTotal("Fact A", "I105")
        vI210FactA = CalcularTotal("Fact A", "I210")
        vI270FactA = CalcularTotal("Fact A", "I270")
        vNetoFactA = CalcularTotal("Fact A", "Neto")
        vTotalFactA = CalcularTotal("Fact A", "Total")
        
        vI105Monotributo = CalcularTotal("Fact B", "I105")
        vI210Monotributo = CalcularTotal("Fact B", "I210")
        vI270Monotributo = CalcularTotal("Fact B", "I270")
        vNetoMonotributo = CalcularTotal("Fact B", "Neto") - vI105Monotributo - vI210Monotributo - vI270Monotributo
        vTotalMonotributo = CalcularTotal("Fact B", "Total") - vI105Monotributo - vI210Monotributo - vI270Monotributo
    
        vI105NotaC = -Val(CalcularTotal("Nota C", "I105"))
        vI210NotaC = -Val(CalcularTotal("Nota C", "I210"))
        vI270NotaC = -Val(CalcularTotal("Nota C", "I270"))
        vNetoNotaC = -Val(CalcularTotal("Nota C", "Neto"))
        vTotalNotaC = -Val(CalcularTotal("Nota C", "Total"))
    
        vNetoNotaD = Val(CalcularTotal("Nota D", "Neto"))
        vI105NotaD = Val(CalcularTotal("Nota D", "I105"))
        vI210NotaD = Val(CalcularTotal("Nota D", "I210"))
        vI270NotaD = Val(CalcularTotal("Nota D", "I270"))
        vTotalNotaD = Val(CalcularTotal("Nota D", "Total"))
        
        With drIvaCompra.Sections("TituloEmpresa")
            .Controls("vmes").Caption = AjustarMes(Month(dtpFecha(0).Value))
            .Controls("vano").Caption = Year(dtpFecha(0).Value)
            
            .Controls("lblNombre").Caption = vDatosEmpresa.Nombre
            .Controls("lblDueno").Caption = "DE " & vDatosEmpresa.Responsable
            .Controls("lblCuit").Caption = "CUIT " & vDatosEmpresa.cuit
        End With
    
        With drIvaCompra.Sections("ReportFooter")
            .Controls("nfacturaa").Caption = Format(CalcularTotal("Fact A", "Neto"), "$########0.00")
            .Controls("nfacturab").Caption = Format(CalcularTotal("Fact B", "Neto"), "$########0.00")
            .Controls("nncredito").Caption = Format(CalcularTotal("Nota C", "Neto"), "$########0.00")
            .Controls("nndebito").Caption = Format(CalcularTotal("Nota D", "Neto"), "$########0.00")
            '.Controls("nfacturae").Caption = Format(CalcularTotal("Fact E", "Neto"), "$#######0.00")
    
            .Controls("ifacturaa105").Caption = Format(CalcularTotal("Fact A", "I105"), "$#######0.00")
            .Controls("ifacturab105").Caption = Format(CalcularTotal("Fact B", "I105"), "$########0.00")
            .Controls("incredito105").Caption = Format(CalcularTotal("Nota C", "I105"), "$########0.00")
            .Controls("indebito105").Caption = Format(CalcularTotal("Nota D", "I105"), "$########0.00")
            '.Controls("ifacturae105").Caption = Format(CalcularTotal("Fact E", "Neto"), "$#######0.00")
    
            .Controls("ifacturaa210").Caption = Format(CalcularTotal("Fact A", "I210"), "$#######0.00")
            .Controls("ifacturab210").Caption = Format(CalcularTotal("Fact B", "I210"), "$########0.00")
            .Controls("incredito210").Caption = Format(CalcularTotal("Nota C", "I210"), "$########0.00")
            .Controls("indebito210").Caption = Format(CalcularTotal("Nota D", "I210"), "$########0.00")
            '.Controls("ifacturae210").Caption = Format(CalcularTotal("Fact E", "Neto"), "$#######0.00")
            
            .Controls("ifacturaa270").Caption = Format(CalcularTotal("Fact A", "I270"), "$#######0.00")
            .Controls("ifacturab270").Caption = Format(CalcularTotal("Fact B", "I270"), "$########0.00")
            .Controls("incredito270").Caption = Format(CalcularTotal("Nota C", "I270"), "$########0.00")
            .Controls("indebito270").Caption = Format(CalcularTotal("Nota D", "I270"), "$########0.00")
            '.Controls("ifacturae270").Caption = Format(CalcularTotal("Fact E", "Iva270"), "$#######0.00")
    
            .Controls("tfacturaa").Caption = Format(CalcularTotal("Fact A", "Total"), "$#######0.00")
            .Controls("tfacturab").Caption = Format(CalcularTotal("Fact B", "Total"), "$########0.00")
            .Controls("tncredito").Caption = Format(CalcularTotal("Nota C", "Total"), "$########0.00")
            .Controls("tndebito").Caption = Format(CalcularTotal("Nota D", "Total"), "$########0.00")
            '.Controls("tfacturae").Caption = Format(CalcularTotal("Fact E", "Total"), "$#######0.00")
        
            .Controls("ntotal").Caption = Format(vNetoFactA + vNetoMonotributo + vNetoNotaC + vNetoNotaD, "$#######0.00")
            .Controls("I105total").Caption = Format(vI105FactA + vI105Monotributo + vI105NotaC + vI105NotaD, "$########0.00")
            .Controls("I210total").Caption = Format(vI210FactA + vI210Monotributo + vI210NotaC + vI210NotaD, "$########0.00")
            .Controls("I270total").Caption = Format(vI270FactA + vI270Monotributo + vI270NotaC + vI270NotaD, "$########0.00")
            .Controls("ttotal").Caption = Format(vTotalFactA + vTotalMonotributo + vTotalNotaC + vTotalNotaD, "$#######0.00")
        End With
        
        With drIvaCompra
            .Orientation = rptOrientLandscape
            .Refresh
            .Show
        End With
    Else
        With Mantenimiento.rsIvaPorFecha
            If Not .State = 0 Then .Close
        
            .Source = " SHAPE {SELECT Fecha FROM Temp_Iva GROUP BY Fecha} AS IvaPorFecha APPEND ({SELECT *, (t.`Iva105` + t.`Iva210` + t.`Iva270`) as ivas  FROM Temp_Iva t} AS IvaPorFechaDetalle RELATE 'Fecha' TO 'Fecha') AS IvaPorFechaDetalle"
        
            If Not .State = 1 Then .Open
            .Close
            .Open
        
            If .RecordCount = 0 Then
                MsgBox "No existen datos para visualizar", vbExclamation, "Mensaje ..."
                Exit Sub
            End If
        End With
    
        With drIvaCompraPorFecha.Sections("TituloEmpresa")
            
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
        
        With drIvaCompraPorFecha
            '.Orientation = rptOrientLandscape
            
            .Refresh
            .Show
        End With
    End If
    
    
If Err Then GrabarLog "Reporte", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub GuardarIva()
Dim i As Integer
On Error Resume Next

    If Not TraerDato("IvaCompraCerrado", "Periodo = '" & Year(dtpFecha(1).Value) & AjustarMes(Month(dtpFecha(0).Value)) & "'", "Periodo") = "" Then
        MsgBox "El Periodo ya se encuentra generado y Cerrado.", vbExclamation, "Mensaje ..."
        Exit Sub
    End If

    Call EjecutarScript("INSERT INTO IvaCompraCerrado (Periodo) VALUES ('" & Year(dtpFecha(1).Value) & AjustarMes(Month(dtpFecha(0).Value)) & "')")
    
    Call BorrarBase("IvaCompra WHERE (month(fecha) = '" & AjustarMes(Month(dtpFecha(0).Value)) & "' AND year(fecha) = '" & Year(dtpFecha(1).Value) & "') ", pathDBMySQL)
        
    With bIvaCompra
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM IvaCompra WHERE 1=2"
        .Refresh
    End With
        
    With bTemp_Iva
        .Refresh
        If Not .Recordset.EOF = True Then
            .Recordset.MoveFirst
            barra.Value = 0
            barra.Max = .Recordset.RecordCount
        Else
            MsgBox "NO tiene datos pre-cargados para guardar!!", vbExclamation, "Mensaje ..."
            Exit Sub
        End If
        Do Until .Recordset.EOF
            bIvaCompra.Recordset.AddNew
            For i = 1 To (.Recordset.Fields.Count - 2)
                If Not IsNull(.Recordset(i).Value) = True Then
                    bIvaCompra.Recordset(i).Value = .Recordset(i).Value
                End If
            Next
            barra.Value = barra.Value + 1
            bIvaCompra.Recordset.Update
            .Recordset.MoveNext
        Loop
        
        'Debug.Print (UltimaHoja(True, "NroHojaIC", vNumeroPagina))
        
    End With


If Err Then GrabarLog "GuardarIva", Err.Number & " " & Err.Description, Me.Name
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
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Function FormatoNC(ByRef vtipo As String, vNcomp As Long) As String
On Error Resume Next
    
    If (vtipo = "Liq.Prod A") Or (vtipo = "Liq.Prod B") Then
        FormatoNC = "0001" & "-" & String(8 - Len(Trim(vNcomp)), "0") & vNcomp
    Else
        FormatoNC = Trim(bPFactura.Recordset("Sucursal").Value) & "-" & String(8 - Len(Trim(vNcomp)), "0") & vNcomp
    End If

If Err Then GrabarLog "FormatoNC", Err.Number & " " & Err.Description, Me.Name
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
        .ColDisplayFormat(2) = "#,###,##0.00"
        
        .TextMatrix(0, 3) = "Iva10.5"
        .ColWidth(3) = 1000
        .ColDisplayFormat(3) = "#,###,##0.00"
        
        .TextMatrix(0, 4) = "Iva21"
        .ColWidth(4) = 1000
        .ColDisplayFormat(4) = "#,###,##0.00"
        
        .TextMatrix(0, 5) = "Iva27"
        .ColWidth(5) = 1000
        .ColDisplayFormat(5) = "#,###,##0.00"
        
        .TextMatrix(0, 6) = "Reten."
        .ColWidth(6) = 1000
        .ColDisplayFormat(6) = "#,###,##0.00"
        
        .TextMatrix(0, 7) = "Perc."
        .ColWidth(7) = 1000
        .ColDisplayFormat(7) = "#,###,##0.00"
        
        .TextMatrix(0, 8) = "NoGrab."
        .ColWidth(8) = 750
        .ColDisplayFormat(8) = "#,###,##0.00"
        
        .TextMatrix(0, 9) = "ITC"
        .ColWidth(9) = 750
        .ColDisplayFormat(9) = "#,###,##0.00"
                
        .TextMatrix(0, 10) = "Exento"
        .ColWidth(10) = 750
        .ColDisplayFormat(10) = "#,###,##0.00"
        
        .TextMatrix(0, 11) = "Total"
        .ColWidth(11) = 1000
        .ColDisplayFormat(11) = "#,###,##0.00"
        
        .Editable = False

        .EnterKeyBehaviour = klexEKNone


        .Row = 2
    End With
    
If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Caption
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
Private Sub IniciarVariables()
On Error Resume Next

    BorrarBase "Temp_Iva", pathDBMySQL
    BorrarBase "Temp", pathDBMySQL
    
    fdesde = strfechaMySQL(dtpFecha(0).Value)
    fhasta = strfechaMySQL(dtpFecha(1).Value)
    
    vNumeroPagina = 0
    vNumeroInicial = UltimaHoja(False, "NroHojaIC")

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
Private Function CalcularTotal(vtipo, vValor) As Double
On Error Resume Next

    If Not vtipo = "" Then
        CalcularTotal = Val(GenerarDato("SELECT Fecha, Tipo, Sum(PFactura.Subtotal) AS Neto, Sum(IvaFacturaCompra.Iva105) AS I105, Sum(IvaFacturaCompra.Iva210) AS I210, Sum(IvaFacturaCompra.Iva270) AS I270, SubTotal+Iva105+Iva210+Iva270 AS Total FROM PFactura INNER JOIN IvaFacturaCompra ON PFactura.Remito = IvaFacturaCompra.Remito GROUP BY PFactura.tipo, Month(PFactura.Fecha), Year(PFactura.Fecha) HAVING (PFactura.tipo = '" & vtipo & "') AND (Month(PFactura.Fecha) = '" & AjustarMes(Month(dtpFecha(0).Value)) & "') AND (Year(PFactura.Fecha) = '" & (Year(dtpFecha(0).Value)) & "')", vValor))
    Else
        CalcularTotal = Val(GenerarDato("SELECT Sum(PFactura.Subtotal) AS Neto, Sum(IvaFacturaCompra.Iva105) AS I105, Sum(IvaFacturaCompra.Iva210) AS I210, Sum(IvaFacturaCompra.Iva270) AS I270, SubTotal+Iva105+Iva210+Iva270 AS Total FROM PFactura INNER JOIN IvaFacturaCompra ON PFactura.Remito = IvaFacturaCompra.Remito GROUP BY Month(PFactura.Fecha), Year(PFactura.Fecha) HAVING (Month(PFactura.Fecha) = '" & AjustarMes(Month(dtpFecha(0).Value)) & "') AND (Year(PFactura.Fecha) = '" & (Year(dtpFecha(0).Value)) & "')", vValor))
    End If
    
If Err Then GrabarLog "CalcularTotal", Err.Number & " " & Err.Description, Me.Name
End Function
'Private Sub DgTemp_Iva_ButtonClick(ByVal ColIndex As Integer)
'On Error Resume Next
'
'    With bTemp_Iva
'        .Recordset("NroComprobante").Value = FormatoNC(.Recordset("Tipo").Value, Val(InputBox("Ingrese el numero de Documento!!")))
'        .Recordset.Update
'    End With
'
'If Err Then GrabarLog "DgTemp_Iva_ButtonClick", Err.Number & " " & Err.Description, Me.Name
'End Sub
Private Sub CargarGrillaTotales()
On Error Resume Next
    
    Dim rsTotales As New ADODB.Recordset, sqlTotales As String, i As Integer
    
   ' sqlTotales = "SELECT TipoMovimiento, SUM(SubTotal) as ImpBruto,SUM(Iva105) as ImpIva105,SUM(Iva210) as ImpIva210, SUM(Iva270) as ImpIva270, SUM(Retenciones) as ImpRetenciones,SUM(Percepciones) as ImpPercepciones,SUM(NoGravado) as ImpNoGravado, SUM(ITC) as ImpITC, SUM(ImpExento) as ImpExento,  SUM(Total) as ImpNeto FROM PFactura Fa INNER JOIN IvaFacturaCompra Iv ON Fa.Remito=IV.Remito WHERE Month(fecha) =  '" & AjustarMes(Month(dtpFecha(0).Value)) & "' And Year(fecha) = '" & Year(dtpFecha(1).Value) & "' GROUP BY TipoMovimiento;"
    sqlTotales = "SELECT TipoMovimiento, SUM(neto) as ImpNeto,SUM(Iva105) as ImpIva105,SUM(Iva210) as ImpIva210, SUM(Iva270) as ImpIva270, SUM(Retenciones) as ImpRetenciones,SUM(Percepciones) as ImpPercepciones,SUM(NoGravado) as ImpNoGravado, SUM(ITC) as ImpITC, SUM(ImpExento) as ImpExento,  SUM(Total) as ImpTotal FROM temp_IVA WHERE Month(fecha) =  '" & AjustarMes(Month(dtpFecha(0).Value)) & "' And Year(fecha) = '" & Year(dtpFecha(1).Value) & "' GROUP BY TipoMovimiento;"
    
    
    'Pase_Excel (sqlTotales)
    
    With rsTotales
        .CursorLocation = adUseClient
        
        Call .Open(sqlTotales, ConnDDBB, adOpenStatic, adLockBatchOptimistic)
        
        If Not .EOF = True Then .MoveFirst
        
        
        FormatoGrillaTotales (.RecordCount + 1)
        
        i = 0
        Do Until .EOF = True
            klexTotales.TextMatrix(.AbsolutePosition, 1) = EsNulo(.Fields("TipoMovimiento").Value)
            klexTotales.TextMatrix(.AbsolutePosition, 2) = EsNulo(.Fields("ImpNeto").Value)
            klexTotales.TextMatrix(.AbsolutePosition, 3) = EsNulo(.Fields("ImpIva105").Value)
            klexTotales.TextMatrix(.AbsolutePosition, 4) = EsNulo(.Fields("ImpIva210").Value)
            klexTotales.TextMatrix(.AbsolutePosition, 5) = EsNulo(.Fields("ImpIva270").Value)
            klexTotales.TextMatrix(.AbsolutePosition, 6) = EsNulo(.Fields("ImpRetenciones").Value)
            klexTotales.TextMatrix(.AbsolutePosition, 7) = EsNulo(.Fields("ImpPercepciones").Value)
            klexTotales.TextMatrix(.AbsolutePosition, 8) = EsNulo(.Fields("ImpNoGravado").Value)
            klexTotales.TextMatrix(.AbsolutePosition, 9) = EsNulo(.Fields("ImpITC").Value)
            klexTotales.TextMatrix(.AbsolutePosition, 10) = EsNulo(.Fields("ImpExento").Value)
            klexTotales.TextMatrix(.AbsolutePosition, 11) = EsNulo(.Fields("ImpTotal").Value)
            
            .MoveNext
            i = i + 1
        Loop
        
        '--------------------------------------------------------------------------------------------------------------------------------------------------------
        
        'Cargo Los Totales dentro de la misma Grilla
        .Close
        
        sqlTotales = "SELECT SUM(neto) as ImpNeto,SUM(Iva105) as ImpIva105,SUM(Iva210) as ImpIva210, SUM(Iva270) as ImpIva270, SUM(Retenciones) as ImpRetenciones,SUM(Percepciones) as ImpPercepciones,SUM(NoGravado) as ImpNoGravado, SUM(ITC) as ImpITC, SUM(ImpExento) as ImpExento,  SUM(Total) as ImpTotal FROM temp_IVA WHERE Month(fecha) =  '" & AjustarMes(Month(dtpFecha(0).Value)) & "' And Year(fecha) = '" & Year(dtpFecha(1).Value) & "'"
        Call .Open(sqlTotales, ConnDDBB, adOpenStatic, adLockBatchOptimistic)
    
        klexTotales.Row = i + 1

        klexTotales.TextMatrix(i + 1, 1) = "Totales :"
        klexTotales.TextMatrix(i + 1, 2) = EsNulo(.Fields("ImpNeto").Value)
        klexTotales.TextMatrix(i + 1, 3) = EsNulo(.Fields("ImpIva105").Value)
        klexTotales.TextMatrix(i + 1, 4) = EsNulo(.Fields("ImpIva210").Value)
        klexTotales.TextMatrix(i + 1, 5) = EsNulo(.Fields("ImpIva270").Value)
        klexTotales.TextMatrix(i + 1, 6) = EsNulo(.Fields("ImpRetenciones").Value)
        klexTotales.TextMatrix(i + 1, 7) = EsNulo(.Fields("ImpPercepciones").Value)
        klexTotales.TextMatrix(i + 1, 8) = EsNulo(.Fields("ImpNoGravado").Value)
        klexTotales.TextMatrix(i + 1, 9) = EsNulo(.Fields("ImpITC").Value)
        klexTotales.TextMatrix(i + 1, 10) = EsNulo(.Fields("ImpExento").Value)
        klexTotales.TextMatrix(i + 1, 11) = EsNulo(.Fields("ImpTotal").Value)

    End With
    
    sqlTotales = ""
    
    If rsTotales.State = 1 Then
        rsTotales.Close
        Set rsTotales = Nothing
    End If
    
If Err Then GrabarLog "CargarGrillaTotales", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub PbAcciones_Click(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 0
            GenerarListado
            
        
        Case 1
            GuardarIva
            
        Case 2
            Reporte
        
        Case 3
            Unload Me
    
    End Select
If Err Then GrabarLog "PbAcciones_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub PushButton1_Click()
On Error Resume Next
    
  Call grillaToExcel2(Me.KlexFacturas)

If Err Then Exit Sub
End Sub

Private Sub PushButton2_Click()
Call ImprimirFlex(Me.klexTotales, "Totales IVA Compra", "")
End Sub
