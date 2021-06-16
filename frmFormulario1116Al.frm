VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "CODEJO~3.OCX"
Begin VB.Form frmFormulario1116Al 
   Caption         =   "Formulario 1116. Compra / Venta  - Liquidación"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12465
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   12465
   Begin VB.PictureBox PicInferior 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   12465
      TabIndex        =   14
      Top             =   4560
      Width           =   12465
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   0
         Left            =   10080
         TabIndex        =   15
         Top             =   90
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Grabar"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   1
         Left            =   11190
         TabIndex        =   16
         Top             =   75
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Cerrar"
         Appearance      =   6
      End
      Begin VB.Label lblWGestion 
         BackStyle       =   0  'Transparent
         Caption         =   "WGESTION 2010"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
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
         TabIndex        =   18
         Top             =   170
         Width           =   1770
      End
      Begin VB.Label lblWGestion 
         BackStyle       =   0  'Transparent
         Caption         =   "WGESTION 2010"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
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
         TabIndex        =   17
         Top             =   150
         Width           =   1770
      End
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   3885
      Left            =   60
      TabIndex        =   3
      Top             =   660
      Width           =   12435
      _Version        =   851968
      _ExtentX        =   21934
      _ExtentY        =   6853
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AllowReorder    =   -1  'True
      Appearance      =   3
      Color           =   4
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      ItemCount       =   6
      Item(0).Caption =   "Actuo Corredor:"
      Item(0).ControlCount=   4
      Item(0).Control(0)=   "OptSI"
      Item(0).Control(1)=   "OptNO"
      Item(0).Control(2)=   "Frame1"
      Item(0).Control(3)=   "lblIndiqueSi"
      Item(1).Caption =   "Condiciones de la Operación"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "Frame2"
      Item(2).Caption =   "Mercadería Entregada"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "Frame3"
      Item(3).Caption =   "Deducciones"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "Frame4"
      Item(4).Caption =   "Retenciones"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "Frame5"
      Item(5).Caption =   "Total Retenciones"
      Item(5).ControlCount=   1
      Item(5).Control(0)=   "Frame6"
      Begin VB.Frame Frame6 
         Height          =   2985
         Left            =   -69910
         TabIndex        =   71
         Top             =   750
         Visible         =   0   'False
         Width           =   12195
         Begin XtremeSuiteControls.FlatEdit vLocalidad 
            Height          =   285
            Left            =   4680
            TabIndex        =   72
            Top             =   300
            Width           =   7095
            _Version        =   851968
            _ExtentX        =   12515
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vImporteNetoAPagar 
            Height          =   285
            Left            =   4680
            TabIndex        =   73
            Top             =   630
            Width           =   7095
            _Version        =   851968
            _ExtentX        =   12515
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vPagoIVARES1394 
            Height          =   285
            Left            =   4680
            TabIndex        =   74
            Top             =   960
            Width           =   7095
            _Version        =   851968
            _ExtentX        =   12515
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vFecha 
            Height          =   285
            Index           =   1
            Left            =   4680
            TabIndex        =   75
            Top             =   1290
            Width           =   7095
            _Version        =   851968
            _ExtentX        =   12515
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vPagoSCondiciones 
            Height          =   285
            Left            =   4680
            TabIndex        =   76
            Top             =   1620
            Width           =   7095
            _Version        =   851968
            _ExtentX        =   12515
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vNombreApellidoFirmaDNIVendedor 
            Height          =   285
            Index           =   2
            Left            =   4680
            TabIndex        =   82
            Top             =   1950
            Width           =   7095
            _Version        =   851968
            _ExtentX        =   12515
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vNombreApellidoFirmaDNIComprador 
            Height          =   285
            Left            =   4680
            TabIndex        =   83
            Top             =   2280
            Width           =   7095
            _Version        =   851968
            _ExtentX        =   12515
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin VB.Label lblFechaafsafs 
            Alignment       =   1  'Right Justify
            Caption         =   "> Nombre y Apellido, Firma y N° Documento del Vendedor:"
            Height          =   255
            Left            =   360
            TabIndex        =   85
            Top             =   1950
            Width           =   4185
         End
         Begin VB.Label lblPagoS45646 
            Alignment       =   1  'Right Justify
            Caption         =   "> Nombre y Apellido, Firma y N° Documento del Comprador:"
            Height          =   255
            Left            =   240
            TabIndex        =   84
            Top             =   2280
            Width           =   4305
         End
         Begin VB.Label lblLocalidad 
            Alignment       =   1  'Right Justify
            Caption         =   "> Localidad:"
            Height          =   255
            Left            =   2760
            TabIndex        =   81
            Top             =   300
            Width           =   1815
         End
         Begin VB.Label lblImporteNeto 
            Alignment       =   1  'Right Justify
            Caption         =   "> Importe Neto a Pagar:"
            Height          =   255
            Left            =   2760
            TabIndex        =   80
            Top             =   630
            Width           =   1785
         End
         Begin VB.Label lblPagoIVA 
            Alignment       =   1  'Right Justify
            Caption         =   "> Pago IVA RES. 1394:"
            Height          =   255
            Left            =   2760
            TabIndex        =   79
            Top             =   960
            Width           =   1785
         End
         Begin VB.Label lblFecha2 
            Alignment       =   1  'Right Justify
            Caption         =   "> Fecha:"
            Height          =   255
            Left            =   2760
            TabIndex        =   78
            Top             =   1290
            Width           =   1785
         End
         Begin VB.Label lblPagoS 
            Alignment       =   1  'Right Justify
            Caption         =   "> Pago s/condiciones:"
            Height          =   255
            Left            =   2760
            TabIndex        =   77
            Top             =   1620
            Width           =   1785
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2985
         Left            =   -69910
         TabIndex        =   62
         Top             =   750
         Visible         =   0   'False
         Width           =   12195
         Begin XtremeSuiteControls.FlatEdit vIngresosBrutos 
            Height          =   285
            Left            =   2430
            TabIndex        =   63
            Top             =   300
            Width           =   8145
            _Version        =   851968
            _ExtentX        =   14367
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vSellosYRegistro 
            Height          =   285
            Left            =   2430
            TabIndex        =   64
            Top             =   630
            Width           =   8145
            _Version        =   851968
            _ExtentX        =   14367
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vRetencionIVA 
            Height          =   285
            Left            =   2430
            TabIndex        =   65
            Top             =   960
            Width           =   8145
            _Version        =   851968
            _ExtentX        =   14367
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vRetencionGanancias 
            Height          =   285
            Left            =   2430
            TabIndex        =   66
            Top             =   1290
            Width           =   8145
            _Version        =   851968
            _ExtentX        =   14367
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin VB.Label lblIngresosBrutos 
            Alignment       =   1  'Right Justify
            Caption         =   "> Ingresos Brutos:"
            Height          =   255
            Left            =   270
            TabIndex        =   70
            Top             =   300
            Width           =   1815
         End
         Begin VB.Label lblSellosY 
            Alignment       =   1  'Right Justify
            Caption         =   "> Sellos y Registro:"
            Height          =   255
            Left            =   300
            TabIndex        =   69
            Top             =   630
            Width           =   1785
         End
         Begin VB.Label lblRetenciónIVA 
            Alignment       =   1  'Right Justify
            Caption         =   "> Retención IVA:"
            Height          =   255
            Left            =   300
            TabIndex        =   68
            Top             =   960
            Width           =   1785
         End
         Begin VB.Label lblRetenciónGanancias 
            Alignment       =   1  'Right Justify
            Caption         =   "> Retención Ganancias:"
            Height          =   255
            Left            =   300
            TabIndex        =   67
            Top             =   1290
            Width           =   1785
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2985
         Left            =   -69910
         TabIndex        =   55
         Top             =   750
         Visible         =   0   'False
         Width           =   12195
         Begin XtremeSuiteControls.FlatEdit vAlmacenaje 
            Height          =   285
            Left            =   2430
            TabIndex        =   56
            Top             =   300
            Width           =   6675
            _Version        =   851968
            _ExtentX        =   11774
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vOtros 
            Height          =   285
            Left            =   2430
            TabIndex        =   57
            Top             =   630
            Width           =   6675
            _Version        =   851968
            _ExtentX        =   11774
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vIVASDeducciones 
            Height          =   285
            Left            =   2430
            TabIndex        =   58
            Top             =   960
            Width           =   6675
            _Version        =   851968
            _ExtentX        =   11774
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin VB.Label lblAlmacenaje 
            Alignment       =   1  'Right Justify
            Caption         =   "> Almacenaje:"
            Height          =   255
            Left            =   270
            TabIndex        =   61
            Top             =   300
            Width           =   1815
         End
         Begin VB.Label lblOtros 
            Alignment       =   1  'Right Justify
            Caption         =   "> Otros:"
            Height          =   255
            Left            =   300
            TabIndex        =   60
            Top             =   630
            Width           =   1785
         End
         Begin VB.Label lblIVASD 
            Alignment       =   1  'Right Justify
            Caption         =   "> IVA s/deducciones:"
            Height          =   255
            Left            =   300
            TabIndex        =   59
            Top             =   960
            Width           =   1785
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3345
         Left            =   -69910
         TabIndex        =   36
         Top             =   390
         Visible         =   0   'False
         Width           =   12195
         Begin XtremeSuiteControls.FlatEdit vNrosDeCertificadosDepositoIntransferible 
            Height          =   285
            Left            =   3750
            TabIndex        =   37
            Top             =   300
            Width           =   8145
            _Version        =   851968
            _ExtentX        =   14367
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vGrado 
            Height          =   285
            Index           =   1
            Left            =   3750
            TabIndex        =   38
            Top             =   630
            Width           =   8145
            _Version        =   851968
            _ExtentX        =   14367
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vContenidoProteico 
            Height          =   285
            Left            =   3750
            TabIndex        =   39
            Top             =   960
            Width           =   8145
            _Version        =   851968
            _ExtentX        =   14367
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vFactor 
            Height          =   285
            Left            =   3750
            TabIndex        =   40
            Top             =   1290
            Width           =   8145
            _Version        =   851968
            _ExtentX        =   14367
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vFormaDePago 
            Height          =   285
            Left            =   3750
            TabIndex        =   41
            Top             =   1620
            Width           =   8145
            _Version        =   851968
            _ExtentX        =   14367
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vPrecioOperacion 
            Height          =   285
            Left            =   3750
            TabIndex        =   47
            Top             =   1950
            Width           =   8145
            _Version        =   851968
            _ExtentX        =   14367
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vPesoNeto 
            Height          =   285
            Left            =   3750
            TabIndex        =   49
            Top             =   2280
            Width           =   8145
            _Version        =   851968
            _ExtentX        =   14367
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vImporteBruto 
            Height          =   285
            Left            =   3750
            TabIndex        =   51
            Top             =   2610
            Width           =   8145
            _Version        =   851968
            _ExtentX        =   14367
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vIVASImporteBruto 
            Height          =   285
            Left            =   3750
            TabIndex        =   53
            Top             =   2940
            Width           =   8145
            _Version        =   851968
            _ExtentX        =   14367
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin VB.Label lblIVAS 
            Alignment       =   1  'Right Justify
            Caption         =   "> IVA s/Importe Bruto:"
            Height          =   255
            Left            =   1710
            TabIndex        =   54
            Top             =   2940
            Width           =   1785
         End
         Begin VB.Label lblImporteBruto 
            Alignment       =   1  'Right Justify
            Caption         =   "> Importe Bruto:"
            Height          =   255
            Left            =   1710
            TabIndex        =   52
            Top             =   2610
            Width           =   1785
         End
         Begin VB.Label lblPesoNeto 
            Alignment       =   1  'Right Justify
            Caption         =   "> Peso Neto:"
            Height          =   255
            Left            =   1710
            TabIndex        =   50
            Top             =   2280
            Width           =   1785
         End
         Begin VB.Label lblPrecioOperación 
            Alignment       =   1  'Right Justify
            Caption         =   "> Precio Operación:"
            Height          =   255
            Left            =   1710
            TabIndex        =   48
            Top             =   1950
            Width           =   1785
         End
         Begin VB.Label lblNrosDe 
            Alignment       =   1  'Right Justify
            Caption         =   "> Nros. De certificados Deposito Intransferible :"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   300
            Width           =   3375
         End
         Begin VB.Label lblGrado2 
            Alignment       =   1  'Right Justify
            Caption         =   "> Grado:"
            Height          =   255
            Left            =   1710
            TabIndex        =   45
            Top             =   630
            Width           =   1785
         End
         Begin VB.Label lblContenidoProteico 
            Alignment       =   1  'Right Justify
            Caption         =   "> Contenido Proteico:"
            Height          =   255
            Left            =   1710
            TabIndex        =   44
            Top             =   960
            Width           =   1785
         End
         Begin VB.Label lblFactor 
            Alignment       =   1  'Right Justify
            Caption         =   "> Factor:"
            Height          =   255
            Left            =   1710
            TabIndex        =   43
            Top             =   1290
            Width           =   1785
         End
         Begin VB.Label lblFormaDe 
            Alignment       =   1  'Right Justify
            Caption         =   "> Forma de Pago:"
            Height          =   255
            Left            =   1710
            TabIndex        =   42
            Top             =   1620
            Width           =   1785
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2985
         Left            =   -69910
         TabIndex        =   25
         Top             =   750
         Visible         =   0   'False
         Width           =   12195
         Begin XtremeSuiteControls.FlatEdit vFecha 
            Height          =   285
            Index           =   0
            Left            =   3000
            TabIndex        =   26
            Top             =   300
            Width           =   6555
            _Version        =   851968
            _ExtentX        =   11562
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vPrecioReferencia 
            Height          =   285
            Left            =   3000
            TabIndex        =   27
            Top             =   630
            Width           =   6555
            _Version        =   851968
            _ExtentX        =   11562
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vGrado 
            Height          =   285
            Index           =   0
            Left            =   3000
            TabIndex        =   28
            Top             =   960
            Width           =   6555
            _Version        =   851968
            _ExtentX        =   11562
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vFleteC100Kgrs 
            Height          =   285
            Left            =   3000
            TabIndex        =   29
            Top             =   1290
            Width           =   6555
            _Version        =   851968
            _ExtentX        =   11562
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vPuertoOLugarDeReferencia 
            Height          =   285
            Left            =   3000
            TabIndex        =   30
            Top             =   1620
            Width           =   6555
            _Version        =   851968
            _ExtentX        =   11562
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin VB.Label lblFecha 
            Alignment       =   1  'Right Justify
            Caption         =   "> Fecha:"
            Height          =   255
            Left            =   895
            TabIndex        =   35
            Top             =   300
            Width           =   1815
         End
         Begin VB.Label lblPrecioReferencia 
            Alignment       =   1  'Right Justify
            Caption         =   "> Precio Referencia:"
            Height          =   255
            Left            =   895
            TabIndex        =   34
            Top             =   630
            Width           =   1785
         End
         Begin VB.Label lblGrado 
            Alignment       =   1  'Right Justify
            Caption         =   "> Grado:"
            Height          =   255
            Left            =   895
            TabIndex        =   33
            Top             =   960
            Width           =   1785
         End
         Begin VB.Label lblFleteC 
            Alignment       =   1  'Right Justify
            Caption         =   "> Flete c/100 Kgrs:"
            Height          =   255
            Left            =   895
            TabIndex        =   32
            Top             =   1290
            Width           =   1785
         End
         Begin VB.Label lblPuertoO 
            Alignment       =   1  'Right Justify
            Caption         =   "> Puerto o Lugar de Referencia:"
            Height          =   255
            Left            =   300
            TabIndex        =   31
            Top             =   1620
            Width           =   2385
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2985
         Left            =   90
         TabIndex        =   6
         Top             =   750
         Width           =   12195
         Begin XtremeSuiteControls.FlatEdit vNroRegistroSAPyA 
            Height          =   285
            Left            =   2430
            TabIndex        =   9
            Top             =   300
            Width           =   5805
            _Version        =   851968
            _ExtentX        =   10239
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vCuit 
            Height          =   285
            Left            =   2430
            TabIndex        =   11
            Top             =   630
            Width           =   5805
            _Version        =   851968
            _ExtentX        =   10239
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vComisionCorredor 
            Height          =   285
            Left            =   2430
            TabIndex        =   19
            Top             =   960
            Width           =   5805
            _Version        =   851968
            _ExtentX        =   10239
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vRazonSocial 
            Height          =   285
            Left            =   2430
            TabIndex        =   21
            Top             =   1290
            Width           =   5805
            _Version        =   851968
            _ExtentX        =   10239
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vDomicilio 
            Height          =   285
            Left            =   2430
            TabIndex        =   23
            Top             =   1620
            Width           =   5805
            _Version        =   851968
            _ExtentX        =   10239
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin VB.Label lblDomicilio 
            Alignment       =   1  'Right Justify
            Caption         =   "> Domicilio:"
            Height          =   255
            Left            =   300
            TabIndex        =   24
            Top             =   1620
            Width           =   1785
         End
         Begin VB.Label lblRazónSocial 
            Alignment       =   1  'Right Justify
            Caption         =   "> Razón Social:"
            Height          =   255
            Left            =   300
            TabIndex        =   22
            Top             =   1290
            Width           =   1785
         End
         Begin VB.Label lblComisiónCorredor 
            Alignment       =   1  'Right Justify
            Caption         =   "> Comisión Corredor:"
            Height          =   255
            Left            =   300
            TabIndex        =   20
            Top             =   960
            Width           =   1785
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "> Cuit:"
            Height          =   255
            Left            =   300
            TabIndex        =   10
            Top             =   630
            Width           =   1785
         End
         Begin VB.Label lblNroRegistro 
            Alignment       =   1  'Right Justify
            Caption         =   "> Nro. Registro SAPyA:"
            Height          =   255
            Left            =   270
            TabIndex        =   8
            Top             =   300
            Width           =   1815
         End
      End
      Begin VB.OptionButton OptNO 
         Caption         =   "NO"
         Height          =   195
         Left            =   3210
         TabIndex        =   5
         Top             =   510
         Width           =   675
      End
      Begin VB.OptionButton OptSI 
         Caption         =   "SI"
         Height          =   195
         Left            =   2580
         TabIndex        =   4
         Top             =   510
         Width           =   675
      End
      Begin VB.Label lblIndiqueSi 
         Caption         =   "Indique si actuo un corredor: "
         Height          =   285
         Left            =   150
         TabIndex        =   7
         Top             =   510
         Width           =   2295
      End
   End
   Begin XtremeSuiteControls.GroupBox GBTipoComprobante 
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12480
      _Version        =   851968
      _ExtentX        =   22013
      _ExtentY        =   1138
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.ComboBox cboLetra 
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   210
         Width           =   1485
         _Version        =   851968
         _ExtentX        =   2619
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         DropDownItemCount=   5
      End
      Begin XtremeSuiteControls.FlatEdit txtNroComprobante 
         Height          =   315
         Left            =   5520
         TabIndex        =   2
         Top             =   240
         Width           =   4305
         _Version        =   851968
         _ExtentX        =   7594
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
      End
      Begin VB.Label lblNroDel 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. del Comprobante:"
         Height          =   285
         Left            =   3750
         TabIndex        =   13
         Top             =   270
         Width           =   1905
      End
      Begin VB.Label lblLetraDel 
         BackStyle       =   0  'Transparent
         Caption         =   "Letra del Comprobante:"
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1905
      End
   End
End
Attribute VB_Name = "frmFormulario1116Al"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

