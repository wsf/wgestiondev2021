VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmFormulario1116 
   Caption         =   "Formulario 1116. Compra / Venta  - Liquidación"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13335
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4620
   ScaleWidth      =   13335
   Begin MSComctlLib.TreeView TreeView3 
      Height          =   3915
      Left            =   30
      TabIndex        =   13
      Top             =   60
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   6906
      _Version        =   393217
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.PictureBox PicInferior 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   30
      Picture         =   "frmFormulario1116.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   17985
      TabIndex        =   8
      Top             =   4020
      Width           =   17985
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   375
         Index           =   0
         Left            =   10800
         TabIndex        =   9
         Top             =   90
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Grabar"
         Appearance      =   6
         Picture         =   "frmFormulario1116.frx":50B3
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   375
         Index           =   1
         Left            =   11910
         TabIndex        =   10
         Top             =   75
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cerrar"
         Appearance      =   6
         Picture         =   "frmFormulario1116.frx":54BA
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   150
         Width           =   1770
      End
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   3315
      Left            =   3270
      TabIndex        =   2
      Top             =   660
      Width           =   14775
      _Version        =   851968
      _ExtentX        =   26061
      _ExtentY        =   5847
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
      SelectedItem    =   1
      Item(0).Caption =   "Actuo Corredor:"
      Item(0).ControlCount=   4
      Item(0).Control(0)=   "OptSI"
      Item(0).Control(1)=   "OptNO"
      Item(0).Control(2)=   "Frame1"
      Item(0).Control(3)=   "lblIndiqueSi"
      Item(1).Caption =   "Condiciones de la Operación"
      Item(1).ControlCount=   11
      Item(1).Control(0)=   "vFecha(0)"
      Item(1).Control(1)=   "vPrecioReferencia"
      Item(1).Control(2)=   "vGrado(0)"
      Item(1).Control(3)=   "vFleteC100Kgrs"
      Item(1).Control(4)=   "vPuertoOLugarDeReferencia"
      Item(1).Control(5)=   "lblFecha"
      Item(1).Control(6)=   "lblPrecioReferencia"
      Item(1).Control(7)=   "lblGrado"
      Item(1).Control(8)=   "lblFleteC"
      Item(1).Control(9)=   "lblPuertoO"
      Item(1).Control(10)=   "Pus(0)"
      Item(2).Caption =   "Mercadería Entregada"
      Item(2).ControlCount=   18
      Item(2).Control(0)=   "vNrosDeCertificadosDepositoIntransferible"
      Item(2).Control(1)=   "vGrado(1)"
      Item(2).Control(2)=   "vContenidoProteico"
      Item(2).Control(3)=   "vFactor"
      Item(2).Control(4)=   "vFormaDePago"
      Item(2).Control(5)=   "vPrecioOperacion"
      Item(2).Control(6)=   "vPesoNeto"
      Item(2).Control(7)=   "vImporteBruto"
      Item(2).Control(8)=   "vIVASImporteBruto"
      Item(2).Control(9)=   "lblIVAS"
      Item(2).Control(10)=   "lblImporteBruto"
      Item(2).Control(11)=   "lblPesoNeto"
      Item(2).Control(12)=   "lblPrecioOperación"
      Item(2).Control(13)=   "lblNrosDe"
      Item(2).Control(14)=   "lblGrado2"
      Item(2).Control(15)=   "lblContenidoProteico"
      Item(2).Control(16)=   "lblFactor"
      Item(2).Control(17)=   "lblFormaDe"
      Item(3).Caption =   "Deducciones"
      Item(3).ControlCount=   6
      Item(3).Control(0)=   "vAlmacenaje"
      Item(3).Control(1)=   "vOtros"
      Item(3).Control(2)=   "vIVASDeducciones"
      Item(3).Control(3)=   "lblAlmacenaje"
      Item(3).Control(4)=   "lblOtros"
      Item(3).Control(5)=   "lblIVASD"
      Item(4).Caption =   "Retenciones"
      Item(4).ControlCount=   8
      Item(4).Control(0)=   "vIngresosBrutos"
      Item(4).Control(1)=   "vSellosYRegistro"
      Item(4).Control(2)=   "vRetencionIVA"
      Item(4).Control(3)=   "vRetencionGanancias"
      Item(4).Control(4)=   "lblIngresosBrutos"
      Item(4).Control(5)=   "lblSellosY"
      Item(4).Control(6)=   "lblRetenciónIVA"
      Item(4).Control(7)=   "lblRetenciónGanancias"
      Item(5).Caption =   "Total Retenciones"
      Item(5).ControlCount=   15
      Item(5).Control(0)=   "vLocalidad"
      Item(5).Control(1)=   "vImporteNetoAPagar"
      Item(5).Control(2)=   "vPagoIVARES1394"
      Item(5).Control(3)=   "vFecha(1)"
      Item(5).Control(4)=   "vPagoSCondiciones"
      Item(5).Control(5)=   "vNombreApellidoFirmaDNIVendedor(2)"
      Item(5).Control(6)=   "vNombreApellidoFirmaDNIComprador"
      Item(5).Control(7)=   "lblFechaafsafs"
      Item(5).Control(8)=   "lblPagoS45646"
      Item(5).Control(9)=   "lblLocalidad"
      Item(5).Control(10)=   "lblImporteNeto"
      Item(5).Control(11)=   "lblPagoIVA"
      Item(5).Control(12)=   "lblFecha2"
      Item(5).Control(13)=   "lblPagoS"
      Item(5).Control(14)=   "Pus(1)"
      Begin XtremeSuiteControls.PushButton Pus 
         Height          =   255
         Index           =   1
         Left            =   -60520
         TabIndex        =   81
         Top             =   660
         Visible         =   0   'False
         Width           =   435
         _Version        =   851968
         _ExtentX        =   767
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton Pus 
         Height          =   255
         Index           =   0
         Left            =   8460
         TabIndex        =   80
         Top             =   2250
         Width           =   885
         _Version        =   851968
         _ExtentX        =   1561
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Frame Frame1 
         Height          =   2595
         Left            =   -69940
         TabIndex        =   5
         Top             =   630
         Visible         =   0   'False
         Width           =   9555
         Begin XtremeSuiteControls.FlatEdit vNroRegistroSAPyA 
            Height          =   285
            Left            =   3090
            TabIndex        =   14
            Top             =   480
            Width           =   5805
            _Version        =   851968
            _ExtentX        =   10239
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vCuit 
            Height          =   285
            Left            =   3090
            TabIndex        =   15
            Top             =   810
            Width           =   5805
            _Version        =   851968
            _ExtentX        =   10239
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vComisionCorredor 
            Height          =   285
            Left            =   3090
            TabIndex        =   16
            Top             =   1140
            Width           =   5805
            _Version        =   851968
            _ExtentX        =   10239
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vRazonSocial 
            Height          =   285
            Left            =   3090
            TabIndex        =   17
            Top             =   1470
            Width           =   5805
            _Version        =   851968
            _ExtentX        =   10239
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vDomicilio 
            Height          =   285
            Left            =   3090
            TabIndex        =   18
            Top             =   1800
            Width           =   5805
            _Version        =   851968
            _ExtentX        =   10239
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin VB.Label lblNroRegistro 
            Alignment       =   1  'Right Justify
            Caption         =   "> Nro. Registro SAPyA:"
            Height          =   255
            Left            =   930
            TabIndex        =   23
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "> Cuit:"
            Height          =   255
            Left            =   960
            TabIndex        =   22
            Top             =   810
            Width           =   1785
         End
         Begin VB.Label lblComisiónCorredor 
            Alignment       =   1  'Right Justify
            Caption         =   "> Comisión Corredor:"
            Height          =   255
            Left            =   960
            TabIndex        =   21
            Top             =   1140
            Width           =   1785
         End
         Begin VB.Label lblRazónSocial 
            Alignment       =   1  'Right Justify
            Caption         =   "> Razón Social:"
            Height          =   255
            Left            =   960
            TabIndex        =   20
            Top             =   1470
            Width           =   1785
         End
         Begin VB.Label lblDomicilio 
            Alignment       =   1  'Right Justify
            Caption         =   "> Domicilio:"
            Height          =   255
            Left            =   960
            TabIndex        =   19
            Top             =   1800
            Width           =   1785
         End
      End
      Begin VB.OptionButton OptNO 
         Caption         =   "NO"
         Height          =   195
         Left            =   -66790
         TabIndex        =   4
         Top             =   450
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.OptionButton OptSI 
         Caption         =   "SI"
         Height          =   195
         Left            =   -67540
         TabIndex        =   3
         Top             =   450
         Visible         =   0   'False
         Width           =   675
      End
      Begin XtremeSuiteControls.FlatEdit vFecha 
         Height          =   285
         Index           =   0
         Left            =   2760
         TabIndex        =   24
         Top             =   900
         Width           =   6555
         _Version        =   851968
         _ExtentX        =   11562
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vPrecioReferencia 
         Height          =   285
         Left            =   2760
         TabIndex        =   25
         Top             =   1230
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
         Left            =   2760
         TabIndex        =   26
         Top             =   1560
         Width           =   6555
         _Version        =   851968
         _ExtentX        =   11562
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vFleteC100Kgrs 
         Height          =   285
         Left            =   2760
         TabIndex        =   27
         Top             =   1890
         Width           =   6555
         _Version        =   851968
         _ExtentX        =   11562
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vPuertoOLugarDeReferencia 
         Height          =   285
         Left            =   2760
         TabIndex        =   28
         Top             =   2220
         Width           =   5595
         _Version        =   851968
         _ExtentX        =   9869
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vNrosDeCertificadosDepositoIntransferible 
         Height          =   285
         Left            =   -66340
         TabIndex        =   34
         Top             =   360
         Visible         =   0   'False
         Width           =   5895
         _Version        =   851968
         _ExtentX        =   10398
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vGrado 
         Height          =   285
         Index           =   1
         Left            =   -66340
         TabIndex        =   35
         Top             =   690
         Visible         =   0   'False
         Width           =   5895
         _Version        =   851968
         _ExtentX        =   10398
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vContenidoProteico 
         Height          =   285
         Left            =   -66340
         TabIndex        =   36
         Top             =   1020
         Visible         =   0   'False
         Width           =   5895
         _Version        =   851968
         _ExtentX        =   10398
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vFactor 
         Height          =   285
         Left            =   -66340
         TabIndex        =   37
         Top             =   1350
         Visible         =   0   'False
         Width           =   5895
         _Version        =   851968
         _ExtentX        =   10398
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vFormaDePago 
         Height          =   285
         Left            =   -66340
         TabIndex        =   38
         Top             =   1680
         Visible         =   0   'False
         Width           =   5895
         _Version        =   851968
         _ExtentX        =   10398
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vPrecioOperacion 
         Height          =   285
         Left            =   -66340
         TabIndex        =   39
         Top             =   2010
         Visible         =   0   'False
         Width           =   5895
         _Version        =   851968
         _ExtentX        =   10398
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vPesoNeto 
         Height          =   285
         Left            =   -66340
         TabIndex        =   40
         Top             =   2340
         Visible         =   0   'False
         Width           =   5895
         _Version        =   851968
         _ExtentX        =   10398
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vImporteBruto 
         Height          =   285
         Left            =   -66340
         TabIndex        =   41
         Top             =   2670
         Visible         =   0   'False
         Width           =   5895
         _Version        =   851968
         _ExtentX        =   10398
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vIVASImporteBruto 
         Height          =   285
         Left            =   -66340
         TabIndex        =   42
         Top             =   3000
         Visible         =   0   'False
         Width           =   5895
         _Version        =   851968
         _ExtentX        =   10398
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vAlmacenaje 
         Height          =   285
         Left            =   -67540
         TabIndex        =   52
         Top             =   1110
         Visible         =   0   'False
         Width           =   6675
         _Version        =   851968
         _ExtentX        =   11774
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vOtros 
         Height          =   285
         Left            =   -67540
         TabIndex        =   53
         Top             =   1440
         Visible         =   0   'False
         Width           =   6675
         _Version        =   851968
         _ExtentX        =   11774
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vIVASDeducciones 
         Height          =   285
         Left            =   -67540
         TabIndex        =   54
         Top             =   1770
         Visible         =   0   'False
         Width           =   6675
         _Version        =   851968
         _ExtentX        =   11774
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vIngresosBrutos 
         Height          =   285
         Left            =   -67750
         TabIndex        =   58
         Top             =   1020
         Visible         =   0   'False
         Width           =   7215
         _Version        =   851968
         _ExtentX        =   12726
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vSellosYRegistro 
         Height          =   285
         Left            =   -67750
         TabIndex        =   59
         Top             =   1350
         Visible         =   0   'False
         Width           =   7215
         _Version        =   851968
         _ExtentX        =   12726
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vRetencionIVA 
         Height          =   285
         Left            =   -67750
         TabIndex        =   60
         Top             =   1680
         Visible         =   0   'False
         Width           =   7215
         _Version        =   851968
         _ExtentX        =   12726
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vRetencionGanancias 
         Height          =   285
         Left            =   -67750
         TabIndex        =   61
         Top             =   2010
         Visible         =   0   'False
         Width           =   7215
         _Version        =   851968
         _ExtentX        =   12726
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vLocalidad 
         Height          =   285
         Left            =   -65560
         TabIndex        =   66
         Top             =   660
         Visible         =   0   'False
         Width           =   4935
         _Version        =   851968
         _ExtentX        =   8705
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vImporteNetoAPagar 
         Height          =   285
         Left            =   -65560
         TabIndex        =   67
         Top             =   990
         Visible         =   0   'False
         Width           =   5445
         _Version        =   851968
         _ExtentX        =   9604
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vPagoIVARES1394 
         Height          =   285
         Left            =   -65560
         TabIndex        =   68
         Top             =   1320
         Visible         =   0   'False
         Width           =   5445
         _Version        =   851968
         _ExtentX        =   9604
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vFecha 
         Height          =   285
         Index           =   1
         Left            =   -65560
         TabIndex        =   69
         Top             =   1650
         Visible         =   0   'False
         Width           =   5445
         _Version        =   851968
         _ExtentX        =   9604
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vPagoSCondiciones 
         Height          =   285
         Left            =   -65560
         TabIndex        =   70
         Top             =   1980
         Visible         =   0   'False
         Width           =   5445
         _Version        =   851968
         _ExtentX        =   9604
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vNombreApellidoFirmaDNIVendedor 
         Height          =   285
         Index           =   2
         Left            =   -65560
         TabIndex        =   71
         Top             =   2310
         Visible         =   0   'False
         Width           =   5445
         _Version        =   851968
         _ExtentX        =   9604
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vNombreApellidoFirmaDNIComprador 
         Height          =   285
         Left            =   -65560
         TabIndex        =   72
         Top             =   2640
         Visible         =   0   'False
         Width           =   5445
         _Version        =   851968
         _ExtentX        =   9604
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin VB.Label lblPagoS 
         Alignment       =   1  'Right Justify
         Caption         =   "> Pago s/condiciones:"
         Height          =   255
         Left            =   -67480
         TabIndex        =   79
         Top             =   1980
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lblFecha2 
         Alignment       =   1  'Right Justify
         Caption         =   "> Fecha:"
         Height          =   255
         Left            =   -67480
         TabIndex        =   78
         Top             =   1650
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lblPagoIVA 
         Alignment       =   1  'Right Justify
         Caption         =   "> Pago IVA RES. 1394:"
         Height          =   255
         Left            =   -67480
         TabIndex        =   77
         Top             =   1320
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lblImporteNeto 
         Alignment       =   1  'Right Justify
         Caption         =   "> Importe Neto a Pagar:"
         Height          =   255
         Left            =   -67480
         TabIndex        =   76
         Top             =   990
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lblLocalidad 
         Alignment       =   1  'Right Justify
         Caption         =   "> Localidad:"
         Height          =   255
         Left            =   -67480
         TabIndex        =   75
         Top             =   660
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblPagoS45646 
         Alignment       =   1  'Right Justify
         Caption         =   "> Nombre y Apellido, Firma y N° Documento del Comprador:"
         Height          =   255
         Left            =   -69910
         TabIndex        =   74
         Top             =   2640
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Label lblFechaafsafs 
         Alignment       =   1  'Right Justify
         Caption         =   "> Nombre y Apellido, Firma y N° Documento del Vendedor:"
         Height          =   255
         Left            =   -69880
         TabIndex        =   73
         Top             =   2310
         Visible         =   0   'False
         Width           =   4185
      End
      Begin VB.Label lblRetenciónGanancias 
         Alignment       =   1  'Right Justify
         Caption         =   "> Retención Ganancias:"
         Height          =   255
         Left            =   -69880
         TabIndex        =   65
         Top             =   2010
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblRetenciónIVA 
         Alignment       =   1  'Right Justify
         Caption         =   "> Retención IVA:"
         Height          =   255
         Left            =   -69880
         TabIndex        =   64
         Top             =   1680
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblSellosY 
         Alignment       =   1  'Right Justify
         Caption         =   "> Sellos y Registro:"
         Height          =   255
         Left            =   -69880
         TabIndex        =   63
         Top             =   1350
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblIngresosBrutos 
         Alignment       =   1  'Right Justify
         Caption         =   "> Ingresos Brutos:"
         Height          =   255
         Left            =   -69910
         TabIndex        =   62
         Top             =   1020
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lblIVASD 
         Alignment       =   1  'Right Justify
         Caption         =   "> IVA s/deducciones:"
         Height          =   255
         Left            =   -69670
         TabIndex        =   57
         Top             =   1770
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lblOtros 
         Alignment       =   1  'Right Justify
         Caption         =   "> Otros:"
         Height          =   255
         Left            =   -69670
         TabIndex        =   56
         Top             =   1440
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lblAlmacenaje 
         Alignment       =   1  'Right Justify
         Caption         =   "> Almacenaje:"
         Height          =   255
         Left            =   -69700
         TabIndex        =   55
         Top             =   1110
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblFormaDe 
         Alignment       =   1  'Right Justify
         Caption         =   "> Forma de Pago:"
         Height          =   255
         Left            =   -68380
         TabIndex        =   51
         Top             =   1680
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lblFactor 
         Alignment       =   1  'Right Justify
         Caption         =   "> Factor:"
         Height          =   255
         Left            =   -68380
         TabIndex        =   50
         Top             =   1350
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lblContenidoProteico 
         Alignment       =   1  'Right Justify
         Caption         =   "> Contenido Proteico:"
         Height          =   255
         Left            =   -68380
         TabIndex        =   49
         Top             =   1020
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lblGrado2 
         Alignment       =   1  'Right Justify
         Caption         =   "> Grado:"
         Height          =   255
         Left            =   -68380
         TabIndex        =   48
         Top             =   690
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lblNrosDe 
         Alignment       =   1  'Right Justify
         Caption         =   "> Nros. De certificados Deposito Intransferible :"
         Height          =   255
         Left            =   -69970
         TabIndex        =   47
         Top             =   360
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Label lblPrecioOperación 
         Alignment       =   1  'Right Justify
         Caption         =   "> Precio Operación:"
         Height          =   255
         Left            =   -68380
         TabIndex        =   46
         Top             =   2010
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lblPesoNeto 
         Alignment       =   1  'Right Justify
         Caption         =   "> Peso Neto:"
         Height          =   255
         Left            =   -68380
         TabIndex        =   45
         Top             =   2340
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lblImporteBruto 
         Alignment       =   1  'Right Justify
         Caption         =   "> Importe Bruto:"
         Height          =   255
         Left            =   -68380
         TabIndex        =   44
         Top             =   2670
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lblIVAS 
         Alignment       =   1  'Right Justify
         Caption         =   "> IVA s/Importe Bruto:"
         Height          =   255
         Left            =   -68380
         TabIndex        =   43
         Top             =   3000
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lblPuertoO 
         Alignment       =   1  'Right Justify
         Caption         =   "> Puerto o Lugar de Referencia:"
         Height          =   255
         Left            =   60
         TabIndex        =   33
         Top             =   2220
         Width           =   2385
      End
      Begin VB.Label lblFleteC 
         Alignment       =   1  'Right Justify
         Caption         =   "> Flete c/100 Kgrs:"
         Height          =   255
         Left            =   660
         TabIndex        =   32
         Top             =   1890
         Width           =   1785
      End
      Begin VB.Label lblGrado 
         Alignment       =   1  'Right Justify
         Caption         =   "> Grado:"
         Height          =   255
         Left            =   660
         TabIndex        =   31
         Top             =   1560
         Width           =   1785
      End
      Begin VB.Label lblPrecioReferencia 
         Alignment       =   1  'Right Justify
         Caption         =   "> Precio Referencia:"
         Height          =   255
         Left            =   660
         TabIndex        =   30
         Top             =   1230
         Width           =   1785
      End
      Begin VB.Label lblFecha 
         Alignment       =   1  'Right Justify
         Caption         =   "> Fecha:"
         Height          =   255
         Left            =   660
         TabIndex        =   29
         Top             =   900
         Width           =   1815
      End
      Begin VB.Label lblIndiqueSi 
         Caption         =   "Indique si actuo un corredor: "
         Height          =   225
         Left            =   -69850
         TabIndex        =   6
         Top             =   420
         Visible         =   0   'False
         Width           =   2295
      End
   End
   Begin XtremeSuiteControls.GroupBox GBTipoComprobante 
      Height          =   645
      Left            =   3270
      TabIndex        =   0
      Top             =   -30
      Width           =   14760
      _Version        =   851968
      _ExtentX        =   26035
      _ExtentY        =   1138
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.FlatEdit txtNroComprobante 
         Height          =   315
         Left            =   5340
         TabIndex        =   1
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
         Left            =   3630
         TabIndex        =   7
         Top             =   300
         Width           =   1905
      End
   End
End
Attribute VB_Name = "frmFormulario1116"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sqlCampos As String


Function fsqlValor() As String
Dim cadena As String
' Alfredo: tenes que completar esto con los demas valores. Tené en cuenta el orden según los nros de la referencia

cadena = cadena + Me.vCuit + "," '21
cadena = cadena + Me.vCuit + "," '22
'cadena = cadena + Me.votrotextboxquevoshayapuestoenelformulario + ","

fsqlValor = cadena
End Function


Function fsqlCampos() As String
Dim cadena As String
' Alfredo: tenes que completar esto con los demas campos. Tené en cuenta el orden según los nros de la referencia

cadena = cadena + "cuit" + "," '21
cadena = cadena + "otro campo" + "," '22
'cadena = cadena + Me.votrotextboxquevoshayapuestoenelformulario + ","

fsqlCampos = cadena
End Function

Private Sub Form_Load()
Me.Height = 5025
Me.Width = 12615


ArmarArbol

End Sub

Private Sub finit()
' init
End Sub


Private Sub ArmarArbol()
Dim nodX As Node
Dim nodX2 As Node
Dim nodX3 As Node
Dim nodX4 As Node


Set nodX = TreeView3.Nodes.Add(, "P", , "Compra/Venta")


Set nodX = TreeView3.Nodes.Add(1, tvwChild, "D1", "Flete deducido")
Set nodX = TreeView3.Nodes.Add("D1", tvwChild, , "Liq. Final")
Set nodX = TreeView3.Nodes.Add("D1", tvwChild, , "Liq. Parcial")
Set nodX = TreeView3.Nodes.Add("D1", tvwChild, , "Liq. Final del Parcial")

Set nodX = TreeView3.Nodes.Add(1, tvwChild, "N1", "Flete en deducción")
Set nodX = TreeView3.Nodes.Add("N1", tvwChild, , "Liq. Final")
Set nodX = TreeView3.Nodes.Add("N1", tvwChild, , "Liq. Parcial")
Set nodX = TreeView3.Nodes.Add("N1", tvwChild, , "Liq. Final del Parcial")



Set nodX2 = TreeView3.Nodes.Add(, "P", , "Nandato/Consignación")

Set nodX2 = TreeView3.Nodes.Add(1, tvwChild, "DD1", "Flete deducido")
Set nodX2 = TreeView3.Nodes.Add("DD1", tvwChild, , "Liq. Final")
Set nodX2 = TreeView3.Nodes.Add("DD1", tvwChild, , "Liq. Parcial")
Set nodX2 = TreeView3.Nodes.Add("DD1", tvwChild, , "Liq. Final del Parcial")

Set nodX2 = TreeView3.Nodes.Add(1, tvwChild, "NN1", "Flete en deducción")
Set nodX2 = TreeView3.Nodes.Add("NN1", tvwChild, , "Liq. Final")
Set nodX2 = TreeView3.Nodes.Add("NN1", tvwChild, , "Liq. Parcial")
Set nodX2 = TreeView3.Nodes.Add("NN1", tvwChild, , "Liq. Final del Parcial")

End Sub

'Private Sub OptSI_Click()

'If OptSI Then

'    fActuaCarredor

'End If

'End Sub


' -------------------------------- funciones del módulo  ----------------------------

'Private Sub fActuaCarredor()
' fActuaCarredor
' completar
'id_corredor = fTraeDatosTabla("corredor", "id_corredor")

'fllenarCamposCorredor (id_corredor)

'End Sub

'Private Sub PbAcciones_Click(Index As Integer)
'guardarComprador (comprador.fsql)
'End Sub


