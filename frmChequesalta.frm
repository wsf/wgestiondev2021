VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "KlexGrid.ocx"
Object = "{9746E3DA-06E1-4D26-9CE4-D9F6411A9C70}#1.0#0"; "SMGA_OcxTxt2009.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#13.0#0"; "Codejock.DockingPane.v13.0.0.Demo.ocx"
Begin VB.Form frmChequesAlta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario de datos completos  del cheque"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   12225
   Icon            =   "frmChequesalta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   12225
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   555
      Left            =   0
      TabIndex        =   56
      Top             =   -60
      Width           =   12195
      _Version        =   851968
      _ExtentX        =   21511
      _ExtentY        =   979
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton PusControlarInvariantes 
         Height          =   345
         Left            =   9210
         TabIndex        =   67
         Top             =   150
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Controlar Invariantes"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   0
         Left            =   60
         TabIndex        =   57
         Top             =   150
         Width           =   1365
         _Version        =   851968
         _ExtentX        =   2408
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Grabar <F2>"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmChequesalta.frx":6852
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   2
         Left            =   11040
         TabIndex        =   58
         Top             =   150
         Visible         =   0   'False
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Cerrar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmChequesalta.frx":6C59
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   1
         Left            =   1440
         TabIndex        =   59
         Top             =   150
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Limpiar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmChequesalta.frx":7059
      End
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   1215
      Left            =   90
      TabIndex        =   26
      Top             =   420
      Width           =   12105
      _Version        =   851968
      _ExtentX        =   21352
      _ExtentY        =   2143
      _StockProps     =   68
      AllowReorder    =   -1  'True
      PaintManager.OneNoteColors=   -1  'True
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   255
         Left            =   -360
         TabIndex        =   65
         Top             =   0
         Width           =   12465
         _Version        =   851968
         _ExtentX        =   21987
         _ExtentY        =   450
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
      End
      Begin Aplisoft_CajasDeTexto.TxF dtpFecha 
         Height          =   315
         Index           =   0
         Left            =   1620
         TabIndex        =   0
         Top             =   420
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   13
         Left            =   1620
         TabIndex        =   1
         Top             =   780
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   255
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   285
         Index           =   1
         Left            =   6240
         TabIndex        =   45
         Tag             =   "CodigoCliente"
         Top             =   390
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   0
         Left            =   5400
         TabIndex        =   46
         Top             =   390
         Visible         =   0   'False
         Width           =   795
         _Version        =   851968
         _ExtentX        =   1402
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   1
         Left            =   6630
         TabIndex        =   47
         Top             =   390
         Visible         =   0   'False
         Width           =   5355
         _Version        =   851968
         _ExtentX        =   9446
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha2 
         Height          =   315
         Left            =   5400
         TabIndex        =   48
         Top             =   750
         Width           =   1515
         _Version        =   851968
         _ExtentX        =   2672
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha3 
         Height          =   315
         Left            =   7290
         TabIndex        =   49
         Top             =   750
         Width           =   4725
         _Version        =   851968
         _ExtentX        =   8334
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   285
         Index           =   2
         Left            =   6930
         TabIndex        =   50
         Tag             =   "Proveedor"
         Top             =   780
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Index           =   0
         Left            =   4740
         TabIndex        =   52
         Top             =   420
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor:"
         Height          =   195
         Index           =   1
         Left            =   4530
         TabIndex        =   51
         Top             =   810
         Width           =   825
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Importe:"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   8
         Left            =   960
         TabIndex        =   43
         Top             =   810
         Width           =   675
      End
      Begin VB.Label lblLabel1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mensaje al usuario ..."
         Height          =   195
         Left            =   2100
         TabIndex        =   29
         Top             =   60
         Width           =   7125
      End
      Begin VB.Label lblAlta 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Cheque :"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   28
         Top             =   480
         Width           =   1410
      End
   End
   Begin XtremeSuiteControls.TabControl TabAlta 
      Height          =   6615
      Left            =   90
      TabIndex        =   9
      Top             =   1680
      Width           =   12105
      _Version        =   851968
      _ExtentX        =   21352
      _ExtentY        =   11668
      _StockProps     =   68
      AllowReorder    =   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      ItemCount       =   3
      Item(0).Caption =   "Datos del Cheque"
      Item(0).ControlCount=   38
      Item(0).Control(0)=   "lblFicha(3)"
      Item(0).Control(1)=   "txtFicha(4)"
      Item(0).Control(2)=   "lblFicha(5)"
      Item(0).Control(3)=   "lblFicha(6)"
      Item(0).Control(4)=   "lblFicha(7)"
      Item(0).Control(5)=   "txtFicha(5)"
      Item(0).Control(6)=   "txtFicha(6)"
      Item(0).Control(7)=   "txtFicha(7)"
      Item(0).Control(8)=   "txtFicha(8)"
      Item(0).Control(9)=   "txtFicha(9)"
      Item(0).Control(10)=   "txtFicha(10)"
      Item(0).Control(11)=   "pbCarga(3)"
      Item(0).Control(12)=   "lblFicha(4)"
      Item(0).Control(13)=   "lblFicha(2)"
      Item(0).Control(14)=   "pbCarga(4)"
      Item(0).Control(15)=   "txtFicha(11)"
      Item(0).Control(16)=   "lblFicha(9)"
      Item(0).Control(17)=   "dtpFecha(1)"
      Item(0).Control(18)=   "txtFicha(12)"
      Item(0).Control(19)=   "txtFicha(14)"
      Item(0).Control(20)=   "lblFicha(10)"
      Item(0).Control(21)=   "pbCarga(5)"
      Item(0).Control(22)=   "txtFicha(15)"
      Item(0).Control(23)=   "lblFicha(11)"
      Item(0).Control(24)=   "txtAlta(0)"
      Item(0).Control(25)=   "lblAlta(0)"
      Item(0).Control(26)=   "lblFicha(12)"
      Item(0).Control(27)=   "vCodCustodia"
      Item(0).Control(28)=   "PushButton1"
      Item(0).Control(29)=   "VDescCustodia"
      Item(0).Control(30)=   "lblFicha(13)"
      Item(0).Control(31)=   "vSucursal"
      Item(0).Control(32)=   "lblFicha(14)"
      Item(0).Control(33)=   "vmarcaInterna"
      Item(0).Control(34)=   "PushButton2"
      Item(0).Control(35)=   "PusBorrarCustodia"
      Item(0).Control(36)=   "txtFicha(2)"
      Item(0).Control(37)=   "lblFicha(15)"
      Item(1).Caption =   "Foto del cheque"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "DockingPaneFrame1"
      Item(2).Caption =   "Historiales"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "Frame1"
      Item(2).Control(1)=   "FraFiltro"
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   315
         Left            =   7890
         TabIndex        =   64
         Top             =   4350
         Width           =   525
         _Version        =   851968
         _ExtentX        =   926
         _ExtentY        =   556
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Picture         =   "frmChequesalta.frx":D8BB
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   285
         Left            =   3600
         TabIndex        =   8
         Top             =   3960
         Width           =   375
         _Version        =   851968
         _ExtentX        =   661
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit vCodCustodia 
         Height          =   315
         Left            =   2430
         TabIndex        =   24
         Top             =   3930
         Width           =   1155
         _Version        =   851968
         _ExtentX        =   2037
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
      End
      Begin VB.Frame FraFiltro 
         Caption         =   "Filtro:"
         Height          =   675
         Left            =   -69850
         TabIndex        =   32
         Top             =   360
         Visible         =   0   'False
         Width           =   11925
         Begin VB.ComboBox vTipoHistorial 
            Height          =   315
            ItemData        =   "frmChequesalta.frx":DE55
            Left            =   4770
            List            =   "frmChequesalta.frx":DE6E
            TabIndex        =   39
            Top             =   240
            Width           =   4065
         End
         Begin XtremeSuiteControls.PushButton PbAcciones 
            Height          =   345
            Index           =   3
            Left            =   90
            TabIndex        =   35
            Top             =   210
            Width           =   1095
            _Version        =   851968
            _ExtentX        =   1931
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "Filtro1"
            UseVisualStyle  =   -1  'True
            Picture         =   "frmChequesalta.frx":DF65
         End
         Begin XtremeSuiteControls.PushButton PbAcciones 
            Height          =   345
            Index           =   4
            Left            =   2250
            TabIndex        =   36
            Top             =   210
            Width           =   1095
            _Version        =   851968
            _ExtentX        =   1931
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "Filtro3"
            UseVisualStyle  =   -1  'True
            Picture         =   "frmChequesalta.frx":E36C
         End
         Begin XtremeSuiteControls.PushButton PbAcciones 
            Height          =   345
            Index           =   5
            Left            =   1170
            TabIndex        =   37
            Top             =   210
            Width           =   1095
            _Version        =   851968
            _ExtentX        =   1931
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "Filtro2"
            UseVisualStyle  =   -1  'True
            Picture         =   "frmChequesalta.frx":E76C
         End
         Begin VB.Label lblTipoHistorial 
            Caption         =   "> Tipo Historial:"
            Height          =   255
            Left            =   3570
            TabIndex        =   38
            Top             =   270
            Width           =   1125
         End
      End
      Begin VB.Frame Frame1 
         ClipControls    =   0   'False
         Height          =   5475
         Left            =   -69880
         TabIndex        =   30
         Top             =   1020
         Visible         =   0   'False
         Width           =   11925
         Begin Grid.KlexGrid GrillaHistorialCheque 
            Height          =   4605
            Left            =   180
            TabIndex        =   31
            Top             =   690
            Width           =   11655
            _ExtentX        =   20558
            _ExtentY        =   8123
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
            MouseIcon       =   "frmChequesalta.frx":14FCE
            Rows            =   10
         End
         Begin XtremeSuiteControls.ComboBox vModoVisualizacion 
            Height          =   315
            Left            =   1800
            TabIndex        =   33
            Top             =   150
            Width           =   7065
            _Version        =   851968
            _ExtentX        =   12462
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin VB.Label lblTipoDe 
            Caption         =   "Tipo de visualización:"
            Height          =   315
            Left            =   90
            TabIndex        =   34
            Top             =   180
            Width           =   1605
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   5
         Left            =   2460
         TabIndex        =   3
         Top             =   1170
         Width           =   2715
         _Version        =   851968
         _ExtentX        =   4789
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   6
         Left            =   2430
         TabIndex        =   4
         Top             =   1560
         Width           =   9615
         _Version        =   851968
         _ExtentX        =   16960
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   8
         Left            =   4200
         TabIndex        =   18
         Top             =   1980
         Width           =   7845
         _Version        =   851968
         _ExtentX        =   13838
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "No Acreditado"
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   3
         Left            =   3750
         TabIndex        =   15
         Tag             =   "EstadoCheque"
         Top             =   1980
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   7
         Left            =   2430
         TabIndex        =   16
         Top             =   1980
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "4"
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   1725
         Index           =   14
         Left            =   2430
         TabIndex        =   27
         Top             =   4740
         Width           =   6015
         _Version        =   851968
         _ExtentX        =   10610
         _ExtentY        =   3043
         _StockProps     =   77
         BackColor       =   -2147483643
         MaxLength       =   250
         ScrollBars      =   2
      End
      Begin Aplisoft_CajasDeTexto.TxF dtpFecha 
         Height          =   315
         Index           =   1
         Left            =   2430
         TabIndex        =   7
         Top             =   3570
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   4
         Left            =   2460
         TabIndex        =   11
         Top             =   780
         Width           =   1635
         _Version        =   851968
         _ExtentX        =   2884
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   15
         Left            =   5160
         TabIndex        =   12
         Top             =   780
         Width           =   6885
         _Version        =   851968
         _ExtentX        =   12144
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeDockingPane.DockingPaneFrame DockingPaneFrame1 
         Height          =   4755
         Left            =   -69430
         TabIndex        =   41
         Top             =   900
         Visible         =   0   'False
         Width           =   11055
         _Version        =   851968
         _ExtentX        =   19500
         _ExtentY        =   8387
         _StockProps     =   0
      End
      Begin XtremeSuiteControls.FlatEdit txtAlta 
         Height          =   315
         Index           =   0
         Left            =   2460
         TabIndex        =   2
         Top             =   420
         Width           =   9585
         _Version        =   851968
         _ExtentX        =   16907
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   9
         Left            =   2430
         TabIndex        =   19
         Top             =   2430
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   4
         Left            =   3750
         TabIndex        =   5
         Tag             =   "Banco"
         Top             =   2430
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   10
         Left            =   4230
         TabIndex        =   20
         Top             =   2430
         Width           =   7815
         _Version        =   851968
         _ExtentX        =   13785
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   11
         Left            =   2430
         TabIndex        =   21
         Top             =   2790
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   5
         Left            =   3750
         TabIndex        =   53
         Tag             =   "BancoCuenta"
         Top             =   2790
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   12
         Left            =   4230
         TabIndex        =   22
         Top             =   2790
         Width           =   7815
         _Version        =   851968
         _ExtentX        =   13785
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit VDescCustodia 
         Height          =   315
         Left            =   4020
         TabIndex        =   25
         Top             =   3960
         Width           =   4425
         _Version        =   851968
         _ExtentX        =   7805
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Enabled         =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit vSucursal 
         Height          =   315
         Left            =   2430
         TabIndex        =   6
         Top             =   3180
         Width           =   9615
         _Version        =   851968
         _ExtentX        =   16960
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vmarcaInterna 
         Height          =   315
         Left            =   2430
         TabIndex        =   63
         Top             =   4350
         Width           =   5355
         _Version        =   851968
         _ExtentX        =   9446
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton PusBorrarCustodia 
         Height          =   315
         Left            =   8520
         TabIndex        =   66
         Top             =   3960
         Width           =   1575
         _Version        =   851968
         _ExtentX        =   2778
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Limpiar custodia"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmChequesalta.frx":14FEA
      End
      Begin XtremeSuiteControls.FlatEdit txtFicha 
         Height          =   315
         Index           =   2
         Left            =   5940
         TabIndex        =   68
         Top             =   1200
         Width           =   6105
         _Version        =   851968
         _ExtentX        =   10769
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin VB.Label lblFicha 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Endoso:"
         Height          =   195
         Index           =   15
         Left            =   5130
         TabIndex        =   69
         Top             =   1260
         Width           =   765
      End
      Begin VB.Label lblFicha 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Marca interna: "
         Height          =   195
         Index           =   14
         Left            =   540
         TabIndex        =   62
         Top             =   4350
         Width           =   1755
      End
      Begin VB.Label lblFicha 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sucursal:"
         Height          =   195
         Index           =   13
         Left            =   540
         TabIndex        =   61
         Top             =   3240
         Width           =   1755
      End
      Begin VB.Label lblFicha 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "En custodia de:"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   12
         Left            =   480
         TabIndex        =   60
         Top             =   4020
         Width           =   1755
      End
      Begin VB.Label lblFicha 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Banco:"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   6
         Left            =   510
         TabIndex        =   55
         Top             =   2460
         Width           =   1755
      End
      Begin VB.Label lblFicha 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
         Height          =   195
         Index           =   7
         Left            =   510
         TabIndex        =   54
         Top             =   2850
         Width           =   1755
      End
      Begin VB.Label lblAlta 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Numero de Cheque :"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   510
         TabIndex        =   44
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Cod.Barra:"
         Height          =   195
         Index           =   11
         Left            =   4320
         TabIndex        =   42
         Top             =   810
         Width           =   975
      End
      Begin VB.Label lblFicha 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Interno:"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   510
         TabIndex        =   40
         Top             =   840
         Width           =   1755
      End
      Begin VB.Label lblFicha 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones:"
         Height          =   195
         Index           =   10
         Left            =   510
         TabIndex        =   23
         Top             =   4710
         Width           =   1755
      End
      Begin VB.Label lblFicha 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Estado del Cheque:"
         Height          =   195
         Index           =   5
         Left            =   510
         TabIndex        =   17
         Top             =   2040
         Width           =   1755
      End
      Begin VB.Label lblFicha 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ref. Receptor del Cheque:"
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   14
         Top             =   1650
         Width           =   2055
      End
      Begin VB.Label lblFicha 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Acreditación:"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   9
         Left            =   840
         TabIndex        =   13
         Top             =   3600
         Width           =   1485
      End
      Begin VB.Label lblFicha 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Firmante:"
         Height          =   195
         Index           =   3
         Left            =   510
         TabIndex        =   10
         Top             =   1230
         Width           =   1755
      End
   End
End
Attribute VB_Name = "frmChequesAlta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vViene As String
Public vaccion As String
Public vnrointernoGral As Long

Private Sub Form_KeyPress(Keyascii As Integer)
    On Error Resume Next
    
    If Keyascii = 13 Then SendKeys "{TAB}"

    If Err Then GrabarLog "Form_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    If KeyCode = vbKeyF1 Then
        VerAyuda (Me.Name)
    End If
    
    If KeyCode = vbKeyF2 Then
        Grabar
    End If

If Err Then GrabarLog "Form_KeyUp", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub Form_Load()
    On Error Resume Next

    init 'inicializa la entrada

    With Me
       ' .Show
       ' .Top = 0
       ' .Left = 0
        '.Width = 9500
        '.Height = 7110
    End With
    
    Call LimpiarCampos
    
    CentrarFormulario Me
    
    If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub init()
    If vaccion = "" Then vaccion = "Nuevo"
    
    Me.vmarcaInterna = getMarcaIntarna()
    cargardatosdeIE
    cargardatosdeCobros
    CentrarFormulario (Me)
End Sub
Private Sub cargardatosdeCobros()

If Me.vViene = "cobro" Then
gbldsCheques.Codigo = frmCobros.txtCliente(0)
gbldsCheques.Nombre = frmCobros.txtCliente(1)
gbldsCheques.fecha = frmCobros.vfechaCredito.Value
pasarDsAfrom
End If


If Me.vViene = "pago" Then
pasarDsAfrom
End If




End Sub

Public Sub cargardatosdeIE()
' Alfredo: en esta funcion cargo los txt con los datos del data set cargado en el frmIngresosEgresos


If Not Me.Tag = "frmIngresosEgresos" Then Exit Sub
With gbldsCheques
Me.dtpFecha(0) = .fecha
Me.txtAlta(0) = .Ncheque
Me.txtFicha(13) = .monto
Me.dtpFecha(1) = .FechaDeposito
Me.txtFicha(14) = .Observaciones
Me.txtFicha(4) = .NroInterno
End With








End Sub

Private Sub PbAcciones_Click(Index As Integer)
    On Error Resume Next

    Select Case Index
    
        Case 0
        
           
            
            Call Grabar
            Exit Sub
          
        
        Case 1
            Call LimpiarCampos
        
        Case 2
            Unload Me

    End Select

    If Err Then GrabarLog "PbAcciones_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub Grabar()
On Error Resume Next
Dim vsql As String
Dim vIdCheques, vnrointerno As Long


'If Not Me.Tag = "" Then ' en el caso que venga de otro módulo solamente actualizo el dataset
   ' cargarDataSetCheques
    'Exit Sub
'End If

vnrointerno = UltimoNroInterno2 + 1

Me.vnrointernoGral = Val(Me.vmarcaInterna)

    If Not ValidarCampos() = True Then
        Exit Sub
    End If
    
    Dim rsCheques As New ADODB.Recordset, sqlCheques As String
    
    Select Case vaccion

        Case "Nuevo"
            sqlCheques = "SELECT * FROM Cheques WHERE 1=2"
        
        Case "Modificar"
            sqlCheques = "SELECT * FROM Cheques WHERE idCheques = " + Str(gbldsCheques.vid)
        
        Case "Duplicar"
            
    End Select
        
    With rsCheques
        Call .Open(sqlCheques, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not .State = 0 Then
            If vaccion = "Nuevo" Then .AddNew
                    .Fields("NCheque").Value = Trim(txtAlta(0).Text)
                    gbldsCheques.Ncheque = Trim(txtAlta(0).Text)
                    
                    .Fields("FechaDeposito").Value = strfechaMySQL(dtpFecha(1).Value)
                    gbldsCheques.FechaDeposito = strfechaMySQL(dtpFecha(1).Value)
                    
                    .Fields("Fecha").Value = strfechaMySQL(dtpFecha(0).Value)
                    gbldsCheques.fecha = strfechaMySQL(dtpFecha(0).Value)
                
                    .Fields("TipoMovimiento").Value = ""
           
            
            'Ficha
            
            If Not Trim(txtFicha(0).Text) = "" Then
                .Fields("CP").Value = "C"
                .Fields("Codigo").Value = Left(txtFicha(0).Text, 50)
                gbldsCheques.Codigo = Left(txtFicha(0).Text, 50)
                
                
                .Fields("Nombre").Value = Left(txtFicha(1).Text, 255)
                gbldsCheques.Nombre = Left(txtFicha(1).Text, 255)
                
            Else
                If Not Trim(txtFicha2.Text) = "" Then
                    .Fields("CP").Value = "P"
                    .Fields("Codigo").Value = Left(txtFicha2.Text, 50)
                    gbldsCheques.Codigo = Left(txtFicha(0).Text, 50)
                     
                    .Fields("Nombre").Value = Left(txtFicha3.Text, 255)
                    gbldsCheques.Nombre = Left(txtFicha(1).Text, 255)
                End If
            End If
            
            .Fields("idEstadoCheque").Value = Val(txtFicha(7).Text)
            gbldsCheques.idEstadoCheque = Val(txtFicha(7).Text)
            
            .Fields("NroInterno").Value = Val(txtFicha(4).Text)
            gbldsCheques.NroInterno = Val(txtFicha(4).Text)

            .Fields("Firmante").Value = Left(txtFicha(5).Text, 50)
            gbldsCheques.Firmante = Left(txtFicha(5).Text, 50)
            
            .Fields("Endoso").Value = Left(txtFicha(6).Text, 255)
            gbldsCheques.Endoso = Left(txtFicha(6).Text, 255)
            
            
            .Fields("idBancos").Value = Left(txtFicha(9).Text, 3)
            gbldsCheques.idBancos = Left(txtFicha(9).Text, 3)
             
            
            .Fields("idBancosCuentas").Value = Val(txtFicha(11).Text)
            gbldsCheques.idBancosCuentas = Val(txtFicha(11).Text)
            
            .Fields("idCustodia").Value = Me.vCodCustodia
            gbldsCheques.idCustodia = Me.vCodCustodia
            
            .Fields("Observaciones").Value = Trim(txtFicha(14).Text)
            gbldsCheques.Observaciones = Trim(txtFicha(14).Text)
            
            .Fields("Monto").Value = Val(Format(txtFicha(13).Text, "######0.000"))
            gbldsCheques.monto = Val(Format(txtFicha(13).Text, "######0.000"))
            
            .Fields("FechaAcreditacion").Value = strfechaMySQL(dtpFecha(1).Value)
            gbldsCheques.FechaAcreditacion = strfechaMySQL(dtpFecha(1).Value)
                
            'Call GuardarFoto(rsCheques, phtCheque.PhotoFileName)
             .Fields("nrointerno").Value = vnrointerno
             gbldsCheques.NroInterno = vnrointerno
             
             
            vsql = "select max(idCheques) as c from cheques"
            vIdCheques = traerDatos2(vsql, "c", pathDBMySQL) + 1
    
           If Not vViene = "frmie.ingreso" Then Call asignarChequeACaja(vIdCheques)   ' ingresa un movimiento de caja

             vsql = "select max(idBancosMovimientos) as c from bancosmovimientos"
             .Fields("idbancocaja") = traerDatos2(vsql, "c", pathDBMySQL)
             
             
             .Fields("sucursal") = Me.vsucursal.Text
             
             .Fields("marcainterna") = Me.vmarcaInterna
             
             .Fields("propietarios") = fpropietario()
             
        
            .Update
        
        End If
        
    End With
    
     setMarcaInterna (Me.vmarcaInterna) ' graba la marca en el caso que grabe

    
    sqlCheques = ""
    
    If rsCheques.State = 1 Then
        rsCheques.Close
        Set rsCheques = Nothing
    End If
    
    
    
   ' grabar
   
     
    If vViene = "frmie.ingreso" And vnrointernoGral > 0 Then
            
                frmCheques.Show
                frmCheques.vbusca = Str(vnrointernoGral)
                frmCheques.vViene = "frmIngresosEgresos"
                
                vnrointernoGral = 0
                Unload Me
                Exit Sub
                
    End If
            
        
    
    If Err Then
        GrabarLog "Guardar", Err.Number & " " & Err.Description, Me.Name
    Else
        vnrointernoGral = vmarcaInterna
        LimpiarCampos
       ' Unload Me
        'frmCheques.bu
    End If
    
Call frmCheques.fInvariantes

End Sub

Function fpropietario() As String

If Me.txtFicha(0).Text = "" Then
    fpropietario = "Propio"
Else
    fpropietario = "Tercero"
End If

End Function
Private Sub asignarChequeACaja(ByVal vid As Long)
Dim vsql, vcampos, vvalores As String

vcampos = "idcheques,idBancos,fecha,fechaValor,debito,credito,nrocheque,nrointerno,comentario"
vvalores = Str(vid) + ",'" + gbldsCheques.idCustodia + "','" + strfechaMySQL(gbldsCheques.fecha) + "','" + strfechaMySQL(gbldsCheques.FechaDeposito) + "'," + Str(gbldsCheques.monto) + ",0,'" + Str(gbldsCheques.Ncheque) + "'," + Str(gbldsCheques.NroInterno) + ",'" + gbldsCheques.Observaciones + "'"

vsql = "insert into bancosmovimientos (" + vcampos + ") values (" + vvalores + ")"

Call EjecutarScript(vsql, pathDBMySQL)


End Sub

Private Function ValidarCampos() As Boolean
    On Error Resume Next

    Dim i As Integer
    
    ValidarCampos = True
    
   
   
    If Trim(Me.vCodCustodia.Text) = "" Then
            MsgBox "Debe ingresar una custodia del cheque"
            ValidarCampos = Not True
            Exit Function
    End If
    
   
   
   ' For i = 0 To Val(txtAlta.Count - 1)
   '     If Trim(txtAlta(i).Text) = "" Then
   '         MsgBox "Campos obligatorios vacios!", vbExclamation, "Mensaje ..."
   '         ValidarCampos = Not True
   '         Exit Function
   '     End If

    'Next
     If txtFicha(7) = "" Or txtAlta(0) = "" Or txtFicha(13) = "" Or txtFicha(9) = "" Or txtFicha(10) = "" Or vCodCustodia = "" Or VDescCustodia = "" Then
    'If Trim(txtFicha(7).Text) = "" Then
        
        
      
        
        If MsgBox("Campos obligatorios Vacios! " + Chr(13) + " Quiere guardarlos de todas maneras ?", vbYesNo, "Aletención !!!") = vbYes Then
                ValidarCampos = True
                Exit Function
        Else
                ValidarCampos = Not True
                Exit Function
        End If
        
        

    End If
    
 
    
    
    If vaccion = "Nuevo" Then
        If Not Trim(TraerDato("Cheques", "NCheque = '" & Trim(txtAlta(0).Text) & "'", "idCheques")) = "" Then
            MsgBox "Existe un registro con el mismo número de cheque!", vbExclamation, "Mensaje ..."
            ValidarCampos = Not True
            Exit Function
        End If
    End If
    
    If Err Then GrabarLog "ValidarCampos", Err.Number & " " & Err.Description, Me.Caption
End Function
Private Function GuardarFoto(rsFoto As ADODB.Recordset, _
                             vFilename As String) As Boolean
    On Error Resume Next
    
    If Not vFilename = "" Then
    
        Dim stmFoto As New ADODB.Stream
    
        With stmFoto

            If Not .State = adStateClosed Then .Close
            
            .Type = adTypeBinary
            .Open

            .LoadFromFile (vFilename)
        End With
    
    End If
    
    With rsFoto

        If Not (.State = 0) And Not (.EOF = True) Then
            If Not vFilename = "" Then
                .Fields("Foto").Value = stmFoto.Read
            Else
                .Fields("Foto").Value = Null
            End If
        End If
        
    End With
    
    If Not stmFoto.State = adStateClosed Then
        stmFoto.Close
        Set stmFoto = Nothing
    End If
    
    If Err Then
        GuardarFoto = Not True
        GrabarLog "GuardarFoto", Err.Number & " " & Err.Description, "Global"
    Else
        GuardarFoto = True
    End If

End Function
Private Sub LimpiarCampos()
    On Error Resume Next
    
    Dim i As Integer
    
    For i = 0 To txtAlta.Count - 1
        txtAlta(i).Text = ""
    Next
    
    'phtCheque.Reset
    'pbCierraFoto.Visible = Not True
    
    For i = 0 To txtFicha.Count - 1
        txtFicha(i).Text = ""
    Next
    
    txtFicha(7).Text = "2"
    txtFicha(8).Text = "No Acreditado"
    
    
    For i = 0 To dtpFecha.Count - 1
        dtpFecha(i).Value = Date
    Next
    
    Me.vCodCustodia = ""
    Me.VDescCustodia = ""
    
    Me.vmarcaInterna = getMarcaIntarna()
    Me.vsucursal = ""
    
   If Not vaccion = "Modificar" Then vaccion = "Nuevo"
    KeyPreview = True

    If Err Then GrabarLog "Limpia", Err.Number & "-" & Err.Description, Me.Name
End Sub

Private Sub pasarDsAfrom()
'gbldsCheques.Ncheque = txtAlta(0).Text
                    
'gbldsCheques.FechaDeposito = dtpFecha(1).Value
                     
'gbldsCheques.fecha = dtpFecha(0).Value
        
'If gbldsCheques.CP Is Nothing Then Set gbldsCheques.CP = "P"
        
If gbldsCheques.CP = "C" Then

    txtFicha(0).Text = gbldsCheques.Codigo
    Call txtFicha_Click(0)
                
    txtFicha(1).Text = gbldsCheques.Nombre

Else


    txtFicha(0).Text = gbldsCheques.Codigo
    Call txtFicha_Click(0)

    txtFicha(1).Text = gbldsCheques.Nombre
End If

txtAlta(0).Text = gbldsCheques.Ncheque

txtFicha(7).Text = gbldsCheques.idEstadoCheque
Call txtFicha_KeyPress(7, 13)
            
txtFicha(4).Text = gbldsCheques.NroInterno

 txtFicha(5).Text = gbldsCheques.Firmante
            
txtFicha(6).Text = gbldsCheques.Endoso

txtFicha(2).Text = gbldsCheques.Endoso
            
txtFicha(9).Text = gbldsCheques.idBancos
Call txtFicha_KeyPress(9, 13)

txtFicha(11).Text = gbldsCheques.idBancosCuentas
Me.VDescCustodia.Tag = gbldsCheques.idBancosCuentas
Call txtFicha_KeyPress(11, 13)
            
Me.vCodCustodia = gbldsCheques.idCustodia
Me.VDescCustodia.Tag = gbldsCheques.idCustodia
Call vCodCustodia_KeyPress(13)
            
txtFicha(14).Text = gbldsCheques.Observaciones
            
txtFicha(13).Text = gbldsCheques.monto
            
gbldsCheques.FechaAcreditacion = dtpFecha(1).Value


Me.vsucursal = gbldsCheques.sucursal

Me.vmarcaInterna = gbldsCheques.marcainterna

End Sub

Public Sub ModificarCheque(vIdCheques As Long)
    On Error Resume Next
    vaccion = "Modificar"
    pasarDsAfrom 'pasa los datos del ds al formulario
    Exit Sub

    
    Dim rsCheques As New ADODB.Recordset, sqlCheques As String
    
    sqlCheques = "SELECT * FROM Cheques WHERE (idCheques = " & vIdCheques & ")"
    
    With rsCheques
        Call .Open(sqlCheques, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not (.EOF = True) And Not (.BOF = True) Then
        
            'No Opcionales
            txtFicha(0).Text = EsNulo(.Fields("codigo").Value)
           ' txtFicha(0).Locked = True
        
            txtFicha(1).Text = EsNulo(.Fields("Nombre").Value)
           ' txtficha2.Text = EsNulo(.Fields("RazonSocial").Value)
        
         '   If Not IsNull(.Fields("Foto").Value) = True And Not Trim(.Fields("Foto").Value) = "" Then
         '       BorrarArchivo (App.Path & "\" & .Fields("Codigo").Value & ".dat")
         '       phtCheque.BlobToFile rsCheques!Foto, App.Path & "\" & .Fields("Codigo").Value & ".dat"
         '       Call phtCheque.AbrirFotoDesdeArchivo(App.Path & "\" & .Fields("Codigo").Value & ".dat")
         '       BorrarArchivo (App.Path & "\" & .Fields("Codigo").Value & ".dat")
         '       pbCierraFoto.Visible = True
         '   End If

            'Ficha
        
            txtFicha(0).Text = EsNulo(.Fields("Direccion").Value)
            txtFicha(1).Text = EsNulo(.Fields("CodigoPostal").Value)
            txtFicha2.Text = EsNulo(.Fields("Localidad").Value)
            txtFicha3.Text = EsNulo(.Fields("Provincia").Value)
            txtFicha(4).Text = EsNulo(.Fields("Pais").Value)
            txtFicha(5).Text = EsNulo(.Fields("Telefono").Value)
            txtFicha(6).Text = EsNulo(.Fields("Fax").Value)
            txtFicha(7).Text = EsNulo(.Fields("Celular").Value)
            txtFicha(8).Text = EsNulo(.Fields("idVendedor").Value)
            txtFicha(9).Text = EsNulo(TraerDato("Empleados", "Codigo =  '" & .Fields("idVendedor").Value & "'", "Nombre"))
        
            txtFicha(10).Text = EsNulo(.Fields("idReparto").Value)
            txtFicha(11).Text = EsNulo(TraerDato("clireparto", "nreparto =  '" & .Fields("idReparto").Value & "'", "descrip"))
            
            Me.vsucursal = EsNulo(.Fields("sucursal").Value)
            
            Me.vmarcaInterna = EsNulo(.Fields("marcainterna").Value)
            
            
            txtFicha(14).Text = EsNulo(.Fields("observaciones").Value)
            
            txtFicha(2).Text = EsNulo(.Fields("Endoso").Value)
            
        
            'Datos Comerciales
            'txtDatosComerciales(0).Text = EsNulo(.Fields("idTipoIva").Value)
            'txtDatosComerciales(1).Text = EsNulo(TraerDato("TipoIva", "idTipoIva = '" & .Fields("idTipoIva").Value & "'", "TipoIva"))
            'txtDatosComerciales(2).Text = EsNulo(.Fields("Cuit").Value)
            'txtDatosComerciales(3).Text = EsNulo(.Fields("idTipoCliente").Value)
            'txtDatosComerciales(4).Text = EsNulo(TraerDato("TipoClientes", "idTipoClientes = '" & .Fields("idTipoCliente").Value & "'", "Descripcion"))
            'txtDatosComerciales(5).Text = .Fields("idActividad").Value
            'txtDatosComerciales(6).Text = EsNulo(TraerDato("Actividades", "idActividades = '" & .Fields("idActividad").Value & "'", "Descripcion"))
            'txtDatosComerciales(7).Text = EsNulo(.Fields("idListas").Value)
            'txtDatosComerciales(8).Text = EsNulo(TraerDato("Listas", "idListas = '" & .Fields("idListas").Value & "'", "Lista"))
            'txtDatosComerciales(9).Text = EsNulo(.Fields("CreditoMax").Value)
            

            'Otros datos
            If Not IsNull(.Fields("Fecha_Alta").Value) = True Then
                'dtpAlta.Value = .Fields("Fecha_Alta").Value
            End If
            
            If Not IsNull(.Fields("Fecha_Nacimiento").Value) = True Then
                'dtpFechaNacimiento.Value = .Fields("Fecha_Nacimiento").Value
            End If
    
            'txtOtrosDatos(0).Text = EsNulo(.Fields("E-Mail").Value)
            'txtOtrosDatos(1).Text = EsNulo(.Fields("Web").Value)
            'txtOtrosDatos(2).Text = EsNulo(.Fields("Skype").Value)
            'txtOtrosDatos(3).Text = EsNulo(.Fields("MensajeEmergente").Value)
            'txtOtrosDatos(4).Text = EsNulo(.Fields("idEstados").Value)
            'txtOtrosDatos(5).Text = EsNulo(TraerDato("Estados", "idEstados = '" & .Fields("idEstados").Value & "'", "Estado"))
        
            'Observaciones
            'txtObservaciones.Text = EsNulo(.Fields("Observaciones").Value)
        
        End If

    End With
    
    If Err Then GrabarLog "ModificarCliente", Err.Number & " " & Err.Description, Me.Name
End Sub




Private Sub pbCarga_Click(Index As Integer)
    On Error Resume Next


    If Index = 2 Then
        Call fbuscarGrilla("proveedores", "Nombre", "Codigo", Me.txtFicha3.Name, Me, , False)

    Else
    
    
    
    
    vVuelveBusqueda = Me.Name ' Alfredo: en el caso de que sea nuevo tenes que ir a modificar el case de la fncion grilla evento doble clic
    vVieneBusqueda = pbCarga(Index).Tag ' Alfredo: tenes que poner el TAG del botòn el nombre de la tabla. En el caso de que sea nuevo y a modificar el case de la función form evento load

    Select Case Index
    
        Case 0 To pbCarga.Count
            frmBusqueda.Show

    End Select
End If


If Err Then GrabarLog "pbCarga_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub PusBorrarCustodia_Click()
Me.vCodCustodia.Text = ""
Me.VDescCustodia.Text = ""
Me.VDescCustodia.Tag = ""
End Sub

Private Sub PusControlarInvariantes_Click()
 Call frmCheques.fInvariantes
End Sub

Private Sub PushButton1_Click()
Call fbuscarGrilla("(select * from bancos where EsCaja = 'S') as t", "Descripcion", "idBancos", Me.VDescCustodia.Name, Me)
End Sub

Private Sub PushButton2_Click()
Me.vmarcaInterna = getMarcaIntarna()
End Sub

Private Sub txtFicha_Change(Index As Integer)
    On Error Resume Next

    'If Index = 1 Then txtficha2.Text = txtFicha(1).Text

    If Err Then GrabarLog "txtFicha_Change", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtFicha_KeyPress(Index As Integer, Keyascii As Integer)
On Error Resume Next

    If Keyascii = 13 Then
    
        txtFicha(Index).Text = UCase(txtFicha(Index).Text)
        
        Select Case Index
        
            Case 0
                
            Case 2
                txtFicha3.SetFocus
            
            Case 7
                txtFicha(Index + 1).Text = TraerDato("EstadoCheque", "idEstadoCheque = " & Trim(txtFicha(Index).Text) & "", "Descripcion")
            
                txtFicha(Index + 2).SetFocus
   
            Case 9
                 
                If Not Trim(txtFicha(Index).Text) = "" Then
                    txtFicha(Index + 1).Text = TraerDato("Bancos", "idBancos = '" & Trim(txtFicha(Index).Text) & "'", "Descripcion")
                
                    txtFicha(11).Text = ""
                    txtFicha(12).Text = ""


                End If
            
                txtFicha(11).SetFocus
            Case 8
                txtFicha(10).SetFocus
            
            Case 11
                
                If Not txtFicha(Index).Text = "" Then
                    
                    txtFicha(Index + 1).Text = TraerDato("BancosCuentas", "idBancosCuentas = " & Trim(txtFicha(Index).Text) & " AND (idBancos = '" & Trim(txtFicha(Index - 2).Text) & "')", "Cuenta")
                    If txtFicha(Index + 1).Text = "" Then
                        txtFicha(Index).Text = ""
                        txtFicha(Index).SetFocus
                    Else
                        txtFicha(Index + 2).SetFocus
                    End If
                Else
                    txtFicha(Index + 1).Text = ""
                    txtFicha(12).SetFocus
                End If
                

            Case 12
                txtFicha(Index + 1).SetFocus
        End Select
    
    End If

If Err Then GrabarLog "txtFicha_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtFicha_Click(Index As Integer)
On Error Resume Next

    txtFicha(Index).SelStart = 0
    txtFicha(Index).SelLength = Len(txtFicha(Index).Text)

If Err Then GrabarLog "txtFicha_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtFicha_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
    If KeyCode = vbKeyF3 Then
        If Index = 11 Then
           ' pbCarga_Click (5)
        End If
    End If
    
If Err Then GrabarLog "txtFicha_KeyUp", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub cargarDataSetCheques()
 

 With gbldsCheques
 
  .idEstadoCheque = Me.txtFicha(4)
  .fecha = Me.dtpFecha(0)
  
  If (Not Me.txtFicha(0) = "" And Not Me.txtFicha(1) = "") Then
    
    .Codigo = Me.txtFicha(0).TxT
    .Nombre = Me.txtFicha(1).TxT
  
  Else
    .Codigo = Me.txtFicha2.TxT
    .Nombre = Me.txtFicha3.TxT
    
  End If
  
  .idBancos = Me.txtFicha(9).TxT
  .idBancosCuentas = Me.txtFicha(10).TxT
  .Ncheque = Me.txtAlta(0).TxT
  .Firmante = Me.txtFicha(5).TxT
  'CP                 As Chart
  .FechaDeposito = Me.dtpFecha(1)
  .monto = Me.txtFicha(13)
 ' Endoso             As String
 'remito             As Integer
  .NroInterno = Me.txtFicha(4).TxT
  .Observaciones = Me.txtFicha(14).TxT
  .FechaAcreditacion = Me.dtpFecha(1)
  'Foto               As Long
  'TipoMovimiento     As String
End With
  
  
End Sub

Private Sub txtFicha3_Change()
txtFicha2.Text = txtFicha3.Tag
End Sub

Private Sub vCodCustodia_KeyPress(Keyascii As Integer)
Dim vsql As String

If Keyascii = 13 Then
    vsql = "select Descripcion as c from bancos where idBancos='" + vCodCustodia + "'"
    Me.VDescCustodia = traerDatos2(vsql, "c", pathDBMySQL)
    VDescCustodia.Tag = vCodCustodia
End If

End Sub

Private Sub VDescCustodia_Change()
'Me.vCodCustodia.Tag = Format(Me.VDescCustodia.Tag, "000")
Me.vCodCustodia.Text = Me.VDescCustodia.Tag
End Sub
