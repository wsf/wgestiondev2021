VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{9746E3DA-06E1-4D26-9CE4-D9F6411A9C70}#1.0#0"; "SMGA_OcxTxt2008.ocx"
Begin VB.Form frmAsientosAlta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formulario para Guardar Movimiento Contable"
   ClientHeight    =   8280
   ClientLeft      =   -45
   ClientTop       =   240
   ClientWidth     =   13245
   FillStyle       =   0  'Solid
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   13245
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   3375
      Left            =   90
      TabIndex        =   29
      Top             =   570
      Width           =   13065
      _Version        =   851968
      _ExtentX        =   23045
      _ExtentY        =   5953
      _StockProps     =   68
      ItemCount       =   2
      SelectedItem    =   1
      Item(0).Caption =   "Datos <F10>"
      Item(0).ControlCount=   20
      Item(0).Control(0)=   "cmdEventuales"
      Item(0).Control(1)=   "Picture3"
      Item(0).Control(2)=   "cmdProveedor"
      Item(0).Control(3)=   "cmdContribuyente"
      Item(0).Control(4)=   "cmd"
      Item(0).Control(5)=   "vcliprovee"
      Item(0).Control(6)=   "dtpFecha"
      Item(0).Control(7)=   "cboTipoMovimiento"
      Item(0).Control(8)=   "PBCambiarNumero"
      Item(0).Control(9)=   "txtNumero"
      Item(0).Control(10)=   "txtLeyenda"
      Item(0).Control(11)=   "lblAsientos(11)"
      Item(0).Control(12)=   "lblNroInterno"
      Item(0).Control(13)=   "lblAsientos(8)"
      Item(0).Control(14)=   "lblAsientos(7)"
      Item(0).Control(15)=   "lblAsientos(5)"
      Item(0).Control(16)=   "lblAsientos(6)"
      Item(0).Control(17)=   "lblAsientos(0)"
      Item(0).Control(18)=   "vrendicion"
      Item(0).Control(19)=   "PushButton1"
      Item(1).Caption =   "Linea de asiento <F12>"
      Item(1).ControlCount=   14
      Item(1).Control(0)=   "cboCodigo"
      Item(1).Control(1)=   "txtImporte(1)"
      Item(1).Control(2)=   "txtImporte(0)"
      Item(1).Control(3)=   "cboCuenta"
      Item(1).Control(4)=   "lblAsientos(4)"
      Item(1).Control(5)=   "lblAsientos(3)"
      Item(1).Control(6)=   "lblAsientos(1)"
      Item(1).Control(7)=   "lblCodigo"
      Item(1).Control(8)=   "lblImputable"
      Item(1).Control(9)=   "GroupBox1"
      Item(1).Control(10)=   "PbAcciones(2)"
      Item(1).Control(11)=   "GroupBox4"
      Item(1).Control(12)=   "Pus1"
      Item(1).Control(13)=   "PusCopiarImporte"
      Begin XtremeSuiteControls.PushButton PusCopiarImporte 
         Height          =   315
         Left            =   4800
         TabIndex        =   68
         Top             =   1500
         Width           =   3705
         _Version        =   851968
         _ExtentX        =   6535
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Copiar importe del total de la operación"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton Pus1 
         Height          =   315
         Left            =   4680
         TabIndex        =   67
         Top             =   570
         Width           =   345
         _Version        =   851968
         _ExtentX        =   609
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
         Picture         =   "frmAsientosAlta.frx":0000
      End
      Begin VB.TextBox vrendicion 
         Height          =   315
         Left            =   -68110
         TabIndex        =   65
         Top             =   2850
         Visible         =   0   'False
         Width           =   10935
      End
      Begin XtremeSuiteControls.GroupBox GroupBox4 
         Height          =   525
         Left            =   60
         TabIndex        =   61
         Top             =   2250
         Width           =   12825
         _Version        =   851968
         _ExtentX        =   22622
         _ExtentY        =   926
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
         Begin VB.TextBox vconcepto 
            Height          =   345
            Left            =   4590
            TabIndex        =   62
            Top             =   180
            Width           =   8205
         End
         Begin XtremeSuiteControls.PushButton PBAsientoTipo 
            Height          =   345
            Left            =   60
            TabIndex        =   63
            Top             =   180
            Width           =   1605
            _Version        =   851968
            _ExtentX        =   2831
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "A&siento Tipo"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   0
            Picture         =   "frmAsientosAlta.frx":059A
         End
         Begin XtremeSuiteControls.PushButton PusSelConceptos 
            Height          =   345
            Left            =   1680
            TabIndex        =   64
            Top             =   180
            Width           =   2865
            _Version        =   851968
            _ExtentX        =   5054
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "Sel. conceptos de aplicación <F4>"
            ForeColor       =   0
            BackColor       =   -2147483644
            UseVisualStyle  =   -1  'True
         End
      End
      Begin VB.TextBox txtLeyenda 
         Height          =   315
         Left            =   -68110
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   49
         Top             =   2250
         Visible         =   0   'False
         Width           =   10965
      End
      Begin VB.TextBox txtNumero 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -68110
         TabIndex        =   48
         Top             =   510
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox cboTipoMovimiento 
         Height          =   315
         Left            =   -68110
         TabIndex        =   46
         Top             =   1830
         Visible         =   0   'False
         Width           =   10965
      End
      Begin VB.TextBox vcliprovee 
         Height          =   315
         Left            =   -68110
         TabIndex        =   44
         Top             =   1380
         Visible         =   0   'False
         Width           =   4065
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Asistido"
         Height          =   315
         Left            =   -62380
         TabIndex        =   43
         Top             =   1410
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdContribuyente 
         Caption         =   "Contribuyente"
         Height          =   315
         Left            =   -61330
         TabIndex        =   42
         Top             =   1410
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdProveedor 
         Caption         =   "Proveedor"
         Height          =   315
         Left            =   -63430
         TabIndex        =   41
         Top             =   1410
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   -64000
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   40
         Top             =   1320
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton cmdEventuales 
         Caption         =   "Eventuales"
         Height          =   315
         Left            =   -60220
         TabIndex        =   39
         Top             =   1410
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox cboCuenta 
         Height          =   315
         Left            =   5130
         Style           =   1  'Simple Combo
         TabIndex        =   33
         Top             =   570
         Width           =   7845
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1830
         TabIndex        =   32
         Top             =   1020
         Width           =   2775
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1830
         TabIndex        =   31
         Top             =   1500
         Width           =   2775
      End
      Begin VB.ComboBox cboCodigo 
         Height          =   315
         Left            =   1830
         Style           =   1  'Simple Combo
         TabIndex        =   30
         Top             =   570
         Width           =   2775
      End
      Begin XtremeSuiteControls.Label lblCodigo 
         Height          =   315
         Left            =   6540
         TabIndex        =   37
         Top             =   390
         Width           =   4365
         _Version        =   851968
         _ExtentX        =   7699
         _ExtentY        =   556
         _StockProps     =   79
         ForeColor       =   4210752
         BackColor       =   -2147483644
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoEllipsis    =   -1  'True
         EnableMarkup    =   -1  'True
      End
      Begin Aplisoft_CajasDeTexto.TxF dtpFecha 
         Height          =   315
         Left            =   -68110
         TabIndex        =   45
         Top             =   960
         Visible         =   0   'False
         Width           =   2865
         _ExtentX        =   5054
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
         MinValor        =   40179
      End
      Begin XtremeSuiteControls.PushButton PBCambiarNumero 
         Height          =   315
         Left            =   -66340
         TabIndex        =   47
         Top             =   540
         Visible         =   0   'False
         Width           =   1035
         _Version        =   851968
         _ExtentX        =   1826
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Cambiar Nº"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   405
         Left            =   90
         TabIndex        =   57
         Top             =   1860
         Width           =   12795
         _Version        =   851968
         _ExtentX        =   22569
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Tipos de asientos:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
         Begin XtremeSuiteControls.RadioButton rdInterno 
            Height          =   195
            Left            =   1560
            TabIndex        =   58
            Top             =   180
            Width           =   1485
            _Version        =   851968
            _ExtentX        =   2619
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Asiento interno"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rdNormal 
            Height          =   255
            Left            =   3600
            TabIndex        =   59
            Top             =   150
            Width           =   1485
            _Version        =   851968
            _ExtentX        =   2619
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Asiento normal"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   435
         Index           =   2
         Left            =   60
         TabIndex        =   60
         Top             =   2940
         Width           =   12915
         _Version        =   851968
         _ExtentX        =   22781
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Cargar movimiento a la siguiente lista <F6> "
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   345
         Left            =   -69880
         TabIndex        =   66
         Top             =   2820
         Visible         =   0   'False
         Width           =   1635
         _Version        =   851968
         _ExtentX        =   2884
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Rendición  <F9>"
         ForeColor       =   0
         BackColor       =   -2147483644
         UseVisualStyle  =   -1  'True
         Picture         =   "frmAsientosAlta.frx":0B34
      End
      Begin VB.Label lblAsientos 
         Alignment       =   1  'Right Justify
         Caption         =   "Nro Asiento :"
         Height          =   195
         Index           =   0
         Left            =   -69070
         TabIndex        =   56
         Top             =   540
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label lblAsientos 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha :"
         Height          =   195
         Index           =   6
         Left            =   -68890
         TabIndex        =   55
         Top             =   990
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblAsientos 
         Alignment       =   1  'Right Justify
         Caption         =   "Leyenda del Asiento :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   -69730
         TabIndex        =   54
         Top             =   2310
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label lblAsientos 
         Alignment       =   1  'Right Justify
         Caption         =   "Nro Interno :"
         Height          =   195
         Index           =   7
         Left            =   -60520
         TabIndex        =   53
         Top             =   480
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label lblAsientos 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo de Comprobante :"
         Height          =   195
         Index           =   8
         Left            =   -69790
         TabIndex        =   52
         Top             =   1890
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label lblNroInterno 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   -58960
         TabIndex        =   51
         Top             =   420
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblAsientos 
         Alignment       =   1  'Right Justify
         Caption         =   "Persona:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   -69070
         TabIndex        =   50
         Top             =   1440
         Visible         =   0   'False
         Width           =   930
      End
      Begin XtremeSuiteControls.Label lblImputable 
         Height          =   285
         Left            =   6540
         TabIndex        =   38
         Top             =   690
         Width           =   4365
         _Version        =   851968
         _ExtentX        =   7699
         _ExtentY        =   503
         _StockProps     =   79
         ForeColor       =   255
         BackColor       =   -2147483644
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         AutoEllipsis    =   -1  'True
      End
      Begin VB.Label lblAsientos 
         Alignment       =   1  'Right Justify
         Caption         =   " Cuenta :"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   36
         Top             =   630
         Width           =   1575
      End
      Begin VB.Label lblAsientos 
         Alignment       =   1  'Right Justify
         Caption         =   "Importe del Debe :"
         Height          =   195
         Index           =   3
         Left            =   60
         TabIndex        =   35
         Top             =   1080
         Width           =   1650
      End
      Begin VB.Label lblAsientos 
         Alignment       =   1  'Right Justify
         Caption         =   "Importe del Haber:"
         Height          =   195
         Index           =   4
         Left            =   300
         TabIndex        =   34
         Top             =   1575
         Width           =   1395
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   165
      Left            =   0
      TabIndex        =   28
      Top             =   390
      Width           =   13215
      _Version        =   851968
      _ExtentX        =   23310
      _ExtentY        =   291
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   615
      Left            =   -30
      TabIndex        =   25
      Top             =   -90
      Width           =   13245
      _Version        =   851968
      _ExtentX        =   23363
      _ExtentY        =   1085
      _StockProps     =   79
      Appearance      =   2
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   3
         Left            =   90
         TabIndex        =   26
         Top             =   120
         Width           =   2265
         _Version        =   851968
         _ExtentX        =   3995
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Guardar Asiento <F2>"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PusCancelar 
         Height          =   345
         Index           =   4
         Left            =   11910
         TabIndex        =   27
         Top             =   120
         Width           =   1305
         _Version        =   851968
         _ExtentX        =   2302
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Cancelar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmAsientosAlta.frx":0F01
      End
   End
   Begin VB.TextBox vTxtnrobalance 
      Height          =   315
      Left            =   6330
      TabIndex        =   24
      Top             =   7260
      Width           =   2025
   End
   Begin XtremeSuiteControls.GroupBox GBModificar 
      Height          =   2595
      Left            =   270
      TabIndex        =   14
      Top             =   4080
      Visible         =   0   'False
      Width           =   12615
      _Version        =   851968
      _ExtentX        =   22251
      _ExtentY        =   4577
      _StockProps     =   79
      Caption         =   "Modificar"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.FlatEdit txtModificar 
         Height          =   315
         Index           =   0
         Left            =   1680
         TabIndex        =   15
         Top             =   270
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   0
         Left            =   3030
         TabIndex        =   17
         Tag             =   "CodigoCuenta"
         Top             =   270
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtModificar 
         Height          =   315
         Index           =   2
         Left            =   1680
         TabIndex        =   19
         Top             =   630
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtModificar 
         Height          =   315
         Index           =   3
         Left            =   1680
         TabIndex        =   20
         Top             =   990
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.PushButton cmdActualizarLinea 
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   2100
         Width           =   1395
         _Version        =   851968
         _ExtentX        =   2461
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Actualizar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmAsientosAlta.frx":149B
      End
      Begin XtremeSuiteControls.FlatEdit txtModificar 
         Height          =   315
         Index           =   1
         Left            =   3480
         TabIndex        =   18
         Top             =   270
         Width           =   4635
         _Version        =   851968
         _ExtentX        =   8176
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.Label lblModificar 
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   22
         Top             =   990
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Importe Haber:"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblModificar 
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   630
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Importe Debe:"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblModificar 
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   16
         Top             =   315
         Width           =   1500
         _Version        =   851968
         _ExtentX        =   2646
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Cuenta Contable:"
         Alignment       =   1
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.PushButton PbAcciones 
      Height          =   315
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   7200
      Width           =   1275
      _Version        =   851968
      _ExtentX        =   2249
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Borrar Linea"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PBAyuda 
      Height          =   420
      Left            =   11820
      TabIndex        =   9
      Top             =   6990
      Width           =   1245
      _Version        =   851968
      _ExtentX        =   2196
      _ExtentY        =   741
      _StockProps     =   79
      Caption         =   "&Ayuda"
      UseVisualStyle  =   -1  'True
      PushButtonStyle =   2
   End
   Begin XtremeSuiteControls.PushButton PBCuentas 
      Height          =   375
      Left            =   10860
      TabIndex        =   2
      Top             =   1140
      Width           =   2235
      _Version        =   851968
      _ExtentX        =   3942
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Ver Plan de Cuentas"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PbAcciones 
      Height          =   315
      Index           =   1
      Left            =   1470
      TabIndex        =   1
      Top             =   7200
      Width           =   1845
      _Version        =   851968
      _ExtentX        =   3254
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "&Limpiar Toda la grilla"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Frame fraVieneDe 
      Caption         =   "Viene de :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   120
      TabIndex        =   3
      Top             =   7590
      Width           =   13095
      Begin XtremeSuiteControls.CheckBox chkControlar 
         Height          =   210
         Left            =   11700
         TabIndex        =   13
         Top             =   270
         Width           =   1275
         _Version        =   851968
         _ExtentX        =   2249
         _ExtentY        =   379
         _StockProps     =   79
         Caption         =   "Controlar Total"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.TextBox txtImporteVieneDe 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   210
         Width           =   1875
      End
      Begin VB.TextBox txtCuentaVieneDe 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   210
         Width           =   3135
      End
      Begin VB.Label lblAsientos 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "> Importe : $"
         Height          =   195
         Index           =   10
         Left            =   8610
         TabIndex        =   12
         Top             =   270
         Width           =   885
      End
      Begin VB.Label lblAsientos 
         AutoSize        =   -1  'True
         Caption         =   "> Módulo :"
         Height          =   195
         Index           =   9
         Left            =   30
         TabIndex        =   11
         Top             =   255
         Width           =   750
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid FGAsientoDetalle 
      Height          =   2865
      Left            =   90
      TabIndex        =   5
      Top             =   3960
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   5054
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColorFixed  =   -2147483644
      ForeColorFixed  =   4210752
      BackColorSel    =   1375373
      ForeColorSel    =   255
      BackColorUnpopulated=   -2147483644
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      HighLight       =   2
      GridLines       =   2
      GridLinesUnpopulated=   1
      ScrollBars      =   2
      AllowUserResizing=   1
      GridLineWidthFixed=   2
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).TextStyleBand=   0
   End
   Begin XtremeSuiteControls.Label lblDisplay 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   6870
      Width           =   6015
      _Version        =   851968
      _ExtentX        =   10610
      _ExtentY        =   450
      _StockProps     =   79
      ForeColor       =   16744576
      BackColor       =   -2147483638
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      AutoEllipsis    =   -1  'True
      EnableMarkup    =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit FETotal 
      Height          =   285
      Index           =   0
      Left            =   6300
      TabIndex        =   7
      Top             =   6870
      Width           =   975
      _Version        =   851968
      _ExtentX        =   1720
      _ExtentY        =   503
      _StockProps     =   77
      ForeColor       =   16711680
      BackColor       =   -2147483643
      Alignment       =   1
   End
   Begin XtremeSuiteControls.FlatEdit FETotal 
      Height          =   285
      Index           =   1
      Left            =   7320
      TabIndex        =   8
      Top             =   6870
      Width           =   975
      _Version        =   851968
      _ExtentX        =   1720
      _ExtentY        =   503
      _StockProps     =   77
      ForeColor       =   255
      BackColor       =   -2147483643
      Alignment       =   1
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "Ayuda"
      Visible         =   0   'False
      Begin VB.Menu mnuAyuda01 
         Caption         =   "F1 ..... Ayuda"
      End
      Begin VB.Menu mnuAyuda02 
         Caption         =   "F2 ..... Guardar Movimiento"
      End
      Begin VB.Menu mnuAyuda03 
         Caption         =   "F3 ..... Guardar Asiento"
      End
      Begin VB.Menu mnuAyuda04 
         Caption         =   "F4 ..... Limpiar Campos"
      End
      Begin VB.Menu mnuAyuda05 
         Caption         =   "F5"
      End
      Begin VB.Menu mnuAyuda06 
         Caption         =   "F6 ..... Se posiciona en Campo Leyenda"
      End
      Begin VB.Menu mnuAyuda07 
         Caption         =   "F7 ..... Cambia Nro Interno"
      End
      Begin VB.Menu mnuAyuda08 
         Caption         =   "F8"
      End
      Begin VB.Menu mnuAyuda09 
         Caption         =   "F9 ..... Comprobante Anterior"
      End
      Begin VB.Menu mnuAyuda10 
         Caption         =   "F10 ... Comprobante Siguiente"
      End
   End
End
Attribute VB_Name = "frmAsientosAlta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vVieneTabla As String
Public vVieneIdNombre As String
Public vVieneIdValor As Long
Public vVieneNroInterno As Long
' ------------- esto es lo que hay que guardar en el registro de asientos ------
Public vnrointerno As Integer
Public vtipomovimiento, vCodigoProveedor, vCodigoCliente As String
Public vModificando As Boolean
Public vidpersonas, vidclientes, vidproveedores As Long
'------------------------------------------------------------------------------
Public vnrobalance As Integer


Public Sub ConfigurarGrilla()
On Error Resume Next

    With FGAsientoDetalle
       
        .Cols = 10
        .FixedRows = 1
        .FixedCols = 1
    
        .TextMatrix(0, 1) = "Fecha"
        .TextMatrix(0, 2) = "Nº Asiento"
        .TextMatrix(0, 3) = "C. Cuenta"
        .TextMatrix(0, 4) = "Cuenta"
        .TextMatrix(0, 5) = "Debe"
        .TextMatrix(0, 6) = "Haber"
    
        .ColWidth(0) = 600
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        .ColWidth(3) = 2000
        .ColWidth(4) = 5000
        .ColWidth(5) = 1500
        .ColWidth(6) = 1500
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        
        .ColAlignment(5) = 6
        .ColAlignment(6) = 6
       
        .Row = FGAsientoDetalle.Rows - 1
    
    End With

If Err Then GrabarLog "ConfigurarGrilla", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CalcularTotales()
On Error Resume Next

    Dim i As Integer

    FETotal(0).Text = 0
    FETotal(1).Text = 0
        
    With FGAsientoDetalle
        .Row = 1
        For i = 1 To .Rows - 1
            FETotal(0).Text = FETotal(0).Text + Val(.TextMatrix(i, 5))
            FETotal(1).Text = FETotal(1).Text + Val(.TextMatrix(i, 6))
        Next
    End With
                
If Err Then GrabarLog "CalcularTotales", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Grabar(Index As Integer)
On Error Resume Next

Dim vlogAsientosGuardado As String


    Dim i As Integer, vBancoCodigo As String, vBancoLeyenda As String
    
    Select Case Index
    
        Case 2
            Call CargarDetalleAsiento("", 0)
    
        Case 3
           'Aca guardo el Asiento
           ' Habilitar (False)
            
            lbldisplay.Caption = "Guardando Numero Interno"
            'Call GrabarNroInterno
            
            lbldisplay.Caption = "Guardando Asiento"
            Call GrabarAsiento ' graba los datos en la tabla Asientos
            
            lbldisplay.Caption = "Guardando Detalle"
            
         
            
                        With FGAsientoDetalle  ' guarda los datos en la tabla AsientosDetalle
                            .Row = 1
                            For i = 1 To .Rows - 1
                                    ' grabo el detalle del asiento en asientodetalle
                                    Call EjecutarScript("INSERT INTO AsientosDetalle (Numero, Linea, CodigoCuenta, Debe, Haber, nrobalance) VALUES (" & Val(txtNumero.Text) & "," & i & ", '" & .TextMatrix(i, 3) & "', " & Val(.TextMatrix(i, 5)) & ", " & Val(.TextMatrix(i, 6)) & "," + Str(vnrobalance) + ")")
                            Next
                        End With
    
           
    
          ' vlogAsientosGuardado = vlogAsientosGuardado + "Nro. Asiento: " + TraerDato("asientos", "Numero=" + txtNumero.Text, pathDBMySQL) + Chr(13)

           'vlogAsientosGuardado = vlogAsientosGuardado + "Nro. Asiento Detalle: " + TraerDato("asientosdetalle", "Numero=" + txtNumero.Text, pathDBMySQL)


            'MsgBox "El asiento fue guardado con los siguientes datos: " + Chr(13) + vlogAsientosGuardado


            
    End Select


'vlogAsientosGuardado = vlogAsientosGuardado + "Nro. Asiento: " + TraerDato("asiento", "numero=" + txtNumero.Text, pathDBMySQL) + Chr(13)

'vlogAsientosGuardado = vlogAsientosGuardado + "Nro. Asiento Detalle: " + TraerDato("asientodetalle", "numero=" + txtNumero.Text, pathDBMySQL)


'MsgBox "El asiento fue guardado con los siguientes datos: " + Chr(13) + vlogAsientosGuardado



If Err Then
    MsgBox Err.Description, vbCritical, "Error"
End If
End Sub


Public Sub setConceptos(vidconceptos As Integer, vimporte As Double)
Dim vr As Integer
On Error Resume Next
    
        Dim i As Integer, vCantidadDetalles As Integer
        Dim rsAsientosTipo As New ADODB.Recordset, sqlAsientosTipo As String
        
        sqlAsientosTipo = "SELECT * FROM conceptosctas where idconceptos=" + Str(vidconceptos)
        
        With rsAsientosTipo
            Call .Open(sqlAsientosTipo, ConnDDBB, adOpenStatic, adLockReadOnly)
            
            If Not .EOF = True Then
                .MoveFirst
                    
                  '  FGAsientoDetalle.Rows = FGAsientoDetalle.Rows + 1
                    vr = FGAsientoDetalle.Rows - 1
                    
                'If Val(.RecordCount) >= 0 Then
                'If Val(.RecordCount) >= Val(FGAsientoDetalle.Rows) Then
                   ' FGAsientoDetalle.Rows = FGAsientoDetalle.Rows + 1
                'End If
                    
                    FGAsientoDetalle.TextMatrix(vr, 1) = dtpFecha.Value
                    FGAsientoDetalle.TextMatrix(vr, 2) = txtNumero.Text
                    FGAsientoDetalle.TextMatrix(vr, 3) = EsNulo(.Fields("CodigoCuenta").Value)
                    FGAsientoDetalle.TextMatrix(vr, 4) = EsNulo(.Fields("Cuenta").Value)
                    
                    
                    If (.Fields("debe").Value) Then
                        FGAsientoDetalle.TextMatrix(vr, 5) = vimporte
                        FGAsientoDetalle.TextMatrix(vr, 6) = 0
                    Else
                        FGAsientoDetalle.TextMatrix(vr, 6) = vimporte
                        FGAsientoDetalle.TextMatrix(vr, 5) = 0
                    End If
                    
                    
                    vr = vr + 1
                    FGAsientoDetalle.Rows = FGAsientoDetalle.Rows + 1
                    
                    
                    FGAsientoDetalle.TextMatrix(vr, 1) = dtpFecha.Value
                    FGAsientoDetalle.TextMatrix(vr, 2) = txtNumero.Text
                    FGAsientoDetalle.TextMatrix(vr, 3) = EsNulo(.Fields("CodigoCuenta2").Value)
                    FGAsientoDetalle.TextMatrix(vr, 4) = EsNulo(.Fields("Cuenta2").Value)
                    
                    
                    If (.Fields("debe").Value) Then
                        FGAsientoDetalle.TextMatrix(vr, 6) = vimporte
                        FGAsientoDetalle.TextMatrix(vr, 5) = 0
                    Else
                        FGAsientoDetalle.TextMatrix(vr, 5) = vimporte
                        FGAsientoDetalle.TextMatrix(vr, 6) = 0
                    End If
                    
                  
            
            End If
            
            CalcularTotales
        End With
        
   ' End If
    
    sqlAsientosTipo = ""
    
    If rsAsientosTipo.State = 1 Then
        rsAsientosTipo.Close
        Set rsAsientosTipo = Nothing
    End If
    
If Err Then GrabarLog "CargarDetalleAsiento", Err.Number & " " & Err.Description, Me.Caption
End Sub



Public Sub CargarDetalleAsiento(vNroAsientoTipo As String, vImporteAsientoTipo As Double)
On Error Resume Next
    
    If Trim(vNroAsientoTipo) = "" Then
    
        With FGAsientoDetalle
            If Not .Rows = 2 Or Not (.TextMatrix(.Rows - 1, 2) = "") Then .AddItem (.Rows)
    
            '.TextMatrix(.Rows - 1, 1) = dtpFecha.Value
            .TextMatrix(.Rows - 1, 2) = txtNumero.Text
            .TextMatrix(.Rows - 1, 3) = cboCodigo.Text
            .TextMatrix(.Rows - 1, 4) = cboCuenta.Text
            .TextMatrix(.Rows - 1, 5) = Format(Val(txtImporte(0).Text), "#####0.00")
            .TextMatrix(.Rows - 1, 6) = Format(Val(txtImporte(1).Text), "#####0.00")
        End With
    
    Else
        Dim i As Integer, vCantidadDetalles As Integer
        Dim rsAsientosTipo As New ADODB.Recordset, sqlAsientosTipo As String
        
        sqlAsientosTipo = "SELECT idAsientosTipo, Numero, AsientosTipo.CodigoCuenta, Cuentas.Cuenta, DebeHaber, Porcentaje FROM AsientosTipo INNER JOIN Cuentas ON AsientosTipo.CodigoCuenta=Cuentas.CodigoCuenta WHERE (Numero = '" & vNroAsientoTipo & "') ORDER BY DebeHaber ASC"
        
        With rsAsientosTipo
            Call .Open(sqlAsientosTipo, ConnDDBB, adOpenStatic, adLockReadOnly)
            
            If Not .EOF = True Then
                .MoveFirst
            
                If Val(.RecordCount) >= Val(FGAsientoDetalle.Rows) Then
                    FGAsientoDetalle.Rows = .RecordCount + 1
                End If
                    
                Do Until .EOF = True
                    FGAsientoDetalle.TextMatrix(.AbsolutePosition, 1) = dtpFecha.Value
                    FGAsientoDetalle.TextMatrix(.AbsolutePosition, 2) = txtNumero.Text
                    FGAsientoDetalle.TextMatrix(.AbsolutePosition, 3) = EsNulo(.Fields("CodigoCuenta").Value)
                    FGAsientoDetalle.TextMatrix(.AbsolutePosition, 4) = EsNulo(.Fields("Cuenta").Value)
                    
                    If Not vImporteAsientoTipo = 0 Then
                        If EsNulo(.Fields("DebeHaber").Value) = "D" Then
                            FGAsientoDetalle.TextMatrix(.AbsolutePosition, 5) = Format(.Fields("Porcentaje").Value * Val(vImporteAsientoTipo) / 100, "######0.00")
                        Else
                            FGAsientoDetalle.TextMatrix(.AbsolutePosition, 6) = Format(.Fields("Porcentaje").Value * Val(vImporteAsientoTipo) / 100, "######0.00")
                        End If
                    Else
                        FGAsientoDetalle.TextMatrix(.AbsolutePosition, 5) = 0
                        FGAsientoDetalle.TextMatrix(.AbsolutePosition, 6) = 0
                    End If
                    
                    .MoveNext
                Loop
            
            End If
            
            CalcularTotales
        End With
        
    End If
    
    sqlAsientosTipo = ""
    
    If rsAsientosTipo.State = 1 Then
        rsAsientosTipo.Close
        Set rsAsientosTipo = Nothing
    End If
    
If Err Then GrabarLog "CargarDetalleAsiento", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub GrabarAsiento()
On Error Resume Next
    
    Dim rsAsientos As New ADODB.Recordset, sqlAsientos As String
    
    sqlAsientos = "SELECT * FROM Asientos WHERE Numero = " & Val(txtNumero.Text) & ""
    
    With rsAsientos
        .CursorLocation = adUseClient
        
        Call .Open(sqlAsientos, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If .EOF = True Then .AddNew

        .Fields("Fecha").Value = dtpFecha.Value
        .Fields("Numero").Value = Val(txtNumero.Text)
        .Fields("Leyenda").Value = Mid(Trim(txtLeyenda.Text), 1, Len(Trim(txtLeyenda.Text)))
        .Fields("TipoMovimiento").Value = Trim(cboTipoMovimiento.Tag)
        .Fields("NroBalance").Value = vnrobalance
        
        If Not Val(Me.lblNroInterno.Caption) > 0 Then Me.lblNroInterno.Caption = Str(UltimoNroInterno2 + 1)
        
        .Fields("NroInterno").Value = Val(lblNroInterno.Caption)
        
        .Fields("CodigoProveedor").Value = EsNulo(vCodigoProveedor)
        .Fields("CodigoCliente").Value = EsNulo(vCodigoCliente)
        
        
        If Not EsNulo(vCodigoCliente) = "" Then .Fields("cp").Value = "c"
        If Not EsNulo(vCodigoProveedor) = "" Then .Fields("cp").Value = "p"
        
        If Me.rdinterno.Value Then
            .Fields("marca").Value = "INTERNO" ' negro
        Else
            .Fields("marca").Value = "NORMAL" ' negro
        End If
        
            
'        Select Case Trim(txtCuentaVieneDe.Text)
'
'            Case "Documentos de Compras", "Cuentas Corrientes de Proveedores"
'                .Fields("cp").Value = "p"
'                .Fields("CodigoProveedor").Value = txtCuentaVieneDe.Tag
'
'            Case "Documentos de Ventas", "Cuenta Corrientes de Clientes"
'                .Fields("cp").Value = "c"
'                .Fields("CodigoCliente").Value = txtCuentaVieneDe.Tag
'
'        End Select

        .Fields("idrendiciones").Value = EsNulo(Me.vrendicion.Tag)

        .Update
        
        Dim vnroremito As Long
        
        If Not vVieneTabla = "" Then
            Call EjecutarScript("UPDATE " & vVieneTabla & " SET NroAsiento = " & .Fields("Numero").Value & " WHERE " & vVieneIdNombre & " = " & vVieneIdValor & "")
            
            Select Case vVieneTabla
                
                Case "Factura"
                    vnroremito = TraerDato("Factura", "idFactura = " & Val(vVieneIdValor) & "", "Remito")
                    
                    Call EjecutarScript("UPDATE CuentasCorrientes SET NroAsiento = " & .Fields("Numero").Value & " WHERE (Remito = " & vnroremito & ")")
                    
                Case "PFactura"
                    vnroremito = TraerDato("PFactura", "idPFactura = " & Val(vVieneIdValor) & "", "Remito")
                    
                    Call EjecutarScript("UPDATE CuentasCorrientes SET NroAsiento = " & .Fields("Numero").Value & " WHERE (Remito = " & vnroremito & ")")
                
                
            End Select
        
        End If
        
    End With
    
    sqlAsientos = ""
    vCodigoProveedor = ""
    vCodigoCliente = ""
    
    If rsAsientos.State = 1 Then
        rsAsientos.Close
        Set rsAsientos = Nothing
    End If

If Err < 0 Then
    'MsgBox Err.Description
'GrabarLog "GrabarAsiento", Err.Number & " " & Err.Description, Me.Name
End If

End Sub
Private Sub GrabarNroInterno()
On Error Resume Next

    Dim rsNroInterno As New ADODB.Recordset, sqlNroInterno As String
    
    sqlNroInterno = "SELECT NroInterno FROM NroInterno"
    
    With rsNroInterno
        .CursorLocation = adUseClient
        
        Call .Open(sqlNroInterno, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If .EOF = True Then .AddNew
    
        .Fields("NroInterno").Value = Val(lblNroInterno.Caption)
        
        .Update
    
    End With
    
    sqlNroInterno = ""
    
    If rsNroInterno.State = 1 Then
        rsNroInterno.Close
        Set rsNroInterno = Nothing
    End If

If Err Then GrabarLog "GrabarNroInterno", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Habilitar(vHabilita As Boolean)
On Error Resume Next

    With Me
        .PbAcciones(0).Enabled = vHabilita
        .PbAcciones(1).Enabled = vHabilita
        .PbAcciones(2).Enabled = vHabilita
        .PbAcciones(3).Enabled = vHabilita

        .PBCambiarNumero.Enabled = vHabilita
        .PBCuentas.Enabled = vHabilita

        .cboCodigo.Enabled = vHabilita
        .cboCuenta.Enabled = vHabilita
        
        .txtImporte(0).Enabled = vHabilita
        .txtImporte(1).Enabled = vHabilita
        
        '.txtNumero.Enabled = vHabilita
        .txtLeyenda.Enabled = vHabilita
        
        .dtpFecha.Enabled = vHabilita
        
        .FGAsientoDetalle.Enabled = vHabilita
        
        If vHabilita = True Then cboCodigo.SetFocus
        
    End With

If Err Then GrabarLog "Habilitar", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Limpiar(Index As Integer)
On Error Resume Next
    
    Me.lblNroInterno.Caption = ""
    Me.vconcepto = ""
    Me.vconcepto.Tag = ""
    cboCuenta.Text = ""
    cboCodigo.Text = ""
    txtImporte(0).Text = ""
    txtImporte(1).Text = ""
    lblImputable.Caption = ""
    lblCodigo.Caption = ""
    cboCuenta.Tag = ""
    cboCodigo.Tag = ""
        
    If (Index = 0) Or (Index = 1) Or (Index = 3) Then
        lbldisplay.Caption = "Espere ..."
        
        
        'txtNumero.Text = Val(GenerarDato("SELECT MAX(Numero) AS UAsiento FROM Asientos ", "UAsiento")) + 1
        txtNumero.Text = Val(GenerarDato("SELECT MAX(Numero) AS UAsiento FROM Asientos", "UAsiento")) + 1 ' los numeros absolutos
        
        
        dtpFecha.Value = Date
        
        FGAsientoDetalle.Clear
        FGAsientoDetalle.Rows = 2
    
        ConfigurarGrilla
        lbldisplay.Caption = ""
        
        If txtCuentaVieneDe.Tag = "" Then
            cboTipoMovimiento.Tag = ""
            txtLeyenda.Text = ""
            'lblNroInterno.Caption = Val(GenerarDato("SELECT * FROM NroInterno", "NroInterno")) + 1
    
            vVieneTabla = ""
            vVieneIdNombre = ""
            vVieneIdValor = 0
            txtCuentaVieneDe.Text = ""
            txtCuentaVieneDe.Tag = ""
            txtImporteVieneDe.Text = ""
            chkControlar.Value = xtpUnchecked
        
        End If
    End If
    
    CalcularTotales
    
    If cboCodigo.Enabled = True Then cboCodigo.SetFocus
    
   ' Me.TabControl1.SelectedItem = 0
    
   ' Me.txtLeyenda.SetFocus
    
If Err < 0 Then GrabarLog "Limpiar", Err.Number & " " & Err.Description, Me.Name
    End Sub
Private Function NumeroAsiento() As String
On Error Resume Next

    Dim rsNumeroAsiento As New ADODB.Recordset, sqlNumeroAsiento As String
    
    sqlNumeroAsiento = "SELECT MAX(Numero) as UNumero FROM AsientosDetalle GROUP BY Numero ORDER BY Numero ASC"
    
    With rsNumeroAsiento
        .CursorLocation = adUseClient
        
        Call .Open(sqlNumeroAsiento, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then
            .MoveLast
            NumeroAsiento = .Fields("UNumero").Value + 1
        Else
            NumeroAsiento = 1
        End If
    
    End With
    
    sqlNumeroAsiento = ""
    
    If rsNumeroAsiento.State = 1 Then
        rsNumeroAsiento.Close
        Set rsNumeroAsiento = Nothing
    End If
    
If Err Then GrabarLog "NumeroAsiento", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub PopupButtonMenu(btn As PushButton) ' sacado 060318
On Error Resume Next

    With btn
        If .Style = xtpButtonDropDownRight Then
            PopupMenu mnuAyuda, , .Left, .Top    ' - 500
        End If
    End With

If Err Then GrabarLog "PopupButtonMenu", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Function ValidarAsiento(Index As Integer) As Boolean
On Error Resume Next

    ValidarAsiento = True
    
    Select Case Index
    
        Case 2
            Dim i As Integer
            
            
           ' If Me.vrendicion.Tag = "" Then
           '     MsgBox "Debe Ingresar el campo de rendición", vbInformation
           '     ValidarAsiento = False
           '     Exit Function
           ' End If
            
            
           ' If Me.vrendicion.Tag = "" Then
           '     MsgBox "Debe Ingresar el campo de rendición", vbInformation
           '     ValidarAsiento = False
           '     Exit Function
           ' End If
            
            
            If (Trim(cboCodigo.Text) = "" Or Trim(cboCuenta.Text) = "") And Me.vconcepto.Tag = "" Then
                MsgBox "Debe Ingresar Codigo/Nombre de la Cuenta!!!"
                ValidarAsiento = False
                Exit Function
            End If
            
            If (Val(txtImporte(0).Text) = 0 And Val(txtImporte(1).Text) = 0) And Me.vconcepto.Tag = "" Then
                If Not MsgBox("Quiere guardar este asiento con valores en cero ?", vbYesNo, "Atención !") = vbYes Then
                ValidarAsiento = False
                txtImporte(0).SetFocus
                Exit Function
                End If
            End If
            
            If Val(txtImporte(0).Text) > 0 And Val(txtImporte(1).Text) > 0 Then
                MsgBox "No Puede Ingresar un Importe en DEBE y HABER al Mismo TIEMPO!!!!"
                ValidarAsiento = False
                txtImporte(0).SetFocus
                Exit Function
            End If
        
            If cboCodigo.Tag = "N" Then
                MsgBox "No Puede Ingresar un Movimiento con una Cuenta NO IMPUTABLE"
                ValidarAsiento = False
                cboCodigo.SetFocus
                Exit Function
            End If

        Case 3
            If dtpFecha.Value > Date Then
                ValidarAsiento = False
                Exit Function
            End If
            
            If (FGAsientoDetalle.Rows = 2) And (FGAsientoDetalle.TextMatrix(FGAsientoDetalle.Rows - 1, 3) = "") Then
                ValidarAsiento = False
                Exit Function
            End If
            
            If Val(txtNumero.Text) = 0 Then
                ValidarAsiento = False
                Exit Function
            End If

          
            If Not Val(FETotal(0).Text) = Val(FETotal(1).Text) Then
                ValidarAsiento = False
                MsgBox "El Asiento no se encuenta Balanceado!!", vbInformation, "Mensaje ..."
                Exit Function
            End If
            
        
            
           ' If chkControlar.Value = xtpChecked Then
           '     If Not (Val(txtImporteVieneDe.Text) = Val(FETotal(0).Text)) Then
           '         ValidarAsiento = False
           '         MsgBox "El Total del Asiento NO COINCIDE con el Total del Documento!!", vbInformation, "Mensaje ..."
           '         Exit Function
           '     End If
           ' End If
   End Select
    
If Err Then GrabarLog "ValidarAsiento", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub cboCodigo_Click()
On Error Resume Next

    cboCuenta.Text = TraerDato("Cuentas", "CodigoCuenta = '" & Trim(cboCodigo.Text) & "'", "Cuenta")
        
    If TraerDato("Cuentas", "Cuenta = '" & Trim(cboCuenta.Text) & "'", "Imputable") = "N" Then
        cboCodigo.Tag = "N"
        lblImputable.Caption = "NO IMPUTABLE"
        lblImputable.ForeColor = vbRed
    Else
        cboCodigo.Tag = "S"
        lblImputable.Caption = "IMPUTABLE"
        lblImputable.ForeColor = vbGreen
    End If
    
    lblCodigo.Caption = "" & MostrarCodigoCuenta(cboCodigo.Text)


If Err Then GrabarLog "cboCodigo_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cboCodigo_Change()
Me.txtImporte(0).SetFocus
End Sub

Private Sub cboCodigo_GotFocus()
On Error Resume Next

    Call CargarCombo("Cuentas", "CodigoCuenta", cboCodigo, True)
    
    

If Err Then GrabarLog "cboCodigo_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cboCodigo_KeyPress(KeyAscii As Integer)
On Error Resume Next
    
    If KeyAscii = 13 Then
    
        Dim rsCuentas As New ADODB.Recordset, sqlCuentas As String
    
    
        If Me.cboCodigo.Text = "" Then
            Call Pus1_Click
            frmConsultas.SetFocus
            frmConsultas.WindowState = 2
            Exit Sub
            
        End If
        
    
    
        sqlCuentas = "SELECT * FROM Cuentas WHERE (CodigoCuenta = '" & Trim(cboCodigo.Text) & "')"
    
        With rsCuentas
            .CursorLocation = adUseClient
        
            Call .Open(sqlCuentas, ConnDDBB, adOpenStatic, adLockReadOnly)
        
            If Not .EOF = True Then
                cboCuenta.Text = .Fields("Cuenta").Value
                
                If .Fields("Imputable") = "N" Then
                    cboCodigo.Tag = "N"
                    lblImputable.Caption = "NO IMPUTABLE"
                    lblImputable.ForeColor = vbRed
                Else
                    cboCodigo.Tag = "S"
                    lblImputable.Caption = "IMPUTABLE"
                    lblImputable.ForeColor = vbGreen
                End If
                
                lblCodigo.Caption = "" & MostrarCodigoCuenta(cboCodigo.Text)


                txtImporte(0).SetFocus
            Else
                cboCuenta.Text = ""
                cboCodigo.Text = ""
                cboCuenta.SetFocus
            End If
    
        End With

        sqlCuentas = ""
    
        If rsCuentas.State = 1 Then
            rsCuentas.Close
            Set rsCuentas = Nothing
        End If

    End If
    
If Err Then GrabarLog "txtCuenta_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cboCuenta_Click()
On Error Resume Next

    cboCodigo.Text = TraerDato("Cuentas", "Cuenta = '" & Trim(cboCuenta.Text) & "'", "CodigoCuenta")
    
    If TraerDato("Cuentas", "CodigoCuenta = '" & Trim(cboCodigo.Text) & "'", "Imputable") = "N" Then
        cboCodigo.Tag = "N"
        lblImputable.Caption = "NO IMPUTABLE"
        lblImputable.ForeColor = vbRed
    Else
        cboCodigo.Tag = "S"
        lblImputable.Caption = "IMPUTABLE"
        lblImputable.ForeColor = vbGreen
    End If

    lblCodigo.Caption = "" & MostrarCodigoCuenta(cboCodigo.Text)


If Err Then GrabarLog "cboCuenta_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cboCuenta_Change()
    Me.cboCodigo.Text = cboCuenta.Tag
End Sub

Private Sub cboCuenta_GotFocus()
On Error Resume Next

    Call CargarCombo("Cuentas", "Cuenta", cboCuenta, True)

If Err Then GrabarLog "cboCuenta_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cboCuenta_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then txtImporte(0).SetFocus

If Err Then GrabarLog "cboCuenta_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cboTipoMovimiento_Click()
On Error Resume Next

    cboTipoMovimiento.Tag = TraerDato("TipoMovimientos", "TipoMovimiento = '" & Trim(cboTipoMovimiento.Text) & "'", "Codigo")
    
If Err Then GrabarLog "cboTipoMovimiento_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cboTipoMovimiento_GotFocus()
On Error Resume Next

    Call CargarCombo("TipoMovimientos", "TipoMovimiento", cboTipoMovimiento, True)

If Err Then GrabarLog "cboTipoMovimiento_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmd_Click()
On Error Resume Next
Call fbuscarGrilla("clientes", "nombre", "codigo", Me.vcliprovee.Name, Me, "apellido", False)
vidproveedores = vcliprovee.Tag ' ema:
If Err Then Exit Sub
End Sub

Private Sub cmdActualizarLinea_Click()
On Error Resume Next
    
    If EsNulo(txtModificar(0).Text) = "" Then
        MsgBox "Debe Ingresar una cuenta", vbExclamation, "Mensaje ..."
        Exit Sub
    End If
    
    If Val(txtModificar(2).Text) = Val(txtModificar(3).Text) Then
        'MsgBox "Valor Debe/Haber Incorrecto", vbExclamation, "Mensaje ..."
        'Exit Sub
    End If
    
    If Val(txtModificar(2).Text) < 0 Then
        MsgBox "Valor Debe/Haber Incorrecto", vbExclamation, "Mensaje ..."
        Exit Sub
    End If
    
    If Val(txtModificar(3).Text) < 0 Then
        MsgBox "Valor Debe/Haber Incorrecto", vbExclamation, "Mensaje ..."
        Exit Sub
    End If
    
   
        
    With FGAsientoDetalle
        .TextMatrix(.Row, 3) = txtModificar(0).Text
        .TextMatrix(.Row, 4) = txtModificar(1).Text
        .TextMatrix(.Row, 5) = Format(txtModificar(2).Text, "##########0.00")
        .TextMatrix(.Row, 6) = Format(txtModificar(3).Text, "##########0.00")
    End With

    CalcularTotales
    
    GBModificar.Visible = False

If Err Then GrabarLog "cmdActualizarLinea_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub cmdContribuyente_Click()
Call fbuscarGrilla("personas", "nombre", "id_personas", Me.vcliprovee.Name, Me, "apellido", True)
vidproveedores = Val(vcliprovee.Tag) ' ema:
End Sub

Private Sub cmdProveedor_Click()
Call fbuscarGrilla("proveedores", "Nombre", "Codigo", Me.vcliprovee.Name, Me, , False)  ' ema:
vidproveedores = Val(vcliprovee.Tag)
End Sub

Private Sub dtpfecha_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
    If KeyCode = 13 Then cboCodigo.SetFocus

If Err Then GrabarLog "dtpfecha_KeyDown", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub FGAsientoDetalle_DblClick()
On Error Resume Next
    
    With GBModificar
        .Visible = True
        txtModificar(0).Text = FGAsientoDetalle.TextMatrix(FGAsientoDetalle.Row, 3)
        txtModificar(1).Text = FGAsientoDetalle.TextMatrix(FGAsientoDetalle.Row, 4)
        txtModificar(2).Text = FGAsientoDetalle.TextMatrix(FGAsientoDetalle.Row, 5)
        txtModificar(3).Text = FGAsientoDetalle.TextMatrix(FGAsientoDetalle.Row, 6)
        
        
        txtModificar(0).SetFocus
    End With
    
If Err Then GrabarLog "FGAsientoDetalle_DblClick", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub FGAsientoDetalle_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then

        With FGAsientoDetalle
            .TextMatrix(.Row, .Col) = .TextMatrix(.Row, .Col)
        End With

    End If
    
If Err Then GrabarLog "FGAsientoDetalle_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    Select Case KeyCode
    
    
        Case vbKeyF11
        Me.TabControl1.SelectedItem = 0
    
    
        Case vbKeyF12
        Me.TabControl1.SelectedItem = 1
        
        Case vbKeyF1
            

        Case vbKeyF2
            PbAcciones_Click (3)
        
        Case vbKeyF4
            Limpiar (0)
    
        Case vbKeyF5
        
        Case vbKeyF6
            Call agregarRenglonAsiento
        Case vbKeyF7
            PBAsientoTipo_Click
        
        Case vbKeyF8
            lblNroInterno_DblClick
            
        Case vbKeyF9
            With cboTipoMovimiento
                If Not .ListCount = 0 Then
                    If .ListIndex = 0 Then
                        .ListIndex = 0
                    Else
                        .ListIndex = .ListIndex - 1
                    End If
                Else
                    cboTipoMovimiento_GotFocus
                End If
            End With

        
        Case vbKeyF10
            With cboTipoMovimiento
                If Not .ListCount = 0 Then
                    If .ListIndex = .ListCount - 1 Then
                        .ListIndex = .ListCount - 1
                    Else
                        .ListIndex = .ListIndex + 1
                    End If
                Else
                    cboTipoMovimiento_GotFocus
                End If
            End With
    
        Case 27
            If MsgBox("Esta a punto de abandonar este formulario. " & vbCrLf & " Tambien se van a borrar los registros con el Numero Interno : " & lblNroInterno.Caption, vbYesNo + vbInformation, "Mensaje ...") = vbYes Then
                If Not Val(vVieneIdValor) = 0 Then
                    Call BorrarBase("CuentasCorrientes WHERE (NroInterno = " & Val(lblNroInterno.Caption) & ")", pathDBMySQL)
                    Call BorrarBase("Factura WHERE (NroInterno = " & Val(lblNroInterno.Caption) & ")", pathDBMySQL)
                    Call BorrarBase("PCuentasCorrientes WHERE (NroInterno = " & Val(lblNroInterno.Caption) & ")", pathDBMySQL)
                    Call BorrarBase("PFactura WHERE (NroInterno = " & Val(lblNroInterno.Caption) & ")", pathDBMySQL)
                    Call BorrarBase("BancosMovimientos WHERE (NroInterno = " & Val(lblNroInterno.Caption) & ")", pathDBMySQL)
                    Call BorrarBase("Asientos WHERE (NroInterno = " & Val(lblNroInterno.Caption) & ")", pathDBMySQL)
                    Call BorrarBase("AsientosDetalle WHERE (NroInterno = " & Val(lblNroInterno.Caption) & ")", pathDBMySQL)
                End If
                
                Unload Me
            End If
            
    End Select
    
    
    If LeerXml("") Then frmAsientosAlta.Show

If Err Then GrabarLog "Form_KeyUp", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Load()
On Error Resume Next
    
    KeyPreview = True
    
    Limpiar (0)
    
    Call CentrarFormulario(Me)
             
    init
    
If Err < 0 Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name

End Sub
Private Sub init()
    vnrobalance = TraerDato("balances", " Activo='S' order by NroBalance Desc", "NroBalance", pathDBMySQL)
    Me.Caption = "Carga de movimientos de Asientos Contables. " + " [Nro. de Balance: " + Str(vnrobalance) + "]"
 
    TabControl1.SelectedItem = 0
    
    
If LeerXml("Puesto") = "comuna" And Not Me.vVieneTabla = "" Then
    Me.txtImporte(0).Enabled = False
    Me.txtImporte(0).Visible = False
End If
    
    
    txtLeyenda.SetFocus
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub lblNroInterno_DblClick()
On Error Resume Next
    
    Dim vnrointerno As Long

    vnrointerno = InputBox("", "Mensaje ...", Val(lblNroInterno.Caption))
    
    If Val(vnrointerno) > 0 Then
        lblNroInterno.Caption = vnrointerno
    Else
        MsgBox "Debe ingresar un numero Interno VALIDO!!!", vbExclamation, "Mensaje ..."
    End If

If Err Then GrabarLog "lblNroInterno_DblClick", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub PbAcciones_Click(Index As Integer)
On Error Resume Next

Dim vimporte As Double

vnrobalance = selectNrobalance(dtpFecha.Value, dtpFecha, vnrobalance)


If Not vModificando Then
    If existeRegistroAsientos(Val(lblNroInterno)) Then Exit Sub
End If


'If Val(lblNroInterno) = 0 And vConfigGral.vIncluyeContabilidad Then
 '   MsgBox "Debe ingresar un nro interno válido", vbInformation, "Cuidado..."
'    Exit Sub
'End If





    Select Case Index
    
        Case 0
            If FGAsientoDetalle.Rows = 2 Then
                FGAsientoDetalle.Clear
                ConfigurarGrilla
            Else
                Call FGAsientoDetalle.RemoveItem(FGAsientoDetalle.RowSel)
            End If

            CalcularTotales
            
        Case 1
            Call Limpiar(Index)


        Case 2, 3
        
        
            If Not Me.vconcepto = "" And Index = 2 Then
                 
                If Val(Me.txtImporteVieneDe.Text) = 0 Then
                       vimporte = Val(InputBox("Ingrese importe de la operación", ""))
                Else
                       vimporte = txtImporteVieneDe.Text
                End If
                
                If vimporte = 0 Then Exit Sub
                
                Call setConceptos(vconcepto.Tag, vimporte)
         
           End If
                
           If Index = 3 And vModificando Then
                'If Not validarTransaccion(Me.lblNroInterno.Caption, vConfigGral.vIdUsuario) Then Exit Sub
            End If
        
        
            If vModificando Then
                borrarAsientoAmodificar (Val(Me.txtNumero))
                Call dTransaccion(Me.lblNroInterno)
            End If
            
            
            If ValidarAsiento(Index) Then
                Grabar (Index)
                Limpiar (Index)
                
                If Index = 3 Then
                
                
                                If vcliprovee.Text = "" Then vcliprovee.Text = Me.vCodigoCliente + Me.vCodigoProveedor
        
        
                             '   If Trim(vcliprovee.Text) = "" Then
                             '           If MsgBox("No hay Cliente / Proveedor para cargar." + Chr(13) + " Continúa grabando ?", vbYesNo) = vbNo Then
                             '               Exit Sub
                             '           End If
                             '   End If
             
                            'frmAsientos.PbAcciones_Click (4)
                            
                    Call wTransaccion(Me.lblNroInterno, vConfigGral.vIdUsuario)
                   ' Unload frmAsientosAlta
                   
                 Me.TabControl1.SelectedItem = 0
    
                 Me.txtLeyenda.SetFocus
                    
                Else
                    Habilitar (True)
                End If
            
        End If

        
End Select



If Err < 0 Then
    MsgBox Err.Description, vbCritical, "Error"
    GrabarLog "PBAcciones_Click", Err.Number & " " & Err.Description, Me.Name
End If
End Sub

Private Sub agregarRenglonAsiento()
On Error Resume Next
Dim vimporte As Double
        
            If Not Me.vconcepto = "" Then
                 
                If Val(Me.txtImporteVieneDe.Text) = 0 Then
                       vimporte = Val(InputBox("Ingrese importe de la operación", ""))
                Else
                       vimporte = txtImporteVieneDe.Text
                End If
                
                If vimporte = 0 Then Exit Sub
                
                Call setConceptos(vconcepto.Tag, vimporte)
         
           End If
                
        
        
            If vModificando Then
                borrarAsientoAmodificar (Val(Me.txtNumero))
                Call dTransaccion(Me.lblNroInterno)
            End If
            
            
            If ValidarAsiento(2) Then
                Grabar (2)
                Limpiar (2)
                
                    Habilitar (True)
            
        End If

If Err Then Exit Sub
End Sub




Private Sub borrarAsientoAmodificar(vnroasiento As Long)
Dim vsql As String

vsql = "delete from asientos where numero=" + Str(vnroasiento) + " and nrobalance=" + Str(vnrobalance)
Call EjecutarScript(vsql, pathDBMySQL)


vsql = "delete from asientosdetalle where numero=" + Str(vnroasiento) + " and nrobalance=" + Str(vnrobalance)
Call EjecutarScript(vsql, pathDBMySQL)

End Sub
Private Sub PBAsientoTipo_Click()
On Error Resume Next

    frmAsientosTipo.Show
    frmAsientosTipo.txtAsiento(1).Text = EsNulo(txtImporteVieneDe.Text)

If Err Then GrabarLog "PBAsientoTipo_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub PBAyuda_DropDown()
On Error Resume Next

    Call PopupButtonMenu(PBAyuda)

If Err Then GrabarLog "PBAyuda_DropDown", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub PBCambiarNumero_Click()
On Error Resume Next

    With txtNumero
        .Text = ""
        .Enabled = True
        .SetFocus
    End With

If Err Then GrabarLog "PBCambiarNumero_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub pbCarga_Click(Index As Integer)
    On Error Resume Next

    vVuelveBusqueda = Me.Name
    vVieneBusqueda = pbCarga(Index).Tag

    Select Case Index
    
        Case 0 To pbCarga.Count - 1
            frmBusqueda.Show

    End Select

If Err Then GrabarLog "pbCarga_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub PBCuentas_Click()
On Error Resume Next

    frmCuentas.Show

If Err Then GrabarLog "PBCuentas_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub Pus1_Click()
    Call fbuscarGrilla("(select * from cuentas where Imputable ='S') as t", "Cuenta", "CodigoCuenta", Me.cboCuenta.Name, Me) ' ema:
    Me.txtImporte(0).SetFocus
    frmConsultas.SetFocus
    frmConsultas.WindowState = 2
    frmConsultas.vbuscando.SetFocus
    
End Sub

Private Sub PusCancelar_Click(Index As Integer)
Unload Me
End Sub

Private Sub PusCopiarImporte_Click()
    Me.txtImporte(1).Text = Me.txtImporteVieneDe.Text
End Sub

Private Sub PushButton1_Click()
Call fbuscarGrilla("rendiciones", "nombre", "idrendiciones", Me.vrendicion.Name, Me)    ' ema:
End Sub

Private Sub PusSelConceptos_Click()
Call fbuscarGrilla("conceptos2", "descripcion", "idconceptos", Me.vconcepto.Name, Me)   ' ema:
End Sub

Private Sub txtImporte_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        If Index = 0 Then
            txtImporte(1).SetFocus
        Else
            If Trim(txtLeyenda.Text) = "" Then
                PbAcciones(2).SetFocus
                
            Else: PbAcciones(2).SetFocus
                PbAcciones(2).SetFocus
            End If
        End If
    
    End If
    
If Err Then GrabarLog "txtLeyenda_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtLeyenda_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        'PbAcciones(2).SetFocus
        txtLeyenda.Text = UCase(Me.txtLeyenda.Text)
        Me.TabControl1.SelectedItem = 1
    End If


If Err Then GrabarLog "txtLeyenda_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub txtModificar_GotFocus(Index As Integer)
On Error Resume Next

    With txtModificar(Index)
        .SelStart = 0
        .SelLength = Len(txtModificar(Index).Text)
    End With

If Err Then GrabarLog "txtModificar_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtModificar_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        Select Case Index
        
            Case 0
                txtModificar(1).Text = TraerDato("Cuentas", "CodigoCuenta = '" & Trim(txtModificar(0).Text) & "'", "Cuenta")
                
                If txtModificar(1).Text = "" Then
                
                Else
                    txtModificar(2).SetFocus
                End If
            
            Case 1
                txtModificar(Index + 1).SetFocus
            Case 2
                txtModificar(Index + 1).SetFocus
            Case 3
                cmdActualizarLinea.SetFocus
        End Select
        
    End If

If Err Then GrabarLog "", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then cboCodigo.SetFocus

If Err Then GrabarLog "txtNumero_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtNumero_LostFocus()
On Error Resume Next

    txtNumero.Enabled = False

   

If Err Then GrabarLog "txtNumero_LostFocus", Err.Number & " " & Err.Description, Me.Name
End Sub

