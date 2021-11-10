VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{63BEADB1-20E1-478A-9B40-DDDAFBF3624F}#1.0#0"; "bsGradientLabel.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "Copia de KlexGrid.ocx"
Object = "{9746E3DA-06E1-4D26-9CE4-D9F6411A9C70}#1.0#0"; "SMGA_OcxTxt2008.ocx"
Begin VB.Form frmIngresosEgresos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario de Ingresos y Egresos de Caja y Banco"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   16320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   16320
   Begin XtremeSuiteControls.FlatEdit vnro_doc_ale 
      Height          =   285
      Left            =   13185
      TabIndex        =   183
      Top             =   45
      Width           =   915
      _Version        =   851968
      _ExtentX        =   1614
      _ExtentY        =   503
      _StockProps     =   77
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   8835
      Left            =   14280
      TabIndex        =   151
      Top             =   120
      Width           =   1935
      _Version        =   851968
      _ExtentX        =   3413
      _ExtentY        =   15584
      _StockProps     =   68
      Appearance      =   8
      Color           =   128
      ItemCount       =   10
      SelectedItem    =   4
      Item(0).Caption =   "10 - Cambiar Valores"
      Item(0).ControlCount=   7
      Item(0).Control(0)=   "Option1"
      Item(0).Control(1)=   "Option2"
      Item(0).Control(2)=   "Option3"
      Item(0).Control(3)=   "Option4"
      Item(0).Control(4)=   "Option11"
      Item(0).Control(5)=   "frame_doc"
      Item(0).Control(6)=   "f2"
      Item(1).Caption =   "11 - Depositar Valores"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "Option5"
      Item(1).Control(1)=   "Option6"
      Item(1).Control(2)=   "a"
      Item(2).Caption =   "12 - Acreditar Valores"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "Option7"
      Item(2).Control(1)=   "Option8"
      Item(3).Caption =   "13 - Val.Rechazados"
      Item(3).ControlCount=   2
      Item(3).Control(0)=   "Option9"
      Item(3).Control(1)=   "Option10"
      Item(4).Caption =   "22 - Cargar Servicios"
      Item(4).ControlCount=   3
      Item(4).Control(0)=   "Option12"
      Item(4).Control(1)=   "Option13"
      Item(4).Control(2)=   "Log22"
      Item(5).Caption =   "Item"
      Item(5).ControlCount=   0
      Item(6).Caption =   "Item"
      Item(6).ControlCount=   0
      Item(7).Caption =   "Item"
      Item(7).ControlCount=   0
      Item(8).Caption =   "Item"
      Item(8).ControlCount=   0
      Item(9).Caption =   "Item"
      Item(9).ControlCount=   0
      Begin XtremeSuiteControls.GroupBox f2 
         Height          =   1665
         Left            =   -69940
         TabIndex        =   171
         Top             =   1080
         Visible         =   0   'False
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   2937
         _StockProps     =   79
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
         Begin VB.TextBox vcomiPorc 
            Height          =   345
            Left            =   60
            TabIndex        =   174
            Top             =   360
            Width           =   1665
         End
         Begin VB.TextBox vcomiFijo 
            Height          =   345
            Left            =   60
            TabIndex        =   173
            Top             =   960
            Width           =   1665
         End
         Begin XtremeSuiteControls.PushButton PusAceptar 
            Height          =   225
            Left            =   90
            TabIndex        =   172
            Top             =   1350
            Width           =   1605
            _Version        =   851968
            _ExtentX        =   2831
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "Aceptar"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label Label7 
            Caption         =   "% Comisión:"
            Height          =   345
            Left            =   150
            TabIndex        =   176
            Top             =   180
            Width           =   1515
         End
         Begin VB.Label Label9 
            Caption         =   "Importe Fijo:"
            Height          =   345
            Left            =   60
            TabIndex        =   175
            Top             =   780
            Width           =   1545
         End
      End
      Begin XtremeSuiteControls.ListBox Log22 
         Height          =   2385
         Left            =   120
         TabIndex        =   170
         Top             =   2340
         Width           =   1725
         _Version        =   851968
         _ExtentX        =   3043
         _ExtentY        =   4207
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin VB.OptionButton Option13 
         Caption         =   "2- Generar archivo."
         Height          =   345
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   169
         Top             =   4980
         Width           =   1665
      End
      Begin VB.OptionButton Option12 
         Caption         =   "1- Comenzar carga por Codigo de Barra"
         Height          =   495
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   168
         Top             =   1620
         Width           =   1575
      End
      Begin XtremeSuiteControls.GroupBox frame_doc 
         Height          =   2535
         Left            =   -69940
         TabIndex        =   164
         Top             =   3780
         Visible         =   0   'False
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   4471
         _StockProps     =   79
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
         Begin MSComCtl2.DTPicker txtvencimiento 
            Height          =   315
            Left            =   90
            TabIndex        =   181
            Top             =   1440
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   556
            _Version        =   393216
            Format          =   208142337
            CurrentDate     =   42828
         End
         Begin VB.TextBox txtvimporte_pagare 
            Height          =   345
            Left            =   60
            TabIndex        =   179
            Top             =   810
            Width           =   1665
         End
         Begin VB.TextBox txtvintereses_pagare 
            Height          =   315
            Left            =   60
            TabIndex        =   177
            Top             =   300
            Width           =   1635
         End
         Begin XtremeSuiteControls.PushButton PushButton16 
            Height          =   315
            Left            =   90
            TabIndex        =   165
            Top             =   2190
            Width           =   1665
            _Version        =   851968
            _ExtentX        =   2937
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Recibo"
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
            Picture         =   "frmIngresosEgresos.frx":0000
            BorderGap       =   10
         End
         Begin XtremeSuiteControls.PushButton PushButton15 
            Height          =   315
            Left            =   120
            TabIndex        =   166
            Top             =   1830
            Width           =   1665
            _Version        =   851968
            _ExtentX        =   2937
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Pagaré"
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
            Picture         =   "frmIngresosEgresos.frx":059A
            BorderGap       =   10
         End
         Begin VB.Label lblFechaVencimiento 
            Caption         =   "Fecha Vencim. Pagaré"
            Height          =   255
            Left            =   60
            TabIndex        =   182
            Top             =   1170
            Width           =   1665
         End
         Begin VB.Label Label11 
            Caption         =   "Total del  Pagaré:"
            Height          =   255
            Left            =   90
            TabIndex        =   180
            Top             =   600
            Width           =   1545
         End
         Begin VB.Label Label10 
            Caption         =   "Intereses:"
            Height          =   345
            Left            =   30
            TabIndex        =   178
            Top             =   120
            Width           =   1545
         End
      End
      Begin VB.OptionButton Option11 
         Caption         =   "Imprimir Docum."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -69910
         Style           =   1  'Graphical
         TabIndex        =   163
         Top             =   3450
         Visible         =   0   'False
         Width           =   1785
      End
      Begin RichTextLib.RichTextBox a 
         Height          =   1215
         Left            =   -69940
         TabIndex        =   162
         Top             =   690
         Visible         =   0   'False
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   2143
         _Version        =   393217
         BackColor       =   12648447
         BorderStyle     =   0
         Appearance      =   0
         TextRTF         =   $"frmIngresosEgresos.frx":0B34
      End
      Begin VB.OptionButton Option10 
         Caption         =   "Gastos del Depósito"
         Height          =   405
         Left            =   -69940
         Style           =   1  'Graphical
         TabIndex        =   161
         Top             =   1770
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Pasar Valor a Rechazado"
         Height          =   525
         Left            =   -69940
         Style           =   1  'Graphical
         TabIndex        =   160
         Top             =   1170
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Gastos del Depósito"
         Height          =   375
         Left            =   -69940
         Style           =   1  'Graphical
         TabIndex        =   159
         Top             =   1710
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Cobrar Valores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69940
         Style           =   1  'Graphical
         TabIndex        =   158
         Top             =   900
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Gastos del Depósito"
         Height          =   375
         Left            =   -69940
         Style           =   1  'Graphical
         TabIndex        =   157
         Top             =   2400
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Despositar Valor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69940
         Style           =   1  'Graphical
         TabIndex        =   156
         Top             =   1980
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Agregar Gastos Extras"
         Height          =   315
         Left            =   -69910
         Style           =   1  'Graphical
         TabIndex        =   155
         Top             =   3120
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Pagaler Valor a Cli."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -69910
         Style           =   1  'Graphical
         TabIndex        =   154
         Top             =   2790
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Fijar Comisión"
         Height          =   345
         Left            =   -69940
         Style           =   1  'Graphical
         TabIndex        =   153
         Top             =   720
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ingresar Valor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -69940
         Style           =   1  'Graphical
         TabIndex        =   152
         Top             =   330
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox4 
      Height          =   195
      Left            =   0
      TabIndex        =   36
      Top             =   510
      Width           =   13155
      _Version        =   851968
      _ExtentX        =   23204
      _ExtentY        =   344
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
   End
   Begin XtremeSuiteControls.GroupBox g3 
      Height          =   405
      Left            =   30
      TabIndex        =   30
      Top             =   0
      Width           =   10935
      _Version        =   851968
      _ExtentX        =   19288
      _ExtentY        =   714
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton cmdGuardar 
         Height          =   315
         Left            =   150
         TabIndex        =   11
         Top             =   60
         Width           =   1635
         _Version        =   851968
         _ExtentX        =   2884
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Guardar <F2>"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmIngresosEgresos.frx":0C16
         BorderGap       =   10
      End
      Begin XtremeSuiteControls.PushButton cmdCerrar 
         Height          =   315
         Left            =   5400
         TabIndex        =   31
         Top             =   60
         Visible         =   0   'False
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Cerrar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmIngresosEgresos.frx":11B0
      End
      Begin XtremeSuiteControls.PushButton PusGuardarSin 
         Height          =   315
         Left            =   1920
         TabIndex        =   32
         Top             =   60
         Visible         =   0   'False
         Width           =   3285
         _Version        =   851968
         _ExtentX        =   5794
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Guardar sin asiento atomático  asiento"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmIngresosEgresos.frx":174A
         BorderGap       =   10
      End
      Begin XtremeSuiteControls.PushButton PusImprimirGuardar 
         Height          =   315
         Left            =   6690
         TabIndex        =   47
         Top             =   60
         Width           =   2295
         _Version        =   851968
         _ExtentX        =   4048
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Imprimir (1 copia) <F10>"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmIngresosEgresos.frx":1B2E
         BorderGap       =   10
      End
      Begin XtremeSuiteControls.PushButton PushButton11 
         Height          =   315
         Left            =   9030
         TabIndex        =   122
         Top             =   60
         Width           =   1845
         _Version        =   851968
         _ExtentX        =   3254
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Imprimir (2 copias)"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmIngresosEgresos.frx":20C8
         BorderGap       =   10
      End
   End
   Begin XtremeSuiteControls.TabControl TabAlta 
      Height          =   8355
      Left            =   0
      TabIndex        =   12
      Top             =   660
      Width           =   14205
      _Version        =   851968
      _ExtentX        =   25056
      _ExtentY        =   14737
      _StockProps     =   68
      PaintManager.BoldSelected=   -1  'True
      ItemCount       =   6
      Item(0).Caption =   "Ingresos-Egresos"
      Item(0).ControlCount=   42
      Item(0).Control(0)=   "KlexMovimientoCaja"
      Item(0).Control(1)=   "GBRBSuperior"
      Item(0).Control(2)=   "cmdCheque"
      Item(0).Control(3)=   "GroupBox1"
      Item(0).Control(4)=   "Frame1"
      Item(0).Control(5)=   "gcustodia"
      Item(0).Control(6)=   "VchequesDisplay"
      Item(0).Control(7)=   "cmdContribuyente"
      Item(0).Control(8)=   "cmd"
      Item(0).Control(9)=   "lblAsientos(11)"
      Item(0).Control(10)=   "vcliprovee"
      Item(0).Control(11)=   "lblIngresarUna"
      Item(0).Control(12)=   "vobservacion"
      Item(0).Control(13)=   "cmdEventuales"
      Item(0).Control(14)=   "tab2"
      Item(0).Control(15)=   "vconcepto"
      Item(0).Control(16)=   "PusSelConceptos"
      Item(0).Control(17)=   "f1"
      Item(0).Control(18)=   "Command1"
      Item(0).Control(19)=   "txtAlta(12)"
      Item(0).Control(20)=   "lblAltaCaja(10)"
      Item(0).Control(21)=   "vrendicion"
      Item(0).Control(22)=   "PushButton4"
      Item(0).Control(23)=   "txtAlta(13)"
      Item(0).Control(24)=   "lblAltaCaja(11)"
      Item(0).Control(25)=   "txtAlta(3)"
      Item(0).Control(26)=   "pbCarga(1)"
      Item(0).Control(27)=   "txtAlta(4)"
      Item(0).Control(28)=   "lblAltaCaja(4)"
      Item(0).Control(29)=   "PusPersonas"
      Item(0).Control(30)=   "vobservacion2"
      Item(0).Control(31)=   "Label4"
      Item(0).Control(32)=   "lblCta"
      Item(0).Control(33)=   "lblCtaSeleccionada"
      Item(0).Control(34)=   "PushButton12"
      Item(0).Control(35)=   "PushButton13"
      Item(0).Control(36)=   "PusBuscarDocumento"
      Item(0).Control(37)=   "PusLimpiar"
      Item(0).Control(38)=   "lbsaldo"
      Item(0).Control(39)=   "gsaldos"
      Item(0).Control(40)=   "PusActSaldo"
      Item(0).Control(41)=   "cmdCerrar2"
      Item(1).Caption =   "Últimos movimientos de CAJA"
      Item(1).ControlCount=   10
      Item(1).Control(0)=   "gultimos"
      Item(1).Control(1)=   "vfiltro"
      Item(1).Control(2)=   "lblBuscarPor"
      Item(1).Control(3)=   "PusImprimir"
      Item(1).Control(4)=   "PushButton9"
      Item(1).Control(5)=   "PusSaldosPor"
      Item(1).Control(6)=   "PushButton10"
      Item(1).Control(7)=   "PusBorrarMovimientos"
      Item(1).Control(8)=   "PushButton14"
      Item(1).Control(9)=   "gdetalle"
      Item(2).Caption =   "CONTABILIDAD"
      Item(2).ControlCount=   4
      Item(2).Control(0)=   "PushButton5"
      Item(2).Control(1)=   "FlatEdit1"
      Item(2).Control(2)=   "MSHFlexGrid1"
      Item(2).Control(3)=   "Label5"
      Item(3).Caption =   "VALORES"
      Item(3).ControlCount=   4
      Item(3).Control(0)=   "PushButton6"
      Item(3).Control(1)=   "FlatEdit2"
      Item(3).Control(2)=   "MSHFlexGrid2"
      Item(3).Control(3)=   "Label6"
      Item(4).Caption =   "VALES"
      Item(4).ControlCount=   2
      Item(4).Control(0)=   "MSHFlexGrid3"
      Item(4).Control(1)=   "MSHFlexGrid4"
      Item(5).Caption =   "Opciones"
      Item(5).ControlCount=   4
      Item(5).Control(0)=   "PushButton8"
      Item(5).Control(1)=   "PusCierreDe"
      Item(5).Control(2)=   "Label8"
      Item(5).Control(3)=   "PushButton7"
      Begin VB.CommandButton cmdCerrar2 
         Caption         =   "Cerrar"
         Height          =   525
         Left            =   13530
         TabIndex        =   148
         Top             =   7710
         Width           =   615
      End
      Begin XtremeSuiteControls.PushButton PusActSaldo 
         Height          =   240
         Left            =   9390
         TabIndex        =   147
         Top             =   390
         Width           =   1275
         _Version        =   851968
         _ExtentX        =   2249
         _ExtentY        =   423
         _StockProps     =   79
         Caption         =   "Actualizar Saldo"
         UseVisualStyle  =   -1  'True
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gsaldos 
         Height          =   2595
         Left            =   10890
         TabIndex        =   146
         Top             =   330
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   4577
         _Version        =   393216
         BackColor       =   6513507
         ForeColor       =   15591427
         FixedCols       =   0
         BackColorSel    =   65280
         ForeColorSel    =   255
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         GridLineWidthFixed=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin XtremeSuiteControls.PushButton PusLimpiar 
         Height          =   225
         Left            =   1380
         TabIndex        =   140
         Top             =   2160
         Width           =   615
         _Version        =   851968
         _ExtentX        =   1085
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Limpiar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PusBuscarDocumento 
         Height          =   345
         Left            =   8580
         TabIndex        =   137
         Top             =   2070
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Sel. Doc."
         UseVisualStyle  =   -1  'True
         TextAlignment   =   5
         Picture         =   "frmIngresosEgresos.frx":2662
      End
      Begin XtremeSuiteControls.PushButton PusCierreDe 
         Height          =   405
         Left            =   -69820
         TabIndex        =   125
         Top             =   1260
         Visible         =   0   'False
         Width           =   2055
         _Version        =   851968
         _ExtentX        =   3625
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Ver Saldos Disponibles"
         UseVisualStyle  =   -1  'True
         PushButtonStyle =   2
      End
      Begin XtremeSuiteControls.PushButton PusBorrarMovimientos 
         Height          =   315
         Left            =   -69850
         TabIndex        =   121
         Top             =   4350
         Visible         =   0   'False
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3096
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Borrar Movimientos"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmIngresosEgresos.frx":2BFC
      End
      Begin XtremeSuiteControls.PushButton PusImprimir 
         Height          =   345
         Left            =   -57940
         TabIndex        =   107
         Top             =   540
         Visible         =   0   'False
         Width           =   945
         _Version        =   851968
         _ExtentX        =   1667
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmIngresosEgresos.frx":3196
      End
      Begin XtremeSuiteControls.FlatEdit vfiltro 
         Height          =   345
         Left            =   -67750
         TabIndex        =   105
         Top             =   540
         Visible         =   0   'False
         Width           =   4215
         _Version        =   851968
         _ExtentX        =   7435
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid gultimos 
         Height          =   3345
         Left            =   -69850
         TabIndex        =   91
         Top             =   990
         Visible         =   0   'False
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   5900
         _Version        =   393216
         BackColorSel    =   65280
         ForeColorSel    =   255
         SelectionMode   =   1
         AllowUserResizing=   1
         GridLineWidthFixed=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin XtremeSuiteControls.TabControl tab2 
         Height          =   1755
         Left            =   45
         TabIndex        =   39
         Top             =   3270
         Width           =   13965
         _Version        =   851968
         _ExtentX        =   24633
         _ExtentY        =   3096
         _StockProps     =   68
         Color           =   -1972949577
         PaintManager.BoldSelected=   -1  'True
         PaintManager.DisableLunaColors=   0   'False
         PaintManager.HotTracking=   -1  'True
         PaintManager.ShowIcons=   -1  'True
         ItemCount       =   3
         Item(0).Caption =   "Cajas  -  Bancos <F11>"
         Item(0).Tooltip =   "ffgfgf"
         Item(0).ControlCount=   3
         Item(0).Control(0)=   "GroupBox2"
         Item(0).Control(1)=   "Picture5"
         Item(0).Control(2)=   "PusCancelarLos"
         Item(1).Caption =   "Datos del Cheque <12>"
         Item(1).ControlCount=   5
         Item(1).Control(0)=   "lblVlcajabanco"
         Item(1).Control(1)=   "fvalores"
         Item(1).Control(2)=   "fcht"
         Item(1).Control(3)=   "fchp"
         Item(1).Control(4)=   "PusBuscarCheque"
         Item(2).Caption =   "Contabilidad <F10>"
         Item(2).ControlCount=   8
         Item(2).Control(0)=   "GroupBox5"
         Item(2).Control(1)=   "Picture4"
         Item(2).Control(2)=   "txtAlta10"
         Item(2).Control(3)=   "txtAlta11"
         Item(2).Control(4)=   "pbCarga(4)"
         Item(2).Control(5)=   "vleyenda"
         Item(2).Control(6)=   "lblAltaCaja(12)"
         Item(2).Control(7)=   "lblAltaCaja(9)"
         Begin XtremeSuiteControls.PushButton PusBuscarCheque 
            Height          =   315
            Left            =   -66370
            TabIndex        =   136
            Top             =   570
            Visible         =   0   'False
            Width           =   1605
            _Version        =   851968
            _ExtentX        =   2831
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Buscar Cheques"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PusCancelarLos 
            Height          =   345
            Left            =   2925
            TabIndex        =   120
            Top             =   420
            Width           =   3795
            _Version        =   851968
            _ExtentX        =   6694
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "Cancelar los vales pendientes de esta persona"
            UseVisualStyle  =   -1  'True
            Picture         =   "frmIngresosEgresos.frx":3730
         End
         Begin VB.PictureBox Picture4 
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   -65320
            Picture         =   "frmIngresosEgresos.frx":3CCA
            ScaleHeight     =   285
            ScaleWidth      =   285
            TabIndex        =   98
            Top             =   600
            Visible         =   0   'False
            Width           =   285
         End
         Begin XtremeSuiteControls.GroupBox fchp 
            Height          =   705
            Left            =   -69850
            TabIndex        =   85
            Top             =   960
            Visible         =   0   'False
            Width           =   3405
            _Version        =   851968
            _ExtentX        =   6006
            _ExtentY        =   1244
            _StockProps     =   79
            Caption         =   "Cheques Propios:"
            UseVisualStyle  =   -1  'True
            BorderStyle     =   1
            Begin XtremeSuiteControls.PushButton PusCargarDatos 
               Height          =   345
               Left            =   30
               TabIndex        =   86
               Top             =   240
               Width           =   3285
               _Version        =   851968
               _ExtentX        =   5794
               _ExtentY        =   609
               _StockProps     =   79
               Caption         =   "Cargar datos cheques propios"
               ForeColor       =   0
               Appearance      =   3
               Picture         =   "frmIngresosEgresos.frx":4254
            End
         End
         Begin XtremeSuiteControls.GroupBox fcht 
            Height          =   585
            Left            =   -69880
            TabIndex        =   83
            Top             =   360
            Visible         =   0   'False
            Width           =   3465
            _Version        =   851968
            _ExtentX        =   6112
            _ExtentY        =   1032
            _StockProps     =   79
            Caption         =   "Cheques de Terceros:"
            UseVisualStyle  =   -1  'True
            BorderStyle     =   1
            Begin XtremeSuiteControls.PushButton PushButton1 
               Height          =   315
               Left            =   30
               TabIndex        =   84
               Top             =   210
               Width           =   3345
               _Version        =   851968
               _ExtentX        =   5900
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "Cheques"
               ForeColor       =   0
               BackColor       =   255
               Appearance      =   3
               Picture         =   "frmIngresosEgresos.frx":47EE
            End
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            CausesValidation=   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   480
            Picture         =   "frmIngresosEgresos.frx":4D88
            ScaleHeight     =   285
            ScaleWidth      =   255
            TabIndex        =   45
            Top             =   840
            Width           =   255
         End
         Begin XtremeSuiteControls.GroupBox GroupBox2 
            Height          =   405
            Left            =   210
            TabIndex        =   40
            Top             =   840
            Width           =   12555
            _Version        =   851968
            _ExtentX        =   22146
            _ExtentY        =   714
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            BorderStyle     =   2
            Begin XtremeSuiteControls.FlatEdit txtAlta6 
               Height          =   315
               Left            =   2700
               TabIndex        =   41
               Top             =   30
               Width           =   1335
               _Version        =   851968
               _ExtentX        =   2355
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.PushButton pbCarga 
               Height          =   285
               Index           =   2
               Left            =   4110
               TabIndex        =   2
               Tag             =   "CajaBanco"
               Top             =   30
               Width           =   405
               _Version        =   851968
               _ExtentX        =   714
               _ExtentY        =   503
               _StockProps     =   79
               Caption         =   "..."
               UseVisualStyle  =   -1  'True
               Picture         =   "frmIngresosEgresos.frx":5312
            End
            Begin XtremeSuiteControls.FlatEdit txtAlta7 
               Height          =   315
               Left            =   4560
               TabIndex        =   42
               Top             =   0
               Width           =   7935
               _Version        =   851968
               _ExtentX        =   13996
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.Label lblAltaCaja 
               Height          =   195
               Index           =   7
               Left            =   600
               TabIndex        =   43
               Top             =   30
               Width           =   1815
               _Version        =   851968
               _ExtentX        =   3201
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "Caja / Bando actual:"
               ForeColor       =   0
               Alignment       =   1
               Transparent     =   -1  'True
            End
         End
         Begin XtremeSuiteControls.GroupBox fvalores 
            Height          =   1215
            Left            =   -66790
            TabIndex        =   61
            Top             =   480
            Visible         =   0   'False
            Width           =   9585
            _Version        =   851968
            _ExtentX        =   16907
            _ExtentY        =   2143
            _StockProps     =   79
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
            BorderStyle     =   2
            Begin XtremeSuiteControls.FlatEdit txtAlta 
               Height          =   315
               Index           =   5
               Left            =   3930
               TabIndex        =   62
               Top             =   390
               Width           =   2385
               _Version        =   851968
               _ExtentX        =   4207
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
               Alignment       =   1
            End
            Begin XtremeSuiteControls.PushButton PushButton2 
               Height          =   285
               Left            =   5280
               TabIndex        =   63
               Top             =   780
               Width           =   345
               _Version        =   851968
               _ExtentX        =   609
               _ExtentY        =   503
               _StockProps     =   79
               Caption         =   "..."
               UseVisualStyle  =   -1  'True
               Picture         =   "frmIngresosEgresos.frx":58AC
            End
            Begin XtremeSuiteControls.FlatEdit vNuevaCustodiaCodigo 
               Height          =   315
               Left            =   3930
               TabIndex        =   64
               Top             =   750
               Width           =   1335
               _Version        =   851968
               _ExtentX        =   2355
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.FlatEdit VNuevaCustodiaNombre 
               Height          =   315
               Left            =   5670
               TabIndex        =   65
               Top             =   750
               Width           =   3795
               _Version        =   851968
               _ExtentX        =   6694
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin Aplisoft_CajasDeTexto.TxF dtpValor 
               Height          =   315
               Left            =   7890
               TabIndex        =   66
               Top             =   420
               Width           =   1560
               _ExtentX        =   2752
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
            Begin XtremeSuiteControls.FlatEdit txtAlta 
               Height          =   315
               Index           =   8
               Left            =   9240
               TabIndex        =   67
               Top             =   600
               Visible         =   0   'False
               Width           =   285
               _Version        =   851968
               _ExtentX        =   503
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.PushButton pbCarga 
               Height          =   285
               Index           =   3
               Left            =   8280
               TabIndex        =   68
               Tag             =   "BancoCuenta"
               Top             =   270
               Visible         =   0   'False
               Width           =   345
               _Version        =   851968
               _ExtentX        =   609
               _ExtentY        =   503
               _StockProps     =   79
               Caption         =   "..."
               UseVisualStyle  =   -1  'True
               Picture         =   "frmIngresosEgresos.frx":5E46
            End
            Begin XtremeSuiteControls.FlatEdit txtAlta 
               Height          =   315
               Index           =   9
               Left            =   8700
               TabIndex        =   69
               Top             =   270
               Visible         =   0   'False
               Width           =   3705
               _Version        =   851968
               _ExtentX        =   6535
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.FlatEdit vCodBanco 
               Height          =   315
               Left            =   3930
               TabIndex        =   70
               Top             =   30
               Width           =   1185
               _Version        =   851968
               _ExtentX        =   2090
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.FlatEdit vDesBanco 
               Height          =   315
               Left            =   5550
               TabIndex        =   71
               Top             =   60
               Width           =   3975
               _Version        =   851968
               _ExtentX        =   7011
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.PushButton PushButton3 
               Height          =   285
               Left            =   5160
               TabIndex        =   72
               Top             =   60
               Width           =   345
               _Version        =   851968
               _ExtentX        =   609
               _ExtentY        =   503
               _StockProps     =   79
               Caption         =   "..."
               UseVisualStyle  =   -1  'True
               Picture         =   "frmIngresosEgresos.frx":63E0
            End
            Begin XtremeSuiteControls.Label lblAltaCaja 
               Height          =   195
               Index           =   5
               Left            =   2550
               TabIndex        =   77
               Top             =   480
               Width           =   1305
               _Version        =   851968
               _ExtentX        =   2302
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "Nro. Documento:"
               ForeColor       =   0
               Alignment       =   1
               Transparent     =   -1  'True
            End
            Begin XtremeSuiteControls.Label Label1 
               Height          =   315
               Left            =   960
               TabIndex        =   76
               Top             =   720
               Width           =   2955
               _Version        =   851968
               _ExtentX        =   5212
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "Caja o Banco donde se imputa:"
               ForeColor       =   255
               Alignment       =   1
            End
            Begin XtremeSuiteControls.Label lblAltaCaja 
               Height          =   195
               Index           =   6
               Left            =   6360
               TabIndex        =   75
               Top             =   450
               Width           =   1455
               _Version        =   851968
               _ExtentX        =   2566
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "Fecha Valor / Caja:"
               ForeColor       =   0
               Transparent     =   -1  'True
            End
            Begin XtremeSuiteControls.Label lblBanco 
               Height          =   315
               Left            =   3180
               TabIndex        =   74
               Top             =   30
               Width           =   615
               _Version        =   851968
               _ExtentX        =   1085
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "Banco:"
               ForeColor       =   0
               Alignment       =   1
            End
            Begin XtremeSuiteControls.Label lblCuenta 
               Height          =   285
               Left            =   8100
               TabIndex        =   73
               Top             =   630
               Visible         =   0   'False
               Width           =   855
               _Version        =   851968
               _ExtentX        =   1508
               _ExtentY        =   503
               _StockProps     =   79
               Caption         =   "Cuenta:"
               ForeColor       =   0
               Alignment       =   1
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox5 
            Height          =   30
            Left            =   -69850
            TabIndex        =   92
            Top             =   1395
            Visible         =   0   'False
            Width           =   12555
            _Version        =   851968
            _ExtentX        =   22146
            _ExtentY        =   53
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            BorderStyle     =   2
            Begin XtremeSuiteControls.FlatEdit vcodgocta2 
               Height          =   315
               Left            =   5940
               TabIndex        =   93
               Top             =   210
               Visible         =   0   'False
               Width           =   1965
               _Version        =   851968
               _ExtentX        =   3466
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.FlatEdit vcta2 
               Height          =   315
               Left            =   8310
               TabIndex        =   94
               Top             =   90
               Visible         =   0   'False
               Width           =   4185
               _Version        =   851968
               _ExtentX        =   7382
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.PushButton pbCarga2 
               Height          =   315
               Index           =   5
               Left            =   7950
               TabIndex        =   95
               Tag             =   "CodigoCuenta"
               Top             =   210
               Visible         =   0   'False
               Width           =   375
               _Version        =   851968
               _ExtentX        =   661
               _ExtentY        =   556
               _StockProps     =   79
               UseVisualStyle  =   -1  'True
               Picture         =   "frmIngresosEgresos.frx":697A
            End
            Begin VB.Label Label3 
               Caption         =   "vlcajabanco"
               Height          =   255
               Left            =   5160
               TabIndex        =   97
               Top             =   240
               Visible         =   0   'False
               Width           =   585
            End
            Begin XtremeSuiteControls.Label lblAltaCaja 
               Height          =   195
               Index           =   13
               Left            =   210
               TabIndex        =   96
               Top             =   240
               Visible         =   0   'False
               Width           =   4215
               _Version        =   851968
               _ExtentX        =   7435
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "- Cta Contable asociada a personas o entidades: "
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
               Transparent     =   -1  'True
            End
         End
         Begin XtremeSuiteControls.FlatEdit txtAlta10 
            Height          =   315
            Left            =   -64960
            TabIndex        =   99
            Top             =   600
            Visible         =   0   'False
            Width           =   2265
            _Version        =   851968
            _ExtentX        =   3995
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit txtAlta11 
            Height          =   315
            Left            =   -61780
            TabIndex        =   100
            Top             =   600
            Visible         =   0   'False
            Width           =   4395
            _Version        =   851968
            _ExtentX        =   7752
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.PushButton pbCarga 
            Height          =   315
            Index           =   4
            Left            =   -62650
            TabIndex        =   101
            Tag             =   "CodigoCuenta"
            Top             =   600
            Visible         =   0   'False
            Width           =   795
            _Version        =   851968
            _ExtentX        =   1402
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "<F7>"
            UseVisualStyle  =   -1  'True
            Picture         =   "frmIngresosEgresos.frx":6F14
         End
         Begin XtremeSuiteControls.FlatEdit vleyenda 
            Height          =   345
            Left            =   -66220
            TabIndex        =   102
            Top             =   990
            Visible         =   0   'False
            Width           =   8865
            _Version        =   851968
            _ExtentX        =   15637
            _ExtentY        =   609
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.Label lblAltaCaja 
            Height          =   195
            Index           =   9
            Left            =   -68290
            TabIndex        =   104
            Top             =   660
            Visible         =   0   'False
            Width           =   2865
            _Version        =   851968
            _ExtentX        =   5054
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "- Cta Contable asociada Caja Banco : "
            ForeColor       =   0
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblAltaCaja 
            Height          =   195
            Index           =   12
            Left            =   -68410
            TabIndex        =   103
            Top             =   1050
            Visible         =   0   'False
            Width           =   2055
            _Version        =   851968
            _ExtentX        =   3625
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Leyeda de la contabilidad: "
            Transparent     =   -1  'True
         End
         Begin VB.Label lblVlcajabanco 
            Caption         =   "vlcajabanco"
            Height          =   255
            Left            =   -64960
            TabIndex        =   44
            Top             =   510
            Visible         =   0   'False
            Width           =   15
         End
      End
      Begin VB.TextBox vobservacion2 
         Height          =   285
         Left            =   2580
         TabIndex        =   87
         Top             =   7980
         Width           =   10845
      End
      Begin XtremeSuiteControls.PushButton PusPersonas 
         Height          =   345
         Left            =   7500
         TabIndex        =   82
         Top             =   2070
         Width           =   1065
         _Version        =   851968
         _ExtentX        =   1879
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Personas"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmIngresosEgresos.frx":74AE
      End
      Begin VB.TextBox vcliprovee 
         Height          =   315
         Left            =   2010
         TabIndex        =   57
         Top             =   2100
         Width           =   4275
      End
      Begin VB.TextBox vrendicion 
         Height          =   315
         Left            =   5700
         TabIndex        =   55
         Top             =   1710
         Width           =   2235
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Event. +"
         Height          =   345
         Left            =   12360
         TabIndex        =   9
         Top             =   2070
         Width           =   735
      End
      Begin VB.TextBox vconcepto 
         Height          =   315
         Left            =   1770
         TabIndex        =   46
         Top             =   1710
         Width           =   2175
      End
      Begin VB.CommandButton cmdEventuales 
         Caption         =   "Eventuales"
         Height          =   345
         Left            =   11460
         TabIndex        =   8
         Top             =   2070
         Width           =   915
      End
      Begin VB.TextBox vobservacion 
         Height          =   285
         Left            =   2580
         TabIndex        =   38
         Top             =   7680
         Width           =   10845
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Asistido"
         Height          =   345
         Left            =   10770
         TabIndex        =   7
         Top             =   2070
         Width           =   675
      End
      Begin VB.CommandButton cmdContribuyente 
         Caption         =   "Contribuyente"
         Height          =   345
         Left            =   9690
         TabIndex        =   6
         Top             =   2070
         Width           =   1095
      End
      Begin XtremeSuiteControls.GroupBox GBRBSuperior 
         Height          =   375
         Left            =   8070
         TabIndex        =   14
         Top             =   1650
         Width           =   2505
         _Version        =   851968
         _ExtentX        =   4419
         _ExtentY        =   661
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton RBIngresoEgresoCaja 
            Height          =   195
            Index           =   0
            Left            =   510
            TabIndex        =   17
            Top             =   120
            Visible         =   0   'False
            Width           =   1845
            _Version        =   851968
            _ExtentX        =   3254
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Movimiento de Ingreso"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RBIngresoEgresoCaja 
            Height          =   255
            Index           =   1
            Left            =   1110
            TabIndex        =   16
            Top             =   120
            Visible         =   0   'False
            Width           =   1845
            _Version        =   851968
            _ExtentX        =   3254
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Movimiento de Egreso"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox chkIEautomatico 
            Height          =   315
            Left            =   4590
            TabIndex        =   29
            Top             =   150
            Visible         =   0   'False
            Width           =   2055
            _Version        =   851968
            _ExtentX        =   3625
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Contabilidad Automática"
            ForeColor       =   0
            BackColor       =   1375373
            UseVisualStyle  =   -1  'True
            Value           =   1
         End
         Begin XtremeSuiteControls.Label lblAltaCaja 
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   18
            Top             =   120
            Visible         =   0   'False
            Width           =   495
            _Version        =   851968
            _ExtentX        =   873
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Tipo :"
            Transparent     =   -1  'True
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   210
         TabIndex        =   26
         Top             =   5070
         Width           =   12825
         Begin XtremeSuiteControls.PushButton cmdAgregar 
            Height          =   285
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   1545
            _Version        =   851968
            _ExtentX        =   2725
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Agregar <F6>"
            UseVisualStyle  =   -1  'True
            Picture         =   "frmIngresosEgresos.frx":7A48
         End
         Begin XtremeSuiteControls.PushButton cmdLimpiar 
            Height          =   285
            Left            =   4830
            TabIndex        =   27
            Top             =   0
            Width           =   1335
            _Version        =   851968
            _ExtentX        =   2355
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Limpiar Grilla"
            UseVisualStyle  =   -1  'True
            Picture         =   "frmIngresosEgresos.frx":7FE2
         End
         Begin XtremeSuiteControls.PushButton cmdBorrar 
            Height          =   285
            Left            =   3540
            TabIndex        =   28
            Top             =   0
            Width           =   1305
            _Version        =   851968
            _ExtentX        =   2302
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Borrar Linea"
            UseVisualStyle  =   -1  'True
            Picture         =   "frmIngresosEgresos.frx":E844
         End
         Begin XtremeSuiteControls.Label lblFalta 
            Height          =   255
            Left            =   10710
            TabIndex        =   139
            Top             =   30
            Width           =   435
            _Version        =   851968
            _ExtentX        =   767
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Falta:"
         End
         Begin XtremeSuiteControls.Label lblSubtotal 
            Height          =   255
            Left            =   7740
            TabIndex        =   138
            Top             =   30
            Width           =   735
            _Version        =   851968
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Subtotal:"
         End
         Begin XtremeSuiteControls.Label lfalta 
            Height          =   285
            Left            =   11220
            TabIndex        =   135
            Top             =   0
            Width           =   1575
            _Version        =   851968
            _ExtentX        =   2778
            _ExtentY        =   503
            _StockProps     =   79
            ForeColor       =   16711935
            BackColor       =   -2147483641
            Alignment       =   1
         End
         Begin XtremeSuiteControls.Label ltotal 
            Height          =   285
            Left            =   8550
            TabIndex        =   134
            Top             =   0
            Width           =   1395
            _Version        =   851968
            _ExtentX        =   2461
            _ExtentY        =   503
            _StockProps     =   79
            ForeColor       =   65535
            BackColor       =   -2147483641
            Alignment       =   1
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   915
         Left            =   150
         TabIndex        =   19
         Top             =   720
         Width           =   10425
         _Version        =   851968
         _ExtentX        =   18389
         _ExtentY        =   1614
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
         Begin XtremeSuiteControls.GroupBox GroupBox6 
            Height          =   825
            Left            =   5640
            TabIndex        =   128
            Top             =   180
            Width           =   5265
            _Version        =   851968
            _ExtentX        =   9287
            _ExtentY        =   1455
            _StockProps     =   79
            Caption         =   "Saldo Disponible:"
            Appearance      =   5
            BorderStyle     =   1
            Begin Project1.bsGradientLabel vsaldoValores 
               Height          =   315
               Left            =   3450
               Top             =   360
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   556
               Caption         =   ""
               BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Terminal"
                  Size            =   9
                  Charset         =   255
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Colour1         =   255
               Colour2         =   255
               CaptionAlignment=   1
            End
            Begin XtremeSuiteControls.PushButton PusActualizar 
               Height          =   225
               Left            =   2550
               TabIndex        =   142
               Top             =   120
               Width           =   825
               _Version        =   851968
               _ExtentX        =   1455
               _ExtentY        =   397
               _StockProps     =   79
               Caption         =   "Actualizar"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.PushButton PusDetalle 
               Height          =   225
               Left            =   3390
               TabIndex        =   130
               Top             =   120
               Width           =   645
               _Version        =   851968
               _ExtentX        =   1138
               _ExtentY        =   397
               _StockProps     =   79
               Caption         =   "Detalle"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.PushButton PusSaldos 
               Height          =   225
               Left            =   4050
               TabIndex        =   144
               Top             =   120
               Width           =   645
               _Version        =   851968
               _ExtentX        =   1138
               _ExtentY        =   397
               _StockProps     =   79
               Caption         =   "Saldos"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.PushButton PusValoresCartera 
               Height          =   195
               Left            =   2550
               TabIndex        =   149
               Top             =   360
               Width           =   885
               _Version        =   851968
               _ExtentX        =   1561
               _ExtentY        =   344
               _StockProps     =   79
               Caption         =   "VerCartera"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.Label vsaldoDisponible 
               Height          =   345
               Left            =   180
               TabIndex        =   129
               Top             =   210
               Width           =   2295
               _Version        =   851968
               _ExtentX        =   4048
               _ExtentY        =   609
               _StockProps     =   79
               ForeColor       =   65280
               BackColor       =   4210752
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   2
            End
         End
         Begin XtremeSuiteControls.CheckBox chkNroInternoFijo 
            Height          =   255
            Left            =   7770
            TabIndex        =   33
            Top             =   780
            Visible         =   0   'False
            Width           =   2235
            _Version        =   851968
            _ExtentX        =   3942
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Nro interno manualmente"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtAlta 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   20
            Top             =   150
            Width           =   975
            _Version        =   851968
            _ExtentX        =   1720
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.PushButton pbCarga 
            Height          =   315
            Index           =   0
            Left            =   2670
            TabIndex        =   4
            Tag             =   "TipoMovimientosBanco"
            Top             =   150
            Width           =   375
            _Version        =   851968
            _ExtentX        =   661
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
            Picture         =   "frmIngresosEgresos.frx":150A6
         End
         Begin XtremeSuiteControls.FlatEdit txtAlta 
            Height          =   315
            Index           =   1
            Left            =   3060
            TabIndex        =   21
            Top             =   150
            Width           =   2505
            _Version        =   851968
            _ExtentX        =   4419
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit txtAlta 
            Height          =   360
            Index           =   2
            Left            =   5685
            TabIndex        =   22
            Top             =   555
            Width           =   2955
            _Version        =   851968
            _ExtentX        =   5212
            _ExtentY        =   635
            _StockProps     =   77
            BackColor       =   -2147483643
            Enabled         =   0   'False
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit vtotalcontrol 
            Height          =   315
            Left            =   3840
            TabIndex        =   3
            Top             =   600
            Width           =   1455
            _Version        =   851968
            _ExtentX        =   2566
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin Aplisoft_CajasDeTexto.TxF dtpFecha 
            Height          =   285
            Left            =   720
            TabIndex        =   167
            Top             =   540
            Width           =   1545
            _ExtentX        =   2725
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
         Begin XtremeSuiteControls.Label lblConciliarTotales 
            Height          =   255
            Left            =   2340
            TabIndex        =   56
            Top             =   600
            Width           =   1365
            _Version        =   851968
            _ExtentX        =   2408
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Conciliar Totales:"
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
         Begin XtremeSuiteControls.Label lblAltaCaja 
            Height          =   195
            Index           =   3
            Left            =   5370
            TabIndex        =   25
            Top             =   660
            Width           =   915
            _Version        =   851968
            _ExtentX        =   1614
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Nro Interno:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblAltaCaja 
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   24
            Top             =   600
            Width           =   615
            _Version        =   851968
            _ExtentX        =   1085
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Fecha :"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblAltaCaja 
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   23
            Top             =   180
            Width           =   1665
            _Version        =   851968
            _ExtentX        =   2937
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Tipo de Movimientos:"
            Alignment       =   1
            Transparent     =   -1  'True
         End
      End
      Begin Grid.KlexGrid KlexMovimientoCaja 
         Height          =   2265
         Left            =   120
         TabIndex        =   13
         Top             =   5400
         Width           =   13995
         _ExtentX        =   24686
         _ExtentY        =   3995
         GridLinesFixed  =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmIngresosEgresos.frx":15640
      End
      Begin XtremeSuiteControls.PushButton cmdCheque 
         Height          =   315
         Left            =   9330
         TabIndex        =   15
         Top             =   4530
         Width           =   2355
         _Version        =   851968
         _ExtentX        =   4154
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Ing. más datos del cheque"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   8
         MultiLine       =   0   'False
         Picture         =   "frmIngresosEgresos.frx":1565C
         TextImageRelation=   4
      End
      Begin XtremeSuiteControls.GroupBox gcustodia 
         Height          =   525
         Left            =   60
         TabIndex        =   34
         Top             =   4410
         Width           =   12945
         _Version        =   851968
         _ExtentX        =   22834
         _ExtentY        =   926
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         BorderStyle     =   2
      End
      Begin XtremeSuiteControls.PushButton PusSelConceptos 
         Height          =   315
         Left            =   210
         TabIndex        =   5
         Top             =   1710
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Conceptos  <F4>"
         ForeColor       =   0
         BackColor       =   -2147483644
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtAlta 
         Height          =   315
         Index           =   12
         Left            =   1590
         TabIndex        =   1
         Top             =   2880
         Width           =   1395
         _Version        =   851968
         _ExtentX        =   2461
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   255
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   315
         Left            =   4200
         TabIndex        =   0
         Top             =   1740
         Width           =   1305
         _Version        =   851968
         _ExtentX        =   2302
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Rendición  <F9>"
         ForeColor       =   0
         BackColor       =   -2147483644
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtAlta 
         Height          =   345
         Index           =   13
         Left            =   4470
         TabIndex        =   59
         Top             =   2940
         Width           =   9675
         _Version        =   851968
         _ExtentX        =   17066
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
         MaxLength       =   250
      End
      Begin XtremeSuiteControls.FlatEdit txtAlta 
         Height          =   315
         Index           =   3
         Left            =   1500
         TabIndex        =   78
         Top             =   2460
         Width           =   1395
         _Version        =   851968
         _ExtentX        =   2461
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   1
         Left            =   2940
         TabIndex        =   79
         Tag             =   "TipoValor"
         Top             =   2460
         Width           =   615
         _Version        =   851968
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtAlta 
         Height          =   315
         Index           =   4
         Left            =   3600
         TabIndex        =   80
         Top             =   2460
         Width           =   3915
         _Version        =   851968
         _ExtentX        =   6906
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton PushButton5 
         Height          =   405
         Left            =   -58960
         TabIndex        =   108
         Top             =   510
         Visible         =   0   'False
         Width           =   1965
         _Version        =   851968
         _ExtentX        =   3466
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit1 
         Height          =   345
         Left            =   -67750
         TabIndex        =   109
         Top             =   540
         Visible         =   0   'False
         Width           =   8565
         _Version        =   851968
         _ExtentX        =   15108
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7065
         Left            =   -69820
         TabIndex        =   110
         Top             =   1140
         Visible         =   0   'False
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   12462
         _Version        =   393216
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin XtremeSuiteControls.PushButton PushButton6 
         Height          =   405
         Left            =   -59020
         TabIndex        =   112
         Top             =   540
         Visible         =   0   'False
         Width           =   1965
         _Version        =   851968
         _ExtentX        =   3466
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit2 
         Height          =   345
         Left            =   -67810
         TabIndex        =   113
         Top             =   570
         Visible         =   0   'False
         Width           =   8565
         _Version        =   851968
         _ExtentX        =   15108
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   7065
         Left            =   -69880
         TabIndex        =   114
         Top             =   1170
         Visible         =   0   'False
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   12462
         _Version        =   393216
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   4395
         Left            =   -69850
         TabIndex        =   116
         Top             =   930
         Visible         =   0   'False
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   7752
         _Version        =   393216
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin XtremeSuiteControls.PushButton PushButton9 
         Height          =   345
         Left            =   -62500
         TabIndex        =   117
         Top             =   540
         Visible         =   0   'False
         Width           =   1515
         _Version        =   851968
         _ExtentX        =   2672
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Saldos por Caja"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PusSaldosPor 
         Height          =   345
         Left            =   -60970
         TabIndex        =   118
         Top             =   540
         Visible         =   0   'False
         Width           =   1515
         _Version        =   851968
         _ExtentX        =   2672
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Saldos por Dias"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton10 
         Height          =   345
         Left            =   -59440
         TabIndex        =   119
         Top             =   540
         Visible         =   0   'False
         Width           =   1515
         _Version        =   851968
         _ExtentX        =   2672
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Saldos por Mes"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton12 
         Height          =   255
         Left            =   2220
         TabIndex        =   123
         Top             =   7680
         Width           =   345
         _Version        =   851968
         _ExtentX        =   609
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton13 
         Height          =   255
         Left            =   2220
         TabIndex        =   124
         Top             =   7980
         Width           =   345
         _Version        =   851968
         _ExtentX        =   609
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton8 
         Height          =   405
         Left            =   -69820
         TabIndex        =   126
         Top             =   690
         Visible         =   0   'False
         Width           =   2055
         _Version        =   851968
         _ExtentX        =   3625
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Cierre de Caja"
         UseVisualStyle  =   -1  'True
         PushButtonStyle =   2
      End
      Begin XtremeSuiteControls.PushButton PushButton14 
         Height          =   345
         Left            =   -63340
         TabIndex        =   131
         Top             =   540
         Visible         =   0   'False
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Todos"
         UseVisualStyle  =   -1  'True
      End
      Begin Grid.KlexGrid gdetalle 
         Height          =   1245
         Left            =   -69850
         TabIndex        =   132
         Top             =   4740
         Visible         =   0   'False
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   2196
         GridLinesFixed  =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmIngresosEgresos.frx":15BF6
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
         Height          =   2415
         Left            =   -69820
         TabIndex        =   143
         Top             =   5850
         Visible         =   0   'False
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   4260
         _Version        =   393216
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin XtremeSuiteControls.PushButton PushButton7 
         Height          =   405
         Left            =   -67660
         TabIndex        =   145
         Top             =   1260
         Visible         =   0   'False
         Width           =   2055
         _Version        =   851968
         _ExtentX        =   3625
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Saldos"
         UseVisualStyle  =   -1  'True
         PushButtonStyle =   2
      End
      Begin XtremeSuiteControls.GroupBox f1 
         Height          =   360
         Left            =   90
         TabIndex        =   48
         Top             =   330
         Width           =   10875
         _Version        =   851968
         _ExtentX        =   19182
         _ExtentY        =   635
         _StockProps     =   79
         BackColor       =   255
         UseVisualStyle  =   -1  'True
         BorderStyle     =   2
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   8010
            ScaleHeight     =   225
            ScaleWidth      =   135
            TabIndex        =   50
            Top             =   90
            Width           =   135
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   5610
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   49
            Top             =   90
            Width           =   225
         End
         Begin XtremeSuiteControls.RadioButton RBDebeHaber 
            Height          =   225
            Index           =   1
            Left            =   4020
            TabIndex        =   51
            Top             =   90
            Width           =   1875
            _Version        =   851968
            _ExtentX        =   3307
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "Credito (Retiro)"
            ForeColor       =   0
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
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RBDebeHaber 
            Height          =   225
            Index           =   0
            Left            =   6405
            TabIndex        =   52
            Top             =   90
            Width           =   1845
            _Version        =   851968
            _ExtentX        =   3254
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "Debe (Ingreso)"
            ForeColor       =   0
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
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   225
            Left            =   60
            TabIndex        =   53
            Top             =   120
            Width           =   3615
            _Version        =   851968
            _ExtentX        =   6376
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "Con <F8> cambia de Ingreso a Egreso:"
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
            Transparent     =   -1  'True
         End
      End
      Begin VB.Label lblAsientos 
         Alignment       =   1  'Right Justify
         Caption         =   "Personas:"
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
         Index           =   11
         Left            =   150
         TabIndex        =   58
         Top             =   2130
         Width           =   1095
      End
      Begin XtremeSuiteControls.Label lbsaldo 
         Height          =   345
         Left            =   6300
         TabIndex        =   141
         Top             =   2070
         Width           =   1125
         _Version        =   851968
         _ExtentX        =   1984
         _ExtentY        =   609
         _StockProps     =   79
         ForeColor       =   1375373
         BackColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   405
         Left            =   -67630
         TabIndex        =   127
         Top             =   690
         Visible         =   0   'False
         Width           =   10215
         _Version        =   851968
         _ExtentX        =   18018
         _ExtentY        =   714
         _StockProps     =   79
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   435
         Left            =   -69850
         TabIndex        =   115
         Top             =   510
         Visible         =   0   'False
         Width           =   2085
         _Version        =   851968
         _ExtentX        =   3678
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Buscar por Fecha - Codigo :"
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   435
         Left            =   -69790
         TabIndex        =   111
         Top             =   480
         Visible         =   0   'False
         Width           =   2085
         _Version        =   851968
         _ExtentX        =   3678
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Buscar por Fecha - Codigo :"
      End
      Begin XtremeSuiteControls.Label lblBuscarPor 
         Height          =   435
         Left            =   -69790
         TabIndex        =   106
         Top             =   480
         Visible         =   0   'False
         Width           =   2085
         _Version        =   851968
         _ExtentX        =   3678
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Buscar por Fecha - Codigo :"
      End
      Begin XtremeSuiteControls.Label lblCtaSeleccionada 
         Height          =   165
         Left            =   7650
         TabIndex        =   90
         Top             =   2520
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   291
         _StockProps     =   79
         Caption         =   "Cta. seleccionada:"
         ForeColor       =   0
      End
      Begin XtremeSuiteControls.Label lblCta 
         Height          =   285
         Left            =   9060
         TabIndex        =   89
         Top             =   2490
         Width           =   3945
         _Version        =   851968
         _ExtentX        =   6959
         _ExtentY        =   503
         _StockProps     =   79
         ForeColor       =   4210752
      End
      Begin VB.Label Label4 
         Caption         =   "Observación Gral. Linea 2:"
         Height          =   225
         Left            =   120
         TabIndex        =   88
         Top             =   8010
         Width           =   1905
      End
      Begin XtremeSuiteControls.Label lblAltaCaja 
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   81
         Top             =   2490
         Width           =   1245
         _Version        =   851968
         _ExtentX        =   2205
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Tipo Valor :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblAltaCaja 
         Height          =   195
         Index           =   11
         Left            =   3240
         TabIndex        =   60
         Top             =   3030
         Width           =   1185
         _Version        =   851968
         _ExtentX        =   2090
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Observaciones:"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblAltaCaja 
         Height          =   195
         Index           =   10
         Left            =   180
         TabIndex        =   54
         Top             =   2910
         Width           =   1365
         _Version        =   851968
         _ExtentX        =   2408
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Importe <F3> :"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin VB.Label lblIngresarUna 
         Caption         =   "Observación Gral. Linea 1:"
         Height          =   225
         Left            =   120
         TabIndex        =   37
         Top             =   7740
         Width           =   1905
      End
      Begin XtremeSuiteControls.Label VchequesDisplay 
         Height          =   255
         Left            =   150
         TabIndex        =   35
         Top             =   2400
         Width           =   45
         _Version        =   851968
         _ExtentX        =   -79
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Vcheques display"
         ForeColor       =   14737632
      End
   End
   Begin XtremeSuiteControls.PushButton PusImprimirComprobante 
      Height          =   345
      Left            =   11970
      TabIndex        =   133
      Top             =   45
      Width           =   1125
      _Version        =   851968
      _ExtentX        =   1984
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Previualizar"
      UseVisualStyle  =   -1  'True
      BorderGap       =   10
   End
   Begin XtremeSuiteControls.PushButton Pagare 
      Height          =   345
      Left            =   11115
      TabIndex        =   150
      Top             =   45
      Width           =   825
      _Version        =   851968
      _ExtentX        =   1455
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Pagare"
      UseVisualStyle  =   -1  'True
      BorderGap       =   10
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   9015
      Left            =   14220
      Top             =   -90
      Width           =   2055
   End
End
Attribute VB_Name = "frmIngresosEgresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vguardaArreglo As Boolean
Dim vnroasiento As Long, vnrobalance As Long
Dim vEstaModificando As Boolean
Dim vCodigoCliente, vCodigoProveedor As String
Public vidcheque As Long
Dim vsql2 As String
Dim vidBancoCaja As Long
Dim vidpersonas, vidclientes, vidproveedores As Long
Dim vtotal As Double
Dim vcp As String
Dim vDraft As Boolean
Public vIdCheques As Long
Dim vnrocomprobante As Long
Dim vnrointerno As Long
' banderas
Dim vimporte_pagare, vintereses_pagare As Double

Dim vhaciendo As String


Dim bandNrocheque As Boolean

Dim vultimo_ingreso As Double

Dim vtipoCheque As String

Dim vcampoPctasctes As String
Dim vltotal As Double

Dim vsqlpago() As String
Dim vsqlpagoAuto() As String


Dim atrans(100) As String

Public vtotalseleccionado As Double

Public vpagoPacial As Boolean



'Public Type recibo
'    numero As Integer
'    lugar As String
'    fecha As Date
'    persona As String
'    total As Double
'    totalEnletra As String
'    saldo As Double
'    comentario As String
'End Type

'Dim vrecibo As recibo
                    
Private Sub cmd_Click()
Call fbuscarGrilla("clientes", "Nombre", "Codigo", Me.vcliprovee.Name, Me, , False)  ' ema:
vcp = "p"
End Sub





Function valDatosRelacionados() As Boolean

Dim vpcb, vsql  As String

vsql = "select ProveedorClienteBanco as c from tipomovimientos where codigo = '" + Me.txtAlta(0).Text + "'"

vpcb = traerDatos2(vsql, "c", pathDBMySQL)
valDatosRelacionados = True

If Me.RBDebeHaber(0).Value And vpcb = "P" Then
    MsgBox "Debe selección un movimiento de Retiro del dinero", vbInformation
     Call initIngreso
    valDatosRelacionados = False
End If


If Me.txtAlta(0).Text = "VL" And Trim(Me.txtAlta6.Text) = "" And Trim(Me.vCodBanco.Text) = "" Then
    MsgBox "Debe ingresar una cuenta de arqueo de Caja/Banco para completar el VALE", vbInformation
    valDatosRelacionados = False
End If

End Function
Private Sub cmdAgregar_Click()
On Error Resume Next

    If Me.vNuevaCustodiaCodigo.Text = "" And Not vCodBanco.Text = "" Then
        MsgBox "Debe seleccionar una caja referida a la custodia del cheque ", vbCritical
        Exit Sub
    End If

    
    If vhaciendo Then
      vimporte_pagare = vimporte_pagare + Val(Me.txtAlta(12))
   End If
    
    vtipoCheque = ""

    '- validaciones --

    If Not cajaAbierta(Me.dtpFecha.Value) Then Exit Sub

    If Not valDatosObligatorio Then Exit Sub

    If Not valDatosRelacionados Then Exit Sub
    
    'If Not valFechaGrilla Then Exit Sub
    '-----------------


    If Me.RBDebeHaber(0).Value = False And Me.RBDebeHaber(1).Value = False Then
        MsgBox "Debe selecionar acción de retiro / ingreso", vbExclamation, "Mensaje ..."
        Exit Sub
    End If
    

    
    If txtAlta(3).Text = "EF" And Not Val(txtAlta(5).Text) = 0 Then
        'MsgBox "No puede ingresar un Nro de Valor cuando carga Efectivo", vbExclamation, "Mensaje ..."
        txtAlta(3).Text = "CH"
        txtAlta(4).Text = "CHEQUE"
        
       ' Exit Sub
    End If

 
      
        If RBIngresoEgresoCaja(0).Value And Me.RBDebeHaber(0) Then       ' ingreso y debito entonces tomo el cliente
           ' vcliprovee = Left(Me.txtAlta7.Text, 25)
           ' vCodigoCliente = Left(Me.txtAlta7.Text, 25)
        End If
        
        
        If RBIngresoEgresoCaja(1).Value And Me.RBDebeHaber(1) Then
            'vcliprovee = Left(Me.txtAlta7.Text, 25)
            'vCodigoProveedor = Left(Me.txtAlta7.Text, 25)
        End If
    
    If Not valImprimeRecibo Then Exit Sub
    
    GuardarRenglon
    
    If Me.txtAlta(0).Text = "TR" Then limpiarTransferencia
    
    
    Me.txtAlta(12).SetFocus
    If Not txtAlta(0).Text = "TR" Then txtAlta(12).Text = ""
     
    
    Me.KlexMovimientoCaja.Row = Me.KlexMovimientoCaja.Rows - 1
    
    Me.KlexMovimientoCaja.TopRow = Me.KlexMovimientoCaja.Rows - 1
    

    vultimo_ingreso = Val(Me.txtAlta(13).Text)
   
   actualizarTotales
   
   
   

If Err Then
        Exit Sub
        GrabarLog "cmdAgregar_Click", Err.Number & " " & Err.Description, Me.Caption
End If

End Sub

Public Sub actualizarTotales()
On Error Resume Next
Dim i As Integer
Dim vValor As Double


With KlexMovimientoCaja
        For i = 1 To .Rows
        
            If Not .TextMatrix(i, 5) = "*" Then
                    
                    If .TextMatrix(i, 8) = "D" Then
                        vValor = vValor + .TextMatrix(i, 9)
                    End If
                
                    If .TextMatrix(i, 8) = "H" Then
                        vValor = vValor - .TextMatrix(i, 9)
                    End If
            End If
        
        
        Next
End With

vValor = Abs(vValor)

Me.ltotal.Caption = Format(vValor, "###,###,##0.00")
Me.ltotal.Tag = vValor

Me.lfalta.Caption = Format(vtotalcontrol - vValor, "###,###,##0.00")
Me.lfalta.Tag = vtotalcontrol - vValor

If Err Then Exit Sub
End Sub

Private Sub limpiarTransferencia()

Dim i As Integer

For i = 5 To 9
    txtAlta(i).Text = ""
    txtAlta(i).Tag = ""
Next

Me.vNuevaCustodiaCodigo.Text = ""
Me.VNuevaCustodiaNombre.Text = ""
Me.vCodBanco.Text = ""
Me.vDesBanco.Text = ""

End Sub

Private Sub cmdBorrar_Click()
On Error Resume Next

    If MsgBox("Esta seguro que desea borrar un registro?", vbYesNo + vbInformation, "Mensaje ...") = vbYes Then
        
        With KlexMovimientoCaja
            
            'Este registro es por si esta modificando
           If True Then ' If .TextMatrix(.Row, 1) = "" Then ' Ale: Revisar
                If Not .Rows = 2 Then
                    .RemoveItem (.Row)
                    Call FormatoGrillaCaja(.Rows - 1, True)
                Else
                    FormatoGrillaCaja (1)
                End If
            Else
                MsgBox "No puede borrar un  registro ya registrado en la Contabilidad", vbExclamation, "Mensaje ..."
            End If
        
        End With
    
    End If
    
If Err Then GrabarLog "cmdBorrar_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cmdCerrar_Click()
On Error Resume Next

    Unload Me
    
If Err Then GrabarLog "cmdCerrar_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub cmdCerrar2_Click()
Unload Me
End Sub

Private Sub cmdContribuyente_Click()
Call fbuscarGrilla("personas", "nombre", "id_personas", Me.vcliprovee.Name, Me, "apellido", True)   ' ema:
vcp = "contri"
End Sub

Private Sub cmdCheque_Click()
On Error Resume Next
'completarDatosCheques
frmChequesAlta.Tag = Me.Caption ' le indico a frmcheques desde donde lo llamo para completar los datos

'pasar por referencia
frmChequesAlta.Show

If Err Then GrabarLog "cmdCheque_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub complerarDatosCheques()
With gbldsCheques

.Codigo = 0
.fecha = Me.dtpFecha
.FechaAcreditacion = Me.dtpValor
.FechaDeposito = Me.dtpValor
.monto = Me.txtAlta(12).TxT
.Ncheque = Me.txtAlta(5)
.NroInterno = Me.txtAlta(2).TxT
.TipoMovimiento = Me.txtAlta(0).TxT
End With
End Sub

Private Sub cmdEventuales_Click()
Call fbuscarGrilla("(select * from proveedores where tipoproveedor='Eventuales') as p", "Nombre", "Codigo", Me.vcliprovee.Name, Me, , False)  ' ema:
vcp = "p"
End Sub

Private Sub guardarBoton()
On Error Resume Next

Dim vvnrointerno As Long


If Not valCajaCerradaMovCaja(Me.dtpFecha.Value) Then
        validado = False
        Exit Sub
End If


If (Me.txtAlta(0) = "AD") Or (Me.txtAlta(0) = "AC") Then
    If Not UCase(InputBox("Ingrese clave para hacer este movimiento", "Clave")) = "DALAS" Then
        Exit Sub
    End If
End If


If Me.KlexMovimientoCaja.TextMatrix(1, 2) = "" Then
    MsgBox "No hay movimientos ingresados", vbInformation
    validado = False
    Exit Sub
End If



If Val(Me.txtAlta(2).Text) > 0 And Me.chkNroInternoFijo.Value Then
    If MsgBox("Quiere mantener el nro interno ingresado manualmente ?", vbYesNo) = vbYes Then
        vvnrointerno = Val(Me.txtAlta(2).Text)
    End If
Else
    vvnrointerno = UltimoNroInterno2 + 1
    Me.txtAlta(2).Text = vvnrointerno
End If

initRollbk (vvnrointerno)

 ' ------------ verifica nro interno ----------------------
 If existeRegistro(Val(txtAlta(2))) Then Exit Sub
 '----------------------------------------------------------

vnrointerno = Val(txtAlta(2))

    If Validar = True Then
           
    g3.Visible = False
    
    Call setNroCorrelativos
  
  
    If Not (Me.txtAlta(0).Text = "TR" Or Me.txtAlta(0).Text = "VL" Or Me.txtAlta(0).Text = "AJ") Then
        If Me.chkIEautomatico.Value = 1 Then
            Call GenerarAsientoAutomaticamente 'Alefredo: con esto activas el guardado del asiento nuevo
        Else
            Call GuardarAsiento 'Ale: guarda el asiento como si lo estuviera haciendo del mòdulo de asiento. Muy mal
        End If
    End If
        '---------------------------------------------------------------------
        Call Guardar '  guarda el movimiento en el módulo caja / banco
        '---------------------------------------------------------------------
        
        Me.txtAlta(2).Text = vvnrointerno + 1
        
    Else
        txtAlta(3).SetFocus
        Exit Sub
    End If
     
   ' Call CarfarAsientos
    
actualizarGrilla ("")

g3.Visible = True

' Call endRollbk(vvnrointerno)

If Err < 0 Then

   ' Call endRollbk(vrollbk_nrointerno, vrollbk_nroasiento)
    
    g3.Visible = True
    'MsgBox "Cuidado !. Verifique si este movimientos fue cargado correctamente.", vbCritical
    GrabarLog "cmdGuardar_Click", Err.Number & " " & Err.Description, Me.Caption
Else
  '  Call endRollbk(vrollbk_nrointerno, vrollbk_nroasiento)
   
End If
End Sub

Private Sub actualizarGrilla(vcondi As String)
On Error Resume Next
Dim vr As New ADODB.Recordset
Dim vr2 As New ADODB.Recordset

Dim vsql, vsql2 As String

If vcondi = "" Then vcondi = " 1=1 "

vsql = "select Fecha, Codigo, Descripcion, format(inicial,2) as SaldoInicial, format(ingreso,2) as Ingresos,format(retiro, 2) as Egresos, format(t_logcaja.saldo, 2) as Saldo, DATE_FORMAT(momento,'%d %b %Y %T:%f') as Tiempo, nrointerno from t_logcaja " + _
" inner join bancos on t_logcaja.codigo = bancos.idbancos " + _
" where not  codigo = '' and " + vcondi + " order by id desc"

Call vr.Open(vsql, ConnDDBB, adOpenStatic, adLockPessimistic)

 Set gultimos.DataSource = vr.DataSource
 
 Me.gultimos.ColWidth(0) = 500
 
 Me.gultimos.ColWidth(1) = 1000
 
 Me.gultimos.ColWidth(2) = 1000
 
 Me.gultimos.ColWidth(3) = 3000
 
 Me.gultimos.ColWidth(4) = 1000
 
 Me.gultimos.ColWidth(5) = 1000
 
 Me.gultimos.ColWidth(6) = 1000
 
 Me.gultimos.ColWidth(7) = 1000
 
 Me.gultimos.ColWidth(8) = 3500
 
 Me.gultimos.ColWidth(9) = 1000
 
 vsql2 = "select ingreso,retiro from t_logcaja where not  codigo = '' and " + vcondi + " order by id desc"
 Call vr2.Open(vsql2, ConnDDBB, adOpenStatic, adLockPessimistic)
 
'Set Me.MSChart1.DataSource = Nothing
'Me.MSChart1.Refresh
 
'Set Mantenimiento.rslogcaja.DataSource = vr2.DataSource
'Set Me.MSChart1.DataSource = vr2.DataSource

 'Me.MSChart1.Refresh

If Err Then Exit Sub
End Sub

Function valDatosObligatorioOld() As Boolean
Dim vmensaje As String

valDatosObligatorioOld = True

If txtAlta(12).Text = "" Then vmensaje = vmensaje + "- Debe ingresar el importe de la operacion." + Chr(13)


If Not txtAlta(0).Text = "VL" And txtAlta10.Text = "" And Not Me.txtAlta(0).Text = "TR" And Not Me.txtAlta(0).Text = "ADT" Then
        vmensaje = vmensaje + "- Debe ingresar una Cta contable desde la solapa de <Cuenta>." + Chr(13)
End If
' If Me.vrendicion.Text = "" Then vmensaje = vmensaje + "- Debe ingresar una Rendición." + Chr(13)


If txtAlta(3).Text = "" Then vmensaje = vmensaje + "- Debe ingresar un Tipo de Valor." + Chr(13)

If Not vmensaje = "" Then
    MsgBox vmensaje, vbCritical
    valDatosObligatorioOld = False
End If


End Function
Function valDatosObligatorio() As Boolean
Dim vmensaje As String

Dim vValor As Long

valDatosObligatorio = True

If txtAlta(12).Text = "" Then vmensaje = vmensaje + "- Debe ingresar el importe de la operacion." + Chr(13)

If txtAlta(3).Text = "" Then vmensaje = vmensaje + "- Debe ingresar un Tipo de Valor." + Chr(13)


 vValor = Val(Trim(Me.txtAlta10.Text)) + Val(Trim(Me.vCodBanco)) + Val(Trim(Me.txtAlta6))

If Not vValor > 0 Then
    MsgBox "Cuenta de Caja - Banco erroneas", vbCritical
    valDatosObligatorio = False
End If


If Not vmensaje = "" Then
    MsgBox vmensaje, vbCritical
    valDatosObligatorio = False
End If


End Function

Private Sub CarfarAsientos()
If vConfigGral.vIncluyeContabilidad = True Then
        
        
        With frmAsientosAlta
            .txtCuentaVieneDe.Text = Me.Caption
            .txtCuentaVieneDe.Tag = Me.vcliprovee.Tag
            .dtpFecha.Value = Me.dtpFecha
            
            .vcliprovee.Tag = Me.vcliprovee.Tag
            .vcliprovee.Text = Me.vcliprovee.Text
            
           ' .chkControlar.Value = xtpChecked
            
           ' If Not opTipoDocumento(7).Value = True Then
           '     .txtImporteVieneDe.Text = Trim(txtTotal.Text)
           ' Else
           '     .txtImporteVieneDe.Text = Trim(txtIB(10).Text)
           ' End If
            
            '.cboTipoMovimiento.Tag = txtTipoMovimiento(0).Text
            '.cboTipoMovimiento.Text = txtTipoMovimiento(1).Text
            
           .lblNroInterno.Caption = Me.txtAlta(2)
        
            '.vVieneTabla = "PFactura"
            '.vVieneIdNombre = "idPfactura"
            '.vVieneIdValor = vIdPFactura
            
            .vidpersonas = vidpersonas
            .vidproveedores = vidproveedores
            .vidclientes = vidclientes
            
            ' ---------------- mas datos del asiento -----------
            '.vCodigoCliente = Me.txtProveedor(0).Text
            '.vCodigoProveedor = txtProveedor(0).Tag
            '----------------------------------------------------
            .txtImporteVieneDe.Text = vtotal
            
        
            .Show
            .ZOrder (0)
            .SetFocus
        End With
    End If
End Sub

Private Sub GenerarAsientoAutomaticamente()
'Alfredo: aca tenes una linea por cada renglon del asiento. Antes establecer el nrodelasiento
'Dim vnroasiento As Long  ' paso 1 para asiento
'vnroasiento = Val(GenerarDato("SELECT MAX(Numero) as NroAsiento FROM Asientos where balance=" + Str(vnrobalance), "NroAsiento")) + 1 ' paso 2 para asiento

vnroasiento = Val(GenerarDato("SELECT MAX(Numero) as NroAsiento FROM Asientos", "NroAsiento")) + 1 ' paso 2 para asiento

'abrirAsiento (vnroasiento) ' paso 3 para asiento

' Alfredo: Paso 3-4. En el caso que haya diferentes tipos de asientos dependiendo de un dato que està en la interfaz
' hay que poner un if para cada uno de ellos

'If Me.txtAlta(3) = "CH" Then
    'paso 4 para asiento
'End If

'If not Me.txtAlta(3) = "" Then
    'paso 4 para asiento
'End If



If Me.RBIngresoEgresoCaja(0).Value = 1 Then ' si es un ingreso
' paso 4 para asiento

' 1) Depósito en efvo.: Banco XX c/c   a Caja

If Me.txtAlta(3) = "EF" Then
    'paso 4 para asiento
    Call nuevoRenglonAsiento(vnroasiento, Me.dtpFecha, 0, Me.txtAlta(2), "", 1, bancoToCuenta(Me.txtAlta6), Me.txtAlta(12), 0, Me.vcliprovee.Tag, vcp)
    Call nuevoRenglonAsiento(vnroasiento, Me.dtpFecha, 0, Me.txtAlta(2), "", 1, "00304", Me.txtAlta(12), 0, Me.vcliprovee.Tag, vcp)
End If



' 2) Depósito de valores: Banco XX c/ca Valores a Depositar
If Me.txtAlta(3) = "CH" Then
    'paso 4 para asiento
    Call nuevoRenglonAsiento(vnroasiento, Me.dtpFecha, 0, Me.txtAlta(2), "", 1, bancoToCuenta(Me.txtAlta6), Me.txtAlta(12), 0, Me.vcliprovee.Tag, vcp)
    Call nuevoRenglonAsiento(vnroasiento, Me.dtpFecha, 0, Me.txtAlta(2), "", 1, "98884", Me.txtAlta(12), 0, Me.vcliprovee.Tag, vcp)
End If

End If


If Me.RBIngresoEgresoCaja(1).Value = 1 Then ' si es un egreso
'acá van los asientos de egreso
' renglones del asiento
'nuevoRenglonAsiento(vnroAsiento,

End If

End Sub
Private Function Validar() As Boolean
On Error Resume Next

    Dim vTotalD As Double, vTotalH, vTotalCtaH, vTotalCtaD As Double, l As Integer
    Dim vValor As Double
    Dim vmensaje As String
    
    vTotalCtaH = 0
    vTotalCtaD = 0
    vTotalH = 0
    
    
    If Val(txtAlta(2).Text) = 0 And vDatosEmpresa.UsarNroInterno = "SI" Then ' Alfredo: (Nro Interno) Esto tenes que poner cada vez que se controle el tema del nro interno
        MsgBox "No puede ingresar un Movimiento de Caja sin un Nro Interno!", vbExclamation, "Mensaje ..."
        Validar = False
        Exit Function
    End If

    vTotalD = 0
    vTotalH = 0
    
    With KlexMovimientoCaja
        For l = 1 To Val(.Rows - 1)
            
            If .TextMatrix(l, 8) = "D" Then
                
                        If Not .TextMatrix(l, 5) = "*" Then
                            vTotalD = vTotalD + Val(.TextMatrix(l, 9))
                            
                        End If
                        
                        If Not .TextMatrix(l, 7) = "" Then
                            vTotalCtaD = vTotalCtaD + Val(.TextMatrix(l, 9))
                        End If
            Else
                        
                        If Not .TextMatrix(l, 5) = "*" Then
                            vTotalH = vTotalH + Val(.TextMatrix(l, 9))
                        End If
                        
                        If Not .TextMatrix(l, 7) = "" Then
                            vTotalCtaH = vTotalCtaH + Val(.TextMatrix(l, 9))
                        End If
            End If
        Next
    End With

   Validar = True
   validado = True



    If Me.txtAlta(0).Text = "VL" Then
        vTotalD = vTotalD
        vTotalH = 0
    End If

    If Not Abs(Abs(vTotalD - vTotalH) - Val(Me.vtotalcontrol)) <= 1 And Val(Me.vtotalcontrol) > 0.1 Then

        MsgBox ("No coincide el total de control")
            Validar = False
            validado = False
        
        Exit Function
     
    End If

vValor = 0
vValor = vTotalCtaH + vTotalCtaD - vTotalH - vTotalD


If Int(vValor) > 1 Then
             vmensaje = " a favor Ctas contable = "
               Validar = False
            validado = False
End If

If Int(vValor) < -0.1 Then
    vmensaje = " a favor Caja-Bancos = "
                Validar = False
                 validado = False
End If


If Not Me.txtAlta(0).Text = "AC" And Not Me.txtAlta(0).Text = "AD" And Not Me.txtAlta(0).Text = "TR" And Not Me.txtAlta(0).Text = "VL" And Not vValor = 0 And Not Trim(Format(Abs(vValor), "###,###,##0.00")) = "0.00" Then
            MsgBox "Hay diferencias entre los montos de cajas-banco y las cuentas contables. " + Chr(13) + _
            "Importe diferenciado : " + vmensaje + Format(Abs(vValor), "###,###,##0.00")

            Validar = False
            validado = False
Else
            Validar = True
            validado = True
End If
     
     
     
     If Me.txtAlta(0).Text = "TR" Then
            Validar = validarTransaccion()
            validado = Validar
     End If
     
     If Me.txtAlta(0).Text = "VL" Or Me.txtAlta(0).Text = "AD" Or Me.txtAlta(0).Text = "AC" Then
            Validar = True
            validado = True
     End If
     
  
 If Me.vobservacion = "" Then
    If MsgBox("Falta ingresar un concepto para este movimiento" + Chr(13) + " Continúa de todas manera ?", vbYesNo) = vbNo Then
       Validar = False
       validado = False
    End If
 
 End If
     
     
If Me.dtpFecha.Text = "" Then
        MsgBox "Falta ingresar fecha del movimiento "
       Validar = False
       validado = False
End If
    
If Err Then GrabarLog "Validar", Err.Number & " " & Err.Description, Me.Caption
End Function
Private Sub Guardar()
On Error Resume Next
    Dim vidBancos As Long
    Dim vsql, vcodcta As String

    Dim rsBancosMovimientos As New ADODB.Recordset, sqlBancosMovimientos As String, m As Integer, vImporteD As Double, vImporteH As Double
    
    sqlBancosMovimientos = "SELECT * FROM BancosMovimientos WHERE (NroInterno = " & Val(txtAlta(2).Text) & ")"
    
    With rsBancosMovimientos
        Call .Open(sqlBancosMovimientos, ConnDDBB, adOpenDynamic, adLockOptimistic)
        
        If .EOF = True Then
            
        Else
            If MsgBox("El Nro Interno que ha ingreso existe en la Base de Datos. Desea Reemplazarlo? ", vbInformation + vbYesNo, "Mensaje ...") = vbYes Then
                Call BorrarBase("BancosMovimientos WHERE (NroInterno = " & Val(txtAlta(2).Text) & ")", pathDBMySQL)
                

                'Esto no es muy aconsejable
            Else
                Exit Sub
            End If
        End If

        vrollbk_nrointerno = Val(txtAlta(2).Text)

        ' recorro la Grilla
        For m = 1 To Val(KlexMovimientoCaja.Rows - 1)  ' recorre la grilla de los movimientos seleccionados
            vImporteD = 0
            vImporteH = 0
    
            
            'If KlexMovimientoCaja.TextMatrix(m, 1) = "" Then

               
                If KlexMovimientoCaja.TextMatrix(m, 8) = "D" Then
                    vImporteD = Val(KlexMovimientoCaja.TextMatrix(m, 9))
                Else
                    vImporteH = Val(KlexMovimientoCaja.TextMatrix(m, 9))
                End If
                    
                If Not KlexMovimientoCaja.TextMatrix(m, 5) = "*" And (Not Trim(KlexMovimientoCaja.TextMatrix(m, 3)) = "" Or Not Trim(KlexMovimientoCaja.TextMatrix(m, 2)) = "" Or Not Trim(KlexMovimientoCaja.TextMatrix(m, 5)) = "") Then
                    .AddNew ' bancosmovimientos
                     vidBancos = Replace(KlexMovimientoCaja.TextMatrix(m, 5), "*", "")
                     'vidBancos = KlexMovimientoCaja.TextMatrix(m, 5)
                    
                     .Fields("nrocomprobante").Value = vnrocomprobante
        
        
                    .Fields("idBancos").Value = Replace(KlexMovimientoCaja.TextMatrix(m, 5), "*", "")
                    
                    
                    .Fields("idBancosCuentas").Value = Val(KlexMovimientoCaja.TextMatrix(m, 6))
                    .Fields("Fecha").Value = strfechaMySQL(KlexMovimientoCaja.TextMatrix(m, 4))
                    
                    
                    .Fields("Debito").Value = vImporteD
                    .Fields("Credito").Value = vImporteH
                    .Fields("Saldo").Value = 0
        
                    .Fields("Comentario").Value = EsNulo(KlexMovimientoCaja.TextMatrix(m, 10))
                    .Fields("NroCheque").Value = Val(KlexMovimientoCaja.TextMatrix(m, 3))
                    
                    .Fields("Comentario2").Value = Me.vobservacion
                    .Fields("Comentario3").Value = Me.vobservacion2
        
                    .Fields("TipoMovimiento").Value = EsNulo(KlexMovimientoCaja.TextMatrix(m, 1))
                    
                    .Fields("idTipomovimientoas").Value = txtAlta(0).Tag
                    
                    
                    .Fields("NroInterno").Value = Val(txtAlta(2).Text)
                    .Fields("NroAsiento").Value = Val(vnroasiento)
        
                    .Fields("idTipoValor").Value = EsNulo(KlexMovimientoCaja.TextMatrix(m, 2))
                    
                    .Fields("idpersonas") = vidpersonas
                    .Fields("idproveedores") = vidproveedores
                    .Fields("idclientes") = vidclientes
                    
                    .Fields("cp") = vcp
                    .Fields("ClienteProveedor") = Me.vcliprovee.Tag
                    .Fields("codpersona") = Me.vcliprovee.Tag
                    
                    .Fields("idCheques").Value = EsNulo(KlexMovimientoCaja.TextMatrix(m, 16))
                    
                   ' .Fields("idRendiciones").Value = EsNulo(KlexMovimientoCaja.TextMatrix(m, 17))
                    
                    
                    If Me.RBDebeHaber(1).Value Then  ' si es un pago con cheques
                        Call cambiarCajaCheque(.Fields("idCheques").Value, "098") ' saca de la caja al cheque en caso que sea un retiro
            
                    End If
                    
                    If Me.RBDebeHaber(0).Value Then  ' si es un pago con cheques
                        Call cambiarCajaCheque(.Fields("idCheques").Value, Replace(KlexMovimientoCaja.TextMatrix(m, 12), "*", ""))
                    End If
                    
                   ' If EsNulo(Me.KlexMovimientoCaja.TextMatrix(m, 18)) = True Then
                   '     Call updateNrocheque(Replace(KlexMovimientoCaja.TextMatrix(m, 5), "*", ""), Val(KlexMovimientoCaja.TextMatrix(m, 3)))
                   ' End If
            
            
                    Call cancelarVales2(m)
            
                    Call setLogCaja(strfechaMySQL(KlexMovimientoCaja.TextMatrix(m, 4)), Replace(KlexMovimientoCaja.TextMatrix(m, 5), "*", ""), vImporteD, vImporteH, Val(txtAlta(2).Text))

            
            
                    .Update
                    
                    vsql2 = "select max(idBancosMovimientos)  as c from bancosmovimientos"
                    vidBancoCaja = traerDatos2(vsql2, "c", pathDBMySQL)
                    
                End If
                
                
                '------ si estoy arreglando el movimiento en el banco no tengo que guardarlo en el cheque--------
                If vguardaArreglo Then Exit Sub
                '--------------------------------------------------------------------
                
                
                ' Ale: completar. hacer que pueda modificar el dato
                                    If Me.chkIEautomatico = 0 Then
                                    
                                            vcodcta = ""
                                            vcodcta = Trim(KlexMovimientoCaja.TextMatrix(m, 7))
                                            
                                            If Not vcodcta = "" Then Call EjecutarScript("INSERT INTO AsientosDetalle (nrobalance,Numero,Linea,CodigoCuenta,Debe,Haber, LeyendaBancoCaja) VALUES (" & Str(vnrobalance) & "," & vnroasiento & "," & m & ",'" & Trim(KlexMovimientoCaja.TextMatrix(m, 7)) & "'," & vImporteD & "," & vImporteH & ",'" & Trim(KlexMovimientoCaja.TextMatrix(m, 10)) & "')")
                                                
                                            vcodcta = ""
                                            vcodcta = Trim(KlexMovimientoCaja.TextMatrix(m, 13))
                                                
                                            If Not vcodcta = "" Then Call EjecutarScript("INSERT INTO AsientosDetalle (nrobalance,Numero,Linea,CodigoCuenta,Debe,Haber, LeyendaBancoCaja) VALUES (" & Str(vnrobalance) & "," & vnroasiento & "," & m & ",'" & Trim(KlexMovimientoCaja.TextMatrix(m, 13)) & "'," & vImporteH & "," & vImporteD & ",'" & Trim(KlexMovimientoCaja.TextMatrix(m, 10)) & "')")
                                                
                                    End If
                                    
                If Val(KlexMovimientoCaja.TextMatrix(m, 11)) > 0 Then
                    Me.vidcheque = Val(KlexMovimientoCaja.TextMatrix(m, 11))
                End If

                Me.vidcheque = Val(KlexMovimientoCaja.TextMatrix(m, 11))
                
                ' --------- ahora guardo los cheques -------------
                
                Select Case EsNulo(KlexMovimientoCaja.TextMatrix(m, 2))
            
                    Case "CH"
                        ' Ale: falta aca hay que reemplazar todos los datos con los del dataset
                       ' dsToTabla ' pasa del ds a la tabla  - no lo estoy utilizando
                       
                       'controlo si el cheque fue seleccionado de la cartera de cheques. Esto ocurre cuando el datagrid tiene el idcheques
                        
                        Me.vidcheque = Val(KlexMovimientoCaja.TextMatrix(m, 11))
                       
                        If Me.vidcheque > 0 Then
                            ' cambia la  custodia al cheque
                            vsql2 = "idCustodia='" + Replace(KlexMovimientoCaja.TextMatrix(m, 12), "*", "") + "', idBancoCaja=" + Str(vidBancoCaja)
                            
                            vsql = "update cheques set " + vsql2 + " where idCheques=" + Str(Me.vidcheque)
                            Call EjecutarScript(vsql, pathDBMySQL)
                        End If
                        
                
                        'Call EjecutarScript("INSERT INTO Cheques (idEstadoCheque, Fecha, NCheque, Monto, NroInterno, Observaciones, FechaAcreditacion, TipoMovimiento) VALUES (1,'" & strfechaMySQL(dtpFecha.Value) & "', " & Val(KlexMovimientoCaja.TextMatrix(m, 3)) & "," & Val(KlexMovimientoCaja.TextMatrix(m, 9)) & "," & Val(txtAlta(2).Text) & ",'" & Me.KlexMovimientoCaja.TextMatrix(m, 10) & "','" & strfechaMySQL(KlexMovimientoCaja.TextMatrix(m, 4)) & "','" & Trim(txtAlta(0).Text) & "')")
                
                    Case "EF"
                
                    Case "PA"
                
                    
                    Case ""
                        'MsgBox "No eligio ningun movimiento!!!", vbExclamation, "Mensaje ..."
                
                End Select
           ' Else
                'Este ya existe
           ' End If
            
            
               If Me.KlexMovimientoCaja.TextMatrix(m, 1) = "ADT" Or Me.KlexMovimientoCaja.TextMatrix(m, 1) = "VL" Then
                    Call guardarAdelanto(m)  ' Adelantos para personas que no son proveedores
               End If
               
                 
        Next
    
    End With
    
    'Call BancoYCaja(Val(txtAltaCaja(2).Text))
                
    
 
   ' Call LimpiarCampos(1)
    
    sqlBancosMovimientos = ""

    If rsBancosMovimientos.State = 1 Then
        rsBancosMovimientos.Close
        Set rsBancosMovimientos = Nothing
    End If
    
    ' actualiza marca de vales cancelados

    'cancelarVales

If Err < 0 Then

   ' vrollbk = True
    GrabarLog "GuardarBancoIE", Err.Number & " " & Err.Description, Me.Caption

End If

End Sub

Private Sub cancelarVales2(m)
On Error Resume Next
Dim vid As Long

vid = Val(Me.KlexMovimientoCaja.TextMatrix(m, 3))

If Me.KlexMovimientoCaja.TextMatrix(m, 2) = "VAL" And vid > 0 Then
Dim vsql As String

vsql = "update  bancosmovimientos set conciliado = '-', idtipoValor= 'VA', tipomovimiento = 'VL' where idbancosmovimientos = " + Str(vid)
Call EjecutarScript(vsql)


End If

If Err Then Exit Sub
End Sub

Private Sub cancelarVales()
On Error Resume Next

Dim vsql As String

vsql = "update  bancosmovimientos set conciliado = '-', idtipoValor= 'VA', tipomovimiento = 'VL' where conciliado = 'cancelado'"
Call EjecutarScript(vsql)

If Err Then Exit Sub
End Sub


Private Sub setLogCaja(vfecha As String, vcodigo As String, vd As Double, vh As Double, Optional vnrointerno As Long)
On Error Resume Next

Dim vsql, vcampo, vvalores As String
Dim vInicial, vsaldo As Double

If vcodigo = "" Then Exit Sub

vsql = "select sum(t.Debito) - sum(t.Credito) as saldo from bancosmovimientos t" + _
" where t.idBancos = '" + vcodigo + "' Group By t.idBancos"

vInicial = Val(traerDatos2(vsql, "saldo", pathDBMySQL))

vsaldo = vInicial + vd - vh

vcampo = "fecha,codigo,inicial,ingreso,retiro,saldo, nrointerno"
vvalores = "'" + vfecha + "','" + vcodigo + "'," + Str(vInicial) + "," + Str(vd) + "," + Str(vh) + "," + Str(vsaldo) + "," + Str(vnrointerno)

vsql = "insert into t_logcaja (" + vcampo + ") values (" + vvalores + ")"

Call EjecutarScript(vsql, pathDBMySQL)

If Err Then Exit Sub
End Sub


Private Sub guardarVale(i As Integer)
' guarda un adelanto en pcuentascorrientes
Dim vsql, vcampos, vvalores As String


vcampos = "fecha,codigo,nombre,debito,credito,comentario,tipomovimiento,nrointerno"

With Me.KlexMovimientoCaja
        vvalores = "'" + strfechaMySQL(Me.dtpFecha.Value) + "','" + _
        .TextMatrix(i, 14) + "','" + _
        Replace(vcliprovee.Text, "''", "") + "'," + _
        "0," + _
        .TextMatrix(i, 9) + "," + _
        "'Adelantos por: " + .TextMatrix(i, 10) + "','" + _
        Me.txtAlta(0) + "'," + Me.txtAlta(2).Text
End With

Call settabla("pcuentascorrientes", vcampos, vvalores)


 
End Sub

Private Sub guardarAdelanto(i As Integer)
' guarda un adelanto en pcuentascorrientes
Dim vsql, vcampos, vvalores As String

If Not esProveedor(KlexMovimientoCaja.TextMatrix(i, 14)) Then Exit Sub

vcampos = "fecha,codigo,nombre,debito,credito,comentario,tipomovimiento,nrointerno"

With Me.KlexMovimientoCaja
        vvalores = "'" + strfechaMySQL(Me.dtpFecha.Value) + "','" + _
        .TextMatrix(i, 14) + "','" + _
        Replace(vcliprovee.Text, "''", "") + "'," + _
        "0," + _
        .TextMatrix(i, 9) + "," + _
        "'Adelantos por: " + .TextMatrix(i, 10) + "','" + _
        Me.txtAlta(0) + "'," + Me.txtAlta(2).Text
End With

Call settabla("pcuentascorrientes", vcampos, vvalores)


 
End Sub

Private Sub dsToTabla()
Dim vvsql, sqlCampos, sqlValores As String

'sqlCampos = "idCheques, idEstadoCheque, Fecha, Codigo, Nombre, idBancos, idBancosCuentas, Ncheque, Firmante, cp, FechaDeposito, Monto, Endoso, Remito, NroInterno, Observaciones, FechaAcreditacion, Foto, TipoMovimiento, TimeStamp"
sqlCampos = "idCheques, idEstadoCheque, Fecha, Codigo, Nombre, idBancos, idBancosCuentas, Ncheque, Firmante, FechaDeposito, Monto, Endoso,NroInterno, Observaciones, FechaAcreditacion,TipoMovimiento"


With gbldsCheques
    'sqlValores = Str(p.idCheques) + "," + Str(.idEstadoCheque) + "," + strfechaMySQL(.fecha) + "," + Str(.Codigo) + "," + Str(.Nombre) + "," + Str(.idBancos) + "," + Str(.idBancosCuentas) + "," + Str(.Ncheque) + "," + Str(.FirmanteStr) + "," + Str(.CP) + "," + strfechaMySQL(.FechaDeposito) + "," + Str(.monto) + "," + Str(.Endoso) + "," + Str(.remito) + "," + Str(.NroInterno) + "," + Str(.Observaciones) + "," + strfechaMySQL(.FechaAcreditacion) + "," + Str(.Foto) + "," + Str(.TipoMovimiento) + "," + Str(.TimeStamp)
     sqlValores = Str(.idCheques) + "," + Str(.idEstadoCheque) + "," + strfechaMySQL(.fecha) + "," + Str(.Codigo) + "," + Str(.Nombre) + "," + Str(.idBancos) + "," + Str(.idBancosCuentas) + "," + Str(.Ncheque) + "," + Str(.Firmante) + "," + strfechaMySQL(.FechaDeposito) + "," + Str(.monto) + "," + Str(.Endoso) + "," + Str(.NroInterno) + "," + Str(.Observaciones) + "," + strfechaMySQL(.FechaAcreditacion) + "," + Str(.TipoMovimiento)
End With

vvsql = ""
'"insert into cheques" + " (" + sqlCampos + ") value (" + sqlValores + ")"
 
Call EjecutarScript(Str(vvsql), pathDBMySQL)

End Sub

Private Sub GuardarAsiento()
On Error Resume Next
    
    Dim rsAsiento As New ADODB.Recordset, sqlAsiento As String
    Dim rsAsientoDetalle As New ADODB.Recordset, sqlAsientoDetalle As String
    
    sqlAsiento = "SELECT * FROM Asientos WHERE 1=2"
    
    'vnroasiento = Val(GenerarDato("SELECT MAX(Numero) as NroAsiento FROM Asientos where nrobalance=" + Str(vnrobalance), "NroAsiento")) + 1
    
    vnroasiento = Val(GenerarDato("SELECT MAX(Numero) as NroAsiento FROM Asientos", "NroAsiento")) + 1
    
    Debug.Print "Nro asiento: " + Str(vnroasiento)
    
    With rsAsiento
        .CursorLocation = adUseServer
        Call .Open(sqlAsiento, ConnDDBB, adOpenStatic, adLockOptimistic)
        
        If .EOF = True Then .AddNew
              
        .Fields("Fecha").Value = strfechaMySQL(dtpFecha.Value)
        .Fields("Numero").Value = vnroasiento
        .Fields("Leyenda").Value = vleyenda.Text
        .Fields("TipoMovimiento").Value = EsNulo(txtAlta(0).Text)
        .Fields("NroBalance").Value = vnrobalance
        .Fields("NroInterno").Value = Val(txtAlta(2).Text)
        
         'If vcliprovee.Text = "" Then
            
           ' MsgBox "No se resgistraron Clientes o Proveedores", vbCritical, "Atención !!!"
         
         'End If
         
        .Fields("codigocliente").Value = vCodigoCliente
        .Fields("codigoproveedor").Value = vCodigoProveedor
        
        .Fields("idrendiciones").Value = Me.vrendicion.Tag
        
        vrollbk_nroasiento = vnroasiento
        
        .Update
        
    End With
    
    sqlAsiento = ""

    If rsAsiento.State = 1 Then
        rsAsiento.Close
        Set rsAsiento = Nothing
    End If

Me.vleyenda.Text = ""

If Err < 0 Then
     
    'GrabarLog "GuardarAsiento", Err.Number & " " & Err.Description, Me.Caption
End If
End Sub

Private Sub cmdGuardar_Click()


Call guardarBoton

If Not validado = True Then Exit Sub

pagosAproveedores

Call terminoGrabarOk

Call endRollbk(Val(Me.txtAlta(2).Text))

Call LimpiarCampos(1)

Call PusActSaldo_Click

'Call invariantes_locales

End Sub


Private Sub invariantes_locales()
Dim vsql As String
Dim valor As Double

Exit Sub

vsql = "select Total from invcheques"

valor = traerDatos2(vsql, "Total", pathDBMySQL)

If valor > 0.1 Then
    
    If MsgBox("Debe conciliar los valores en cartera. " + Chr(13) + _
    "Esto ocurre cuando no coincide el saldo de la caja de valores con los documentos en cartera. " + Chr(13) + _
    " --- " + Chr(13) + _
    "Quiere realizarlo en este momento ?", vbYesNo) = vbYes Then
    
            Call chequesEnCartera(Me.Name, 0, "")
        
    End If
    
End If

Me.vsaldoValores.Caption = traerDatos2("select Total from saldovalores ", "Total", pathDBMySQL)

End Sub

Private Sub pagosAproveedores()

If Not (Me.RBDebeHaber(1) And Not Me.vcliprovee.Tag = "") Then
    Exit Sub
End If

MarcarDocumnetosPagos
 
MarcarDocumnetosPagosAuto

Unload frmBuscarFactura


If Not vtotalseleccionado = 0 And validarCtaCte Or vpagoPacial Then
    Call PagarCtaCteDirecto(Me.ltotal.Tag, Me.dtpFecha.Value, Me.vcliprovee.Tag, Me.vcliprovee.Text, Me.vobservacion.Text, Val(Me.txtAlta(2).Text))
Else
   ' MsgBox "Cuidado !!!! si usted está realizando un pago de factura a proveedores " + Chr(13) + "debe verificar la ctacte del proveedor para verificar su estado "
End If

End Sub

Function validarCtaCte() As Boolean
validarCtaCte = True

If Me.vcliprovee.Tag = "" Or Me.vcliprovee.Text = "" Then
    validarCtaCte = False
End If

End Function

Private Sub terminoGrabarOk()
   'Call validarRollbk(vrollbk_nrointerno, vrollbk_nroasiento)
   
   Call reparaValesCancelado
   Call ControlAsientoBCMovimiento(vnrointerno, Trim(Me.txtAlta(0).Text))
   
   
End Sub


Private Sub reparaValesCancelado()
On Error Resume Next
Dim vsql As String

    vsql = "update bancosmovimientos set conciliado= '' where conciliado = 'cancelado'"
    Call EjecutarScript(vsql, pathDBMySQL)
    
If Err Then Exit Sub
'GrabarLog "cmdLimpiar_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub


Private Sub cmdLimpiar_Click()
On Error Resume Next

    Call LimpiarCampos(0)

If Err Then GrabarLog "cmdLimpiar_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub cmdProveedor_Click()


End Sub

Private Sub Command1_Click()

If Not validarTrabEventuales Then Exit Sub

frmTrabEventuales.Show
frmTrabEventuales.vViene = Me.Name
frmTrabEventuales.tab.SelectedItem = 1

'frmTrabEventuales.tab.Selected = 1

End Sub


Public Function validarTrabEventuales() As Boolean
Dim vmen As String


validarTrabEventuales = True

     vmen = ""
     
     If EsNulo(txtAlta(3).Text) = "" Then vmen = vmen + Chr(13) + "- Tipo de valor"               '(Tipo Valor)
     
     If EsNulo(txtAlta10.Text) = "" Then vmen = vmen + Chr(13) + "- Cta. contable"             ' cta1
   
     If Me.txtAlta6.Text = "" Then vmen = vmen + Chr(13) + "Debe seleccionar una caja para asignar la extracción de dinero para el pago a Eventuales"
     
     If Not vmen = "" Then
        validarTrabEventuales = False
        MsgBox vmen, vbCritical, "Faltan datos:"
    End If

End Function


Private Sub chkNroInternoFijo_Click()
If chkNroInternoFijo Then
    Me.txtAlta(2).Enabled = True
    Me.txtAlta(2).SetFocus
End If
End Sub

Private Sub dtpFecha_Change()
dtpValor.Value = dtpFecha.Value
End Sub

Private Sub dtpFecha_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
       ' RBDebeHaber(0).SetFocus
       ' txtAlta(2).SetFocus
       ' txtAlta(2).SelStart = 0
       ' txtAlta(2).SelLength = Len(txtAlta(2).Text)
    End If
    
If Err Then GrabarLog "dtpFecha_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub dtpFecha_LostFocus()
On Error Resume Next
    dtpValor.Value = Me.dtpFecha.Value
    
'If Err Then GrabarLog "dtpFecha_LostFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub dtpValor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'Me.txtAlta6.SetFocus
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next

    'If KeyAscii = 13 Then Exit Sub
    
    
    

If Err Then GrabarLog "Form_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
    
     If KeyCode = vbKeyF7 Then
            Call pbCarga_Click(4)
     End If


    If KeyCode = vbKeyF6 Then
        Call cmdAgregar_Click
    End If
    
    
    If KeyCode = vbKeyF10 Then
        tab2.SelectedItem = 3
    End If
    
    
    If KeyCode = vbKeyF11 Then
        tab2.SelectedItem = 0
    End If
    
     
    If KeyCode = vbKeyF12 Then
        tab2.SelectedItem = 1
    End If
    
    

    If KeyCode = 13 Then SendKeys "{tab}"
    
    
    If KeyCode = vbKeyF1 Then Exit Sub
    
    If KeyCode = vbKeyF2 Then
        cmdGuardar_Click
    End If
    
    If KeyCode = vbKeyF4 Then
        Call PushButton3_Click
    End If
    
    
    If KeyCode = vbKeyF3 Then
        Me.txtAlta(12).SetFocus
    End If
    

    
    'If KeyCode = vbKeyF4 Then Exit Sub
    If KeyCode = vbKeyF5 Then
        cmdGuardar_Click
    End If
    If KeyCode = vbKeyF6 Then Exit Sub
    
    If KeyCode = vbKeyF7 Then
        If RBIngresoEgresoCaja(0).Value = True Then
            RBIngresoEgresoCaja(1).Value = True
        Else
            RBIngresoEgresoCaja(0).Value = True
        End If
    End If
    
    If KeyCode = vbKeyF8 Then
        If RBDebeHaber(0).Value = True Then
            RBDebeHaber(1).Value = True
        Else
            RBDebeHaber(0).Value = True
        End If
    End If
    
    
   
    If KeyCode = vbKeyF9 Then Exit Sub
    If KeyCode = vbKeyF10 Then Exit Sub
    If KeyCode = vbKeyF11 Then Exit Sub
    If KeyCode = vbKeyF12 Then Exit Sub
    
If Err Then GrabarLog "Form_KeyUp", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Load()
On Error Resume Next

   Call init
   
    With Me
        .Show
        .Left = 0
        .Top = 0
        .Height = 9450
        .Width = 16410
        .KeyPreview = True
    End With
    
    Call CentrarFormulario(Me)
    
    Call LimpiarCampos(1)
    
    Me.RBDebeHaber(0) = False
    Me.RBDebeHaber(1) = False
    
    
   ' Me.txtAlta(2).Text = UltimoNroInterno2 + 1
   
   txtAlta(3).Text = "EF"
   txtAlta(4).Text = "EFECTIVO"

   
   Me.dtpFecha.SetFocus
   
    



If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub init()

On Error Resume Next

    vpagoPacial = False
    
    TabAlta.SelectedItem = 0
    
    Call actualizarGrilla("")
    
    fvalores.Enabled = True
    
    Me.tab2.SelectedItem = 0
    
    bandNrocheque = False
    
    vCodigoCliente = ""
    vCodigoProveedor = ""
    
    vidpersonas = 0
    vidproveedores = 0
    vidclientes = 0
    
    vDraft = True
    
vnrobalance = TraerDato("balances", " Activo='S' order by NroBalance Desc", "NroBalance", pathDBMySQL)
Me.Caption = Me.Caption + "      [Nro. de Balance: " + Str(vnrobalance) + "]"


vguardaArreglo = False

Me.chkIEautomatico = CBool(LeerConfig(28)) ' variable global definida en el módulo global

Call FormatoGrillaCaja(1)
 
Me.dtpFecha.Value = Date

If Me.chkIEautomatico.Value = 1 Then
    Me.Frame1.Enabled = False
    Me.txtAlta10.Enabled = False
    Me.txtAlta11.Enabled = False
    Me.pbCarga(4).Enabled = False
Else
    Me.Frame1.Enabled = True
    Me.txtAlta10.Enabled = True
    Me.txtAlta11.Enabled = True
    Me.pbCarga(4).Enabled = True
End If

Me.vsaldoDisponible.Caption = Format(getSaldoDisponible, "###,###,###,##0.00")

Me.vsaldoValores.Caption = traerDatos2("select Total from saldovalores ", "Total", pathDBMySQL)

Call PusActSaldo_Click

Call invariantes_locales


If UCase(LeerXml("Puesto")) = "PRESTAMISTA" Then
    Me.txtAlta(0).Text = "TR"
    Me.txtAlta(0).Text = "Cambio Cheque"
End If


If Err Then
    'MsgBox "Problema al inicializar   " + Err.Description
    'End
End If

End Sub

Private Sub lbie_DragDrop(Source As control, x As Single, y As Single)

End Sub

Private Sub gultimos_SelChange()
On Error Resume Next
Dim vsql As String
Dim vr As Integer


vr = Me.gultimos.Row
 

vsql = " SELECT " + _
"  `bancosmovimientos`.`Fecha`, " + _
"  `asientosdetalle`.`CodigoCuenta`, " + _
" cuentas.cuenta, " + _
"  `asientosdetalle`.`Debe`, " + _
"  `asientosdetalle`.`Haber`, `bancosmovimientos`.`NroCheque`, " + _
"  `bancosmovimientos`.`codpersona` as CPersona, proveedores.nombre as Persona, " + _
"  `proveedores`.`Nombre`, bancosmovimientos.comentario2, bancosmovimientos.comentario  " + _
" FROM " + _
"  `bancosmovimientos` " + _
"  left JOIN `asientos` ON (`asientos`.`NroInterno` = `bancosmovimientos`.`NroInterno`) " + _
"  left JOIN `asientosdetalle` ON (`asientos`.`Numero` = `asientosdetalle`.`Numero`) " + _
"  left JOIN `cuentas` ON (`asientosdetalle`.`CodigoCuenta` = `cuentas`.`CodigoCuenta`) " + _
"  left JOIN `proveedores` ON (`bancosmovimientos`.`ClienteProveedor` = `proveedores`.`Codigo`) " + _
" where bancosmovimientos.nrointerno = " + Str(gultimos.TextMatrix(vr, 9))


    
Call LlenarGrilla2(Me.gdetalle, vsql, 7, pathDBMySQL)
If Err Then Exit Sub
End Sub

Private Sub KlexMovimientoCaja_KeyPress(KeyAscii As Integer)
On Error Resume Next

If Err Then Exit Sub
End Sub

Private Sub KlexMovimientoCaja_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
On Error Resume Next

If Err Then Exit Sub
End Sub

Private Sub Option1_Click()

    tab2.SelectedItem = 1
    Me.RBDebeHaber(0).Value = True
    
End Sub

Private Sub Option11_Click()
frame_doc.Enabled = True

txtvimporte_pagare.Text = vimporte_pagare

End Sub

Private Sub Option12_Click()
    frmCargarCodigoBarra.Show
End Sub

Private Sub Option2_Click()
f2.Enabled = True

vhaciendo = ""

End Sub

Private Sub Option3_Click()

f2.Enabled = False
tab2.SelectedItem = 0
Me.RBDebeHaber(1).Value = True

txtAlta6.Text = "1001"
txtAlta7.Text = "Efectivo"
    
End Sub

Private Sub Option4_Click()
    'Call goGastosExtrasCambioCheque
    
f2.Enabled = False
tab2.SelectedItem = 0
Me.RBDebeHaber(1).Value = True


txtAlta6.Text = "1001"
txtAlta7.Text = "Efectivo"

End Sub

Private Sub goGastoExtrasCambioCheque()
' poner una linea automática de movimento de caja
'
'
'
'

End Sub


Private Sub Option5_Click()
Call chequesEnCartera(Me.Name, Me.vcliprovee.Tag, Me.vcliprovee.Text)
frmCheques.Show
frmCheques.WindowState = vmaximizar
End Sub

Private Sub Option6_Click()
   ' Call Me.tab2.Item(0).Selected
    
f2.Enabled = False
tab2.SelectedItem = 0
Me.RBDebeHaber(1).Value = True


txtAlta6.Text = "1001"
txtAlta7.Text = "Efectivo"
    

End Sub

Private Sub Option7_Click()
 tab2.SelectedItem = 0
    Me.RBDebeHaber(1).Value = True
    
End Sub

Private Sub Pagare_Click()

drPagare.Show

End Sub

Private Sub pbCarga_Click(Index As Integer)
On Error Resume Next

   ' Call fbuscarGrilla("(select * from cuentas where Imputable ='S') as t", "Cuenta", "CodigoCuenta", "txtAlta(11)", Me)     ' ema:

Select Case Index

    Case 4
        
       If LeerXml("Puesto") = "Comuna" Or LeerXml("Puesto") = "Caja" Then
        If Me.RBDebeHaber(1).Value Then  ' egreso
           ' Call fbuscarGrilla("(select * from cuentas where Imputable ='S' and CodigoCuenta like '02.%') as t", "Cuenta", "CodigoCuenta", Me.txtAlta11.Name, Me)    ' ema:
           Call fbuscarGrilla("(select * from cuentas where Imputable ='S' and (tipo = 'E' or tipo is null) ) as t", "Cuenta", "CodigoCuenta", Me.txtAlta11.Name, Me)    ' ema:
            'Call fbuscarGrilla("(select * from cuentas where Imputable ='S') as t", "Cuenta", "CodigoCuenta", Me.txtAlta11.Name, Me)    ' ema:

        End If
    
        If Me.RBDebeHaber(0).Value Then ' ingreso
            'Call fbuscarGrilla("(select * from cuentas where Imputable ='S' and CodigoCuenta like '01.%') as t", "Cuenta", "CodigoCuenta", Me.txtAlta11.Name, Me)    ' ema:
            Call fbuscarGrilla("(select * from cuentas where Imputable ='S' and (tipo = 'I' or tipo is null) ) as t", "Cuenta", "CodigoCuenta", Me.txtAlta11.Name, Me)    ' ema:
           ' Call fbuscarGrilla("(select * from cuentas where Imputable ='S') as t", "Cuenta", "CodigoCuenta", Me.txtAlta11.Name, Me)    ' ema:
        
        End If
    
    Else
        Call fbuscarGrilla("(select * from cuentas where Imputable ='S') as t", "Cuenta", "CodigoCuenta", Me.txtAlta11.Name, Me)    ' ema:

    End If

    
    Case 2
            txtAlta(12).SetFocus
           Call fbuscarGrilla("(select * from bancos where not EsCaja ='B') as t", "Descripcion", "idBancos", Me.txtAlta7.Name, Me)     ' ema:
         
    Case Else
    
    vVuelveBusqueda = Me.Name
    vVieneBusqueda = pbCarga(Index).Tag

    frmBusqueda.Show
End Select


If Err Then GrabarLog "pbCarga_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub pbCarga2_Click(Index As Integer)
    Call fbuscarGrilla("(select * from cuentas where Imputable ='S') as t", "Cuenta", "CodigoCuenta", Me.vcta2.Name, Me)    ' ema:
End Sub

Private Sub PusAgregar_Click()

End Sub

Private Sub PusAceptar_Click()
Dim vultimo_importe As Double

Me.RBDebeHaber(1).Value = True

tab2.SelectedItem = 2

If Val(vcomiPorc.Text) > 0 Then
   Me.txtAlta(12).Text = vimporte_pagare * Val(Me.vcomiPorc.Text) / 100
End If

If Val(vcomiFijo.Text) > 0 Then
   Me.txtAlta(12).Text = Val(Me.vcomiFijo.Text)
End If

f2.Enabled = True

txtAlta10.Text = "1101"
vimporte_pagare = "Ganancias por comisión de Cheque"


End Sub

Private Sub PusActSaldo_Click()
On Error Resume Next
Dim i As Integer

Dim vr As New ADODB.Recordset
Dim vr2 As New ADODB.Recordset

Dim vsql, vsql2 As String


vsql = "select descripcion as Arqueo,  format(sum(t.Debito) - sum(t.Credito),'##,###,##0.00') as saldo   from bancosmovimientos t " + _
" inner join bancos b " + _
" on b.idBancos = t.idBancos where not t.idBancos='098' " + _
" group by t.idBancos "


Call vr.Open(spcierreTemp2grilla(Date), ConnDDBB, adOpenStatic, adLockPessimistic)

Set Me.gsaldos.DataSource = vr.DataSource
 
'set  Me.gsaldos.DataSource =
 
 
 Me.gsaldos.ColWidth(0) = 1800
 
 Me.gsaldos.ColWidth(1) = 1000
 
 For i = 1 To gsaldos.Rows - 1
     Me.gsaldos.TextMatrix(i, 1) = Format(Me.gsaldos.TextMatrix(i, 1), "###,###,##0.00")
 Next
  
If Err Then Exit Sub
End Sub

Private Sub PusActualizar_Click()
Me.vsaldoDisponible.Caption = Format(getSaldoDisponible, "###,###,###,##0.00")
End Sub

Private Sub PusBorrarMovimientos_Click()
On Error Resume Next
Dim i As Integer
Dim vValor As Long

i = Me.gultimos.Row
vValor = gultimos.TextMatrix(i, 9)

Call verTransacciones(vValor)


gultimos.RemoveItem (i)

If Err < 0 Then Exit Sub
End Sub

Private Sub PusBuscarCheque_Click()
On Error Resume Next

Call chequesEnCartera(Me.Name, Me.vcliprovee.Tag, Me.vcliprovee.Text)

If Err Then Exit Sub
End Sub

Private Sub PusBuscarDocumento_Click()
On Error Resume Next


If Not Me.RBDebeHaber(1).Value Then

    mensaje "Cuidado. Debe seleccionar la opción de comprobante de pago"
    
    Exit Sub
    
End If


    With frmBuscarFactura
            
        '    .vImporteSeleccionado.Tag = Me.TxtTotalAPagar.Tag
        '    .vImporteSeleccionado.Caption = Me.TxtTotalAPagar
           .Show
        If Trim(vcliprovee.Text) <> "" Then
            .txtCliente.Text = Me.vcliprovee.Text
            .txtCliente.Tag = Me.vcliprovee.Tag
            .cpFactura = "pfactura"  '"Factura"
        End If
        .CmdEjecutarCobro.Enabled = True
        
        
        .finit
        .vieneCobro = False
        '.Show
        .chkFechaTodas.Value = True
        .cmdFiltrar_Click
        .Show
        .WindowState = 0
         .viene = "ie"
         

         If Val(Me.txtAlta(12).Text) > 0 Then
           ' .vImporteSeleccionado.Tag = Me.TxtTotalAPagar.Tag
            '.vImporteSeleccionado.Caption = Me.TxtTotalAPagar
        End If
        
    End With
    
    'HabilitarControles (True)
    

If Err Then GrabarLog "PusBuscarDocumento_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub PusCancelarLos_Click()
frmBancoCajaDetalle.chkFecha = xtpChecked
'frmBancoCajaDetalle.vcliprovee.Tag = Me.vcliprovee.Tag
frmBancoCajaDetalle.vcliprovee.Text = "Selecciona una persona para cancelar VALES"
frmBancoCajaDetalle.vctm = "VL"
frmBancoCajaDetalle.RadPendientes.Value = True
'frmBancoCajaDetalle.cmdFiltrar_Click

If frmBancoCajaDetalle.estanCancelandoVale Then
    MsgBox "En estos momentos, en otro puesto se están cancelando Vales y no es posible hacerlo en simultaneo"
    Exit Sub
End If

frmBancoCajaDetalle.Show


End Sub

Private Sub PusCargarDatos_Click()

If RBDebeHaber(0).Value Then
    MsgBox "No se puede cargar datos de un cheque propio en Ingreso", vbInformation, "Operación no permitida"
    Exit Sub
Else

    vtipoCheque = "propio"
    Me.fvalores.Enabled = True
End If

End Sub

Private Sub PusCierreDe_Click()
    
   ' Load frmBancoCajaDetalle
    frmBancoCajaDetalle.Hide
    frmBancoCajaDetalle.vtipolistado = "Agrupado por Cajas-Bancos"
    Call frmBancoCajaDetalle.cmdFiltrar_Click
    Unload frmBancoCajaDetalle
End Sub

Private Sub PusComenzarCarga_Click()

End Sub

Private Sub PusDetalle_Click()
Call PusCierreDe_Click
End Sub

Private Sub PusGuardarSin_Click()

If Not MsgBox("Está seguro que desea guardar este movimiento de ajuste ?", vbYesNo, "Guardar Ajustes.") = vbYes Then Exit Sub

vguardaArreglo = True
Call Guardar '  guarda el movimiento en el módulo caja / banco
vguardaArreglo = False

Call LimpiarCampos(1)

End Sub

Private Sub PushButton1_Click()


If Not RBDebeHaber(1).Value = True Then

    frmChequesAlta.Show
    frmChequesAlta.vViene = "frmie.ingreso"
    
    vhaciendo = "Cheque_a_cambiar"

Else


    Call chequesEnCartera(Me.Name, Me.vcliprovee.Tag, Me.vcliprovee.Text)

End If
Exit Sub

frmCheques.Show
frmCheques.WindowState = vmaximizar
frmCheques.vViene = Me.Name

Me.fvalores.Enable = True

End Sub

Private Sub PushButton10_Click()
On Error Resume Next
Dim vr As New ADODB.Recordset
Dim vr2 As New ADODB.Recordset

Dim vsql, vsql2 As String


vsql = "select (100* year(t.fecha)+month(t.fecha)) as Meses, format(sum(t.Debito), '###,###,##0.00') as Ingresos , format(sum(t.credito), '###,###,##0.00') as Egresos,format(sum(t.Debito) - sum(t.Credito),'###,###,##0.00') as saldo   from bancosmovimientos t " + _
" inner join bancos b " + _
" on b.idBancos = t.idBancos " + _
" group by (100* year(t.fecha)+month(t.fecha)) "



Call vr.Open(vsql, ConnDDBB, adOpenStatic, adLockPessimistic)

 Set gultimos.DataSource = vr.DataSource
 
 Me.gultimos.ColWidth(0) = 1000
 
 Me.gultimos.ColWidth(1) = 1000
 
 Me.gultimos.ColWidth(2) = 1000
 
  Me.gultimos.ColWidth(3) = 1000
 
 Me.gultimos.ColWidth(4) = 1000
 
 
'Set Me.MSChart1.DataSource = Nothing
'Me.MSChart1.Refresh
 
If Err Then Exit Sub

End Sub

Private Sub PushButton11_Click()
    Call Imprimir(2)
End Sub

Private Sub PushButton12_Click()
    Call fbuscarGrilla("(select * from articulos order by Descrip) as t", "Descrip", "Codigo", Me.vobservacion.Name, Me)
End Sub

Private Sub PushButton13_Click()
Call fbuscarGrilla("(select * from articulos order by Descrip) as t", "Descrip", "Codigo", Me.vobservacion2.Name, Me)
End Sub

Private Sub PushButton14_Click()
    Call actualizarGrilla("")
End Sub

Private Sub PushButton15_Click()
   ' Call valirPagare
   
   Call llenarPagare
   
   
   drPagare.Show

   
End Sub

Private Sub llenarPagare()

Dim vcuerpo As String

With drPagare


    vcuerpo = .Sections("detalle").Controls("cuerpo").Caption
    
    

    vcuerpo = Replace$(vcuerpo, "%vdia%", Str(Me.dtpFecha))
    vcuerpo = Replace$(vcuerpo, "%vempresa%", vDatosEmpresa.Nombre)
    vcuerpo = Replace$(vcuerpo, "%vdomicilio%", vDatosEmpresa.Direccion)
    vcuerpo = Replace$(vcuerpo, "%vlocalidad%", vDatosEmpresa.Localidad)
    vcuerpo = Replace$(vcuerpo, "%vimporte%", Me.txtvimporte_pagare)
    vcuerpo = Replace$(vcuerpo, "%vinteres%", Me.txtvintereses_pagare.Text)
    
    .Sections("detalle").Controls("cuerpo").Caption = vcuerpo
    
    .Sections("sección4").Controls("evencimiento").Caption = Str(txtvencimiento)
    .Sections("sección4").Controls("eimporte").Caption = Str(Me.txtvimporte_pagare)

End With

End Sub


Private Sub PushButton16_Click()
    Call Imprimir(2)
End Sub

Private Sub PushButton2_Click()
    Call fbuscarGrilla("(select * from bancos where not  EsCaja = 'B') as t", "Descripcion", "idBancos", Me.VNuevaCustodiaNombre.Name, Me)
End Sub

Private Sub PushButton3_Click()

If Not vtipoCheque = "propio" Then
    
    Call fbuscarGrilla("(select * from bancos where EsCaja ='B')as t", "Descripcion", "idBancos", Me.vDesBanco.Name, Me)

Else

    Call fbuscarGrilla("(select * from bancos where EsCaja ='N')as t", "Descripcion", "idBancos", Me.vDesBanco.Name, Me)

End If


'Call fbuscarGrilla("conceptos2", "descripcion", "idconceptos", Me.vconcepto.Name, Me)   ' ema:
End Sub

Private Sub PushButton4_Click()
Call fbuscarGrilla("rendiciones", "nombre", "idrendiciones", Me.vrendicion.Name, Me)
End Sub

Private Sub PushButton5_Click()
    vtipoCheque = "tercero"
    Me.fvalores.Enabled = True
End Sub

Private Sub PushButton7_Click()
    frmBancoCajaDetalle.Hide
    frmBancoCajaDetalle.vtipolistado = "Saldos"
    Call frmBancoCajaDetalle.cmdFiltrar_Click
    Unload frmBancoCajaDetalle
End Sub

Private Sub PushButton8_Click()
    frmBancoCajaDetalle.Show
    frmBancoCajaDetalle.tabbc.SelectedItem = 2
End Sub

Private Sub PushButton9_Click()
On Error Resume Next
Dim vr As New ADODB.Recordset
Dim vr2 As New ADODB.Recordset

Dim vsql, vsql2 As String


vsql = "select t.idBancos , descripcion,  format(sum(t.Debito) - sum(t.Credito),'###,###,##0.00') as saldo   from bancosmovimientos t " + _
" inner join bancos b " + _
" on b.idBancos = t.idBancos " + _
" group by t.idBancos "



Call vr.Open(vsql, ConnDDBB, adOpenStatic, adLockPessimistic)

 Set gultimos.DataSource = vr.DataSource
 
 Me.gultimos.ColWidth(0) = 500
 
 Me.gultimos.ColWidth(1) = 1000
 
 Me.gultimos.ColWidth(2) = 4000
 
 Me.gultimos.ColWidth(3) = 1500
 
'Set Me.MSChart1.DataSource = Nothing
'Me.MSChart1.Refresh
 
If Err Then Exit Sub
End Sub

Private Sub PusImprimir_Click()
Call imprimirGrilla(Me.gultimos, 6)
End Sub

Private Sub PusImprimirComprobante_Click()
   ' If Not valImprimeRecibo Then Exit Sub
    
    Call llenarDrRecibo
End Sub

Function valImprimeRecibo() As Boolean
Dim vmensaje As String
valImprimeRecibo = True
vmensaje = ""

    If Not Abs(Val(Me.txtAlta(12).Text)) > 0 Then
        valImprimeRecibo = False
        vmensaje = "Falta importe de la operación"
    End If
    
    
    If Not vmensaje = "" Then MsgBox vmensaje, vbCritical, "Error de validación... "
    
End Function



Private Sub PusImprimirGuardar_Click()

    Imprimir (1)


Exit Sub

vDraft = True


'Call setNroCorrelativos

'Call llenarDrRecibo
Call guardarBoton

If Not validado = True Then Exit Sub

pagosAproveedores


Call terminoGrabarOk

If Not validado = True Then Exit Sub

Call llenarDrRecibo

Call setNroCorrelativosLimpiar

Call LimpiarCampos(1)

drRecibo.Show


End Sub

Private Sub Imprimir(Optional vcopia As Integer)
'Call setNroCorrelativos

'Call llenarDrRecibo
Call guardarBoton

If Not validado = True Then Exit Sub

pagosAproveedores

Call terminoGrabarOk

Call llenarDrRecibo

Call setNroCorrelativosLimpiar

Call LimpiarCampos(1)


If vcopia = 1 Then
    drRecibo.PrintReport False
End If

If vcopia = 2 Then
    drRecibo.PrintReport False
    drRecibo.PrintReport False
End If

Unload drRecibo

End Sub
Private Sub setNroCorrelativos()
 If Not vnrocomprobante > 0 Then
    vnrocomprobante = getNroRecibo
End If
End Sub

Private Sub setNroCorrelativosLimpiar()
 vnrocomprobante = 0
End Sub

Private Sub PusLimpiar_Click()
Me.vcliprovee.Text = ""
Me.vcliprovee.Tag = ""
End Sub

Private Sub PusPersonas_Click()

Me.vcliprovee.Tag = ""
Me.vcliprovee.Text = ""

Call fbuscarGrilla("proveedores", "Nombre", "Codigo", Me.vcliprovee.Name, Me, , False)
' ema:

vcp = "p"

End Sub

Private Sub PusSaldos_Click()
  ' Load frmBancoCajaDetalle
    frmBancoCajaDetalle.Hide
    frmBancoCajaDetalle.vtipolistado = "Saldos"
    Call frmBancoCajaDetalle.cmdFiltrar_Click
    Unload frmBancoCajaDetalle
End Sub

Private Sub PusSaldosPor_Click()
On Error Resume Next
Dim vr As New ADODB.Recordset
Dim vr2 As New ADODB.Recordset

Dim vsql, vsql2 As String



vsql = "select  t.fecha as Dia, format(sum(t.Debito), '###,###,##0.00') as Ingresos , format(sum(t.credito), '###,###,##0.00') as Egresos, format(sum(t.Debito) - sum(t.Credito),'###,###,##0.00') as saldo   from bancosmovimientos t " + _
" inner join bancos b " + _
" on b.idBancos = t.idBancos " + _
" group by t.fecha "



Call vr.Open(vsql, ConnDDBB, adOpenStatic, adLockPessimistic)

 Set gultimos.DataSource = vr.DataSource
 
 Me.gultimos.ColWidth(0) = 500
 
 Me.gultimos.ColWidth(1) = 1000
 
 Me.gultimos.ColWidth(2) = 1000
 
  Me.gultimos.ColWidth(3) = 1000
 
 Me.gultimos.ColWidth(4) = 1000
 

'Set Me.MSChart1.DataSource = Nothing
'Me.MSChart1.Refresh
 
If Err Then Exit Sub

End Sub

Private Sub PusSelConceptos_Click()
Call fbuscarGrilla("conceptos2", "descripcion", "idconceptos", Me.vconcepto.Name, Me)   ' ema:
End Sub

Private Sub PusValoresCartera_Click()
    Call chequesEnCartera(Me.Name, 0, "")
End Sub

Public Sub RBDebeHaber_Click(Index As Integer)
 If txtAlta(3).Text = "CH" And Me.RBDebeHaber(1).Value Then
            Me.gcustodia.Visible = True
        Else
            Me.gcustodia.Visible = False
End If

Select Case Index
    Case 1
        Call initIngreso
    Case 0
        Call initEgreso
End Select
End Sub

Private Sub txtAlta_Change(Index As Integer)
On Error Resume Next
Dim vsql, vsn As String
Dim vnrocheque As Long
   
    Select Case Index
    
        Case 0
        
        txtAlta(0).Tag = traerDatos2("select * from tipoMovimientos where codigo='" + Trim(txtAlta(0).Text) + "'", "idtipoMovimientos", pathDBMySQL)
        
        If txtAlta(0).Text = "VL" Then
            
            txtAlta(3).Text = "EF"
            txtAlta(4).Text = "EFECTIVO"
            
            Me.initIngreso
            
        End If
        
        Case 3
        
        If txtAlta(3).Text = "CH" And Me.RBDebeHaber(1).Value Then
            Me.gcustodia.Visible = True
            Me.tab2.SelectedItem = 1
        Else
            Me.tab2.SelectedItem = 0
            Me.gcustodia.Visible = False
        End If
        
        
        If txtAlta(3).Text = "CH" Then
            Me.tab2.SelectedItem = 1
        End If
        
        Case 6
    
        
        Case 7
        
        If Not Me.txtAlta7.Text = "" Then
            Me.vDesBanco.Tag = ""
            Me.vCodBanco.Text = ""
            Me.vDesBanco.Text = ""
        End If

   Case 8

        
        Case 10

            
    End Select
    

If Err Then GrabarLog "txtAltaCaja_Change", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub FormatoGrillaCaja(vCantidadRenglones As Integer, Optional vDejaRenglones As Boolean)
On Error Resume Next

    Dim i As Integer

    With KlexMovimientoCaja
        .FixedRows = 1
        .FixedCols = 1
    
        .Cols = 19
        .Rows = vCantidadRenglones + 1
        
        If vDejaRenglones = False Then
            If vCantidadRenglones = 1 Then
                For i = 0 To .Cols - 1
                    .TextMatrix(1, i) = ""
                    .ColWidth(i) = 0
                Next
            End If
        End If
        
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 100
        
        .TextMatrix(0, 1) = "idBancosMovimientos"
        .ColWidth(1) = 0
        
        .TextMatrix(0, 2) = "Tipo Valor"
        .ColWidth(2) = 850
               
        .TextMatrix(0, 3) = "Nro. Valor"
        .ColWidth(3) = 850
        
        .TextMatrix(0, 4) = "F. Valor"
        .ColWidth(4) = 1000
        
        .TextMatrix(0, 5) = "Bco/Caja"
        .ColWidth(5) = 3000
        
        .TextMatrix(0, 6) = "C. Banco"
        .ColWidth(6) = 500
                
        .TextMatrix(0, 7) = "Cta. Contable"
        .ColWidth(7) = 1700
        
        .TextMatrix(0, 8) = "D/H"
        .ColWidth(8) = 700
        
        .TextMatrix(0, 9) = "Importe"
        .ColWidth(9) = 1000
        .ColDisplayFormat(9) = "#0.00"
        
        .TextMatrix(0, 10) = "Obs"
        .ColWidth(10) = 3000
        
        .TextMatrix(0, 11) = "idCheque"
        .ColWidth(11) = 1000
        
        .TextMatrix(0, 12) = "NCustodia"
        .ColWidth(12) = 1000
        
        .TextMatrix(0, 13) = "Caja - Banco"
        .ColWidth(13) = 3000
                
        .TextMatrix(0, 14) = "idPer"
        .ColWidth(14) = 500
                
        .TextMatrix(0, 15) = "tPer"
        .ColWidth(15) = 500
        
        .TextMatrix(0, 17) = "idrendiciones"
        .ColWidth(17) = 100
        '.BackColorAlternate = &HC0C0C0

    End With
    
If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Caption
End Sub

Function ftrascomentario() As String

If Me.txtAlta(0) = "TR" Then

End If

End Function

Private Sub GuardarRenglon()
On Error Resume Next
    
    Dim i, j As Integer
    Dim vie As String
    
    
    vie = ""
    
    Me.vidcheque = gbldsCheques.vid
    
    With KlexMovimientoCaja
        
        If .Rows <= 2 And .TextMatrix(.Rows - 1, 8) = "" Then
            FormatoGrillaCaja (1)
        Else
            .Rows = .Rows + 1
        End If
        
        j = .Rows - 1
        
        
        
        If EsNulo(txtAlta(0).Text) = "VL" Then
            .TextMatrix(j, 1) = "EF"
        Else
            .TextMatrix(j, 1) = EsNulo(txtAlta(0).Text)
        End If
        
        
        .TextMatrix(j, 2) = EsNulo(txtAlta(3).Text)             '(Tipo Valor)
        .TextMatrix(j, 3) = EsNulo(txtAlta(5).Text)             '(Nro Valor)
        .TextMatrix(j, 4) = EsNulo(dtpFecha.Text)               '(F. Valor)
        
        
        If Not IsDate(.TextMatrix(j, 4)) Then
            MsgBox "Cuidado. La fecha es incorrecta.", vbInformation
        End If
        
        
        If Trim(EsNulo(txtAlta6.Text)) = "" Then              '(idBancos)
            
            .TextMatrix(j, 5) = "*" + Me.vNuevaCustodiaCodigo.Text
         
        Else
           ' .TextMatrix(j, 5) = "[" & EsNulo(txtAlta6.Text) & "]"
            .TextMatrix(j, 5) = "*" + txtAlta6.Text           ' caja banco
            .TextMatrix(j, 13) = EsNulo(txtAlta7.Text)         'vNuevaCustodiaCodigo
        
        End If
        
        .TextMatrix(j, 6) = EsNulo(txtAlta(8).Text)             '(CodigoCuenta)
        .TextMatrix(j, 7) = EsNulo(txtAlta10.Text)            '(CodigoCuenta contable)
        
        If RBDebeHaber(0).Value = True Then
            vie = "Ingreso. "
            .TextMatrix(j, 8) = "D"                              'Debe/Haber
            .TextMatrix(j, 14) = "H"
        
        Else
            vie = "Egreso. "
            .TextMatrix(j, 8) = "H"                                'Debe/Haber
            .TextMatrix(j, 14) = "H"
        
'            vltotal = vltotal  EsNulo(txtAlta(12).Text)
        End If
        
        
        If Val(txtAlta6.Text) > 0 Then
       '  vltotal = vltotal + Val(EsNulo(txtAlta(12).Text))
        End If
         
        .TextMatrix(j, 9) = EsNulo(txtAlta(12).Text)               'Importe
       
        
              
     '.TextMatrix(j, 10) = Trim(vie) + "Tipo: " + Me.txtAlta(3) + " - " + fcomenCaja + " " + fcomenCheque + " " + fcomenCta + txtAlta(13).Text            'Observaciones
             
        
        .TextMatrix(j, 10) = vie + fcomenCaja + " " + fcomenCheque + " " + fcomenCta + "  - " + txtAlta(13).Text             'Observaciones"
      
        
        .TextMatrix(j, 11) = vidcheque                             'idcheque
        
        
            If Me.RBDebeHaber(1) Then
                .TextMatrix(j, 12) = "*" + "098"         'vNuevaCustodiaCodigo
            Else
                .TextMatrix(j, 12) = "*" + EsNulo(Me.vNuevaCustodiaCodigo) 'vNuevaCustodiaCodigo
            End If
            
        
        .TextMatrix(j, 13) = EsNulo(Me.vcodgocta2.Text)             'Código de la cuenta para personas y entidades

        .TextMatrix(j, 14) = EsNulo(Me.vcliprovee.Tag)              'id persona persona
    
        .TextMatrix(j, 15) = vcp                                    'tipo de personas
        
        .TextMatrix(j, 16) = vIdCheques                             'tipo de personas
        
        .TextMatrix(j, 17) = Me.vrendicion.Tag                      'tipo de personas
        
        .TextMatrix(j, 18) = bandNrocheque
        
        
         If bandNrocheque Then
            Call updateNrocheque(Replace(KlexMovimientoCaja.TextMatrix(j, 5), "*", ""), Val(KlexMovimientoCaja.TextMatrix(j, 3)))
         End If
            
        
        
        
        bandNrocheque = False
    
    
    If txtAlta(0) = "VL" Then
    
    
            .Rows = .Rows + 1
        
             j = .Rows - 1
        
        .TextMatrix(j, 1) = EsNulo(txtAlta(0).Text)
        .TextMatrix(j, 2) = "VA"
        .TextMatrix(j, 3) = EsNulo(txtAlta(5).Text)             '(Nro Valor)
        .TextMatrix(j, 4) = EsNulo(dtpFecha.Text)               '(F. Valor)
        
      
        .TextMatrix(j, 5) = "*1001"           ' caja banco
     
            .TextMatrix(j, 8) = "D"                              'Debe/Haber
        
        .TextMatrix(j, 9) = EsNulo(txtAlta(12).Text)               'Importe
        
        .TextMatrix(j, 10) = "Ingreso Vale. - " + fcomenCheque + " " + txtAlta(13).Text + " Persona: " + Left(Me.vcliprovee.Text, 50)              'Observaciones
        
        .TextMatrix(j, 13) = EsNulo(Me.vcodgocta2.Text)             'Código de la cuenta para personas y entidades

        .TextMatrix(j, 14) = EsNulo(Me.vcliprovee.Tag)              'id persona persona
    
        .TextMatrix(j, 15) = vcp                                    'tipo de personas
        
        bandNrocheque = False
    
    
    End If


 End With
    
    limpiarChequesSeleccionados
    limpiarRenglon
    
    
    If Err Then GrabarLog "GrabarRenglon", Left(Err.Number & " " & Err.Description, 99), Me.Name
End Sub


Private Sub limpiarRenglon()
Dim i As Integer

Me.txtAlta(5).Tag = ""
Me.txtAlta(5).Text = ""

vCodBanco.Tag = 0
vCodBanco.Text = 0

'Me.txtAlta(10).Text = ""
'Me.txtAlta(11).Text = ""
Me.txtAlta(13).Text = ""
'Me.txtAlta(14).Text = ""
Me.txtAlta(5).Text = ""

Me.txtAlta7.Tag = ""
Me.txtAlta6.Tag = ""


Me.txtAlta7.Text = ""
Me.txtAlta6.Text = ""

'Me.txtAlta(10).Tag = ""
'Me.txtAlta(11).Tag = ""
Me.txtAlta(13).Tag = ""
Me.txtAlta(5).Tag = ""
Me.txtAlta(5).Text = ""


Me.vleyenda.Text = ""
Me.vCodBanco.Text = ""
Me.vDesBanco.Text = ""
Me.vDesBanco.Tag = ""

Me.VNuevaCustodiaNombre.Tag = ""
Me.vNuevaCustodiaCodigo.Text = ""
Me.VNuevaCustodiaNombre.Text = ""


Me.txtAlta10.Tag = ""
Me.txtAlta11.Tag = ""
Me.txtAlta10.Text = ""
Me.txtAlta11.Text = ""

Me.vCodBanco.Tag = ""
Me.vCodBanco.Text = ""

Me.vDesBanco.Text = ""
Me.vDesBanco.Tag = ""


Me.fvalores.Enable = True

End Sub
Function fcomenCheque() As String
fcomenCheque = ""

If Not vNuevaCustodiaCodigo.Text = "" Then
    fcomenCheque = "Nro.Ch: " + EsNulo(Me.txtAlta(5)) + " - Banco : " + Me.vCodBanco.Text + " " + Me.vDesBanco.Text + " - F.Acred:" + Str(Me.dtpValor.Value)
End If

End Function

Function fcomenCaja() As String
fcomenCaja = ""

If Not Trim(Me.txtAlta6.Text) = "" Then
    fcomenCaja = "Caja: " + Me.txtAlta6 + "-" + Me.txtAlta7
End If

End Function

Function fcomenCta() As String
fcomenCta = ""

If Not Trim(Me.txtAlta11.Text) = "" Then
    fcomenCta = "Concepto : " + txtAlta11.Text
End If

End Function

Private Sub LimpiarCampos(vtipo As Byte)
On Error Resume Next
Dim vvnrointerno As Long

    'Call init
    vpagoPacial = False
    
    Me.vtotalseleccionado = 0
    
    Me.ltotal.Tag = 0
    
    Me.vcliprovee.Tag = 0
    
    vvnrointerno = txtAlta(2).Text ' para que no me lo borre cuando limpia el campo

    'vcodigoCliente = ""
    'vcodigoProveedor = ""


    Dim k As Integer
    Dim vimporte As Double
    
   ' vimporte = txtAlta(12)
    
    For k = 3 To txtAlta.Count - 1
        txtAlta(k).Text = ""
        txtAlta(k).Tag = ""
    Next

    
    bandNrocheque = False
    
    
    txtAlta7 = 0
    txtAlta7.Text = ""
    
    
    txtAlta6.Tag = 0
    txtAlta6.Text = ""
    
    txtAlta10 = 0
    txtAlta10.Text = ""


    txtAlta11 = 0
    txtAlta11.Text = ""

    'txtAlta(12) = vimporte
    'txtAlta(12).Text = ""


    txtAlta(13) = vimporte
    txtAlta(13).Text = ""
    
    
    
   ' RBDebeHaber(0).Value = False
   ' RBDebeHaber(1).Value = True
    
    If vtipo = 1 Then
        
        RBIngresoEgresoCaja(0).Value = False
        RBIngresoEgresoCaja(1).Value = False
    
        dtpFecha.Text = ""
        dtpValor.Text = ""
        
        k = 0
        For k = 0 To 2
            txtAlta(k).Text = ""
            txtAlta(k).Tag = ""
        Next
    
        FormatoGrillaCaja (1)

        txtAlta(0).SetFocus
    Else
        txtAlta(3).SetFocus
    End If
    
    dtpValor.Value = Me.dtpFecha.Value
    dtpFecha.Value = Date
    
    gcustodia.Visible = False
   ' Me.txtAlta(2).Text = vvnrointerno + 1
   
   Me.vcodgocta2.Text = ""
   Me.vcta2.Tag = ""
   Me.vcta2.Text = ""
   
   Me.vconcepto.Text = ""
   Me.vconcepto.Tag = ""
    
    
    Me.txtAlta10.Text = ""
    Me.txtAlta10.Tag = ""
    
    Me.txtAlta11.Text = ""
    Me.txtAlta11.Tag = ""
    
    
    
    vDraft = False
    
    tab2.SelectedItem = 0
    
    txtAlta(3).Text = "EF"
    txtAlta(4).Text = "EFECTIVO"
    
    Me.vCodBanco.Text = ""
    Me.vDesBanco.Text = ""
    Me.vDesBanco.Tag = ""
    
    
    Me.vNuevaCustodiaCodigo.Text = ""
    Me.VNuevaCustodiaNombre.Tag = ""
    Me.VNuevaCustodiaNombre.Text = ""
    
    
    Me.txtAlta(12).Text = ""
    
    vcliprovee.Text = ""
    vcliprovee.Tag = ""
    
    
    vtotalcontrol.Text = ""
    
    Me.vobservacion.Text = ""
    Me.vobservacion2.Text = ""
    
    dtpFecha.Value = Date
    
    Me.vsaldoDisponible.Caption = Format(getSaldoDisponible, "###,###,###,##0.00")
    
If Err Then GrabarLog "LimpiarCampos", Left(Err.Number & " " & Err.Description, 99), Me.Name
End Sub

Private Sub BancoYCaja(vnroasiento As Long)
On Error Resume Next


Exit Sub
    'Caso 1:
        'A: Caja a Caja (2 Lineas)      (Probando : OK)
        'B: Caja a Caja (+2 Lineas)
            
    'Caso 2:
        'A: Banco a Banco (2 Lineas)    (Probando : OK)
        'B: Banco a Banco (+2 Lineas)
    
    'Caso 3:
        'A: Caja a Banco (2 Lineas)     (Probando : OK)
        'B: Caja a Banco (+2 Lineas)
    
    'Caso 4:
        'A: Caja a Otro (2 Lineas)      (Probando : OK)
        'B: Caja a Otro (+2 Lineas)

    'Caso 5:
        'A: Banco a Otro (2 Lineas)     (Probando : OK)
        'B: Banco a Otro (+2 Lineas)

    'Caso 6:
        'Lo estoy pensando....

    Dim vCantidadLineas() As Integer, vEsBanco As String, vEsCaja() As String
    
    Dim rsCajaBanco As New ADODB.Recordset, sqlCajaBanco As String
    
    Dim vidBancos() As String, vIDBancosCuentas() As String
    
    Dim vIDAsientoDetalle As Long, vImporteD() As Double, vImporteH() As Double, vTipoDH() As String, vleyenda() As String
    
    ReDim vCantidadLineas(2)
    ReDim vImporteD(1)
    ReDim vImporteH(1)
    ReDim vTipoDH(1)
    ReDim vleyenda(1)
    ReDim vidBancos(1)
    ReDim vIDBancosCuentas(1)
    
    'sqlCajaBanco = "SELECT CodigoCuenta, Debe, Haber, idBancos, Descripcion, EsCaja FROM AsientosDetalle INNER JOIN Bancos ON CodigoCuenta=CuentaContableAsociada WHERE (Numero = " & vNroAsiento & ") ORDER BY idBancos DESC"
    sqlCajaBanco = "SELECT CodigoCuenta, Debe, Haber, B.idBancos, B.Descripcion, EsCaja,BC.CuentaContableAsociada FROM AsientosDetalle AD LEFT JOIN Bancos B ON CodigoCuenta=CuentaContableAsociada LEFT JOIN BancosCuentas BC ON CodigoCuenta=BC.CuentaContableAsociada WHERE (Numero = " & vnroasiento & ")  And Not ((BC.CuentaContableAsociada Is Null) Or Not (b.CuentaContableAsociada Is Null)) OR NOT (EsCaja IS NULL) ORDER BY idBancos DESC;"
    
    With rsCajaBanco
        Call .Open(sqlCajaBanco, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        'Lineas con Bancos/Cajas Asociadas
        vCantidadLineas(0) = .RecordCount
        
        .Close
        
        'sqlCajaBanco = "SELECT idAsientosDetalle,CodigoCuenta,Debe,Haber,idBancos,Descripcion,EsCaja FROM AsientosDetalle LEFT JOIN Bancos ON CodigoCuenta=CuentaContableAsociada WHERE (Numero = " & vNroAsiento & ") ORDER BY idBancos DESC"
        sqlCajaBanco = "SELECT idAsientosDetalle,CodigoCuenta, Debe, Haber, B.idBancos, B.Descripcion, EsCaja, BC.CuentaContableAsociada FROM AsientosDetalle AD LEFT JOIN Bancos B ON CodigoCuenta=CuentaContableAsociada LEFT JOIN BancosCuentas BC ON CodigoCuenta=BC.CuentaContableAsociada WHERE (Numero = " & vnroasiento & ") ORDER BY idBancos DESC"
        
        Call .Open(sqlCajaBanco, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        'Lineas Totales del Asiento
        vCantidadLineas(1) = .RecordCount
        
        Select Case vCantidadLineas(1)
        
            Case 0, 1
                MsgBox "Error: No puede encontrarse el Asiento al cual hace referencia", vbExclamation, "Mensaje ..."
            
            Case 2
                Select Case vCantidadLineas(0)
                
                    Case 0
                        'No Pasa Nada, el asiento no tenia asociada ninguna linea a Banco/Caja
                        
                    Case 1
                        ReDim vEsCaja(0)
                        
                        vEsCaja(0) = ""
                        vEsCaja(0) = EsNulo(.Fields("EsCaja").Value)
                        
                        'Me quedo con el id de Detalle para hacer un filtro (= Nro de asiento y <> a idAsientoDetalle)
                        vIDAsientoDetalle = EsNulo(.Fields("idAsientosDetalle").Value)
                        
                        If Not IsNull(.Fields("idBancos").Value) = True Then
                            vidBancos(0) = EsNulo(.Fields("idBancos").Value)
                        Else
                            vidBancos(0) = EsNulo(.Fields("idBancos").Value)
                        
                        End If
                        
                        If vEsCaja(0) = "N" Or IsNull(vEsCaja(0)) = True Then
                            vIDBancosCuentas(0) = TraerDato("BancosCuentas", "idBancos = '" & vidBancos(0) & "' AND (CuentaContableAsociada = '11010202')", "idBancosCuentas")
                        Else
                            vIDBancosCuentas(0) = ""
                        End If
                        
                        .Close
                            
                        sqlCajaBanco = "SELECT idAsientosDetalle,CodigoCuenta,Debe,Haber,idBancos,Descripcion,EsCaja FROM AsientosDetalle LEFT JOIN Bancos ON CodigoCuenta=CuentaContableAsociada WHERE (Numero = " & vnroasiento & ") AND NOT (idAsientosDetalle = " & vIDAsientoDetalle & ") ORDER BY idBancos DESC"
                        
                        Call .Open(sqlCajaBanco, ConnDDBB, adOpenStatic, adLockReadOnly)
                            
                        'Caso 4-A o 5-A
                        If Not .EOF = True Then
                            If Val(.Fields("Debe").Value) > 0 Then
                                vImporteD(0) = .Fields("Debe").Value
                                vTipoDH(0) = "D"
                                vleyenda(0) = TraerDato("Cuentas", "CodigoCuenta = '" & .Fields("CodigoCuenta").Value & "'", "Cuenta")
                            Else
                                vImporteH(0) = .Fields("Haber").Value
                                vTipoDH(0) = "H"
                                vleyenda(0) = TraerDato("Cuentas", "CodigoCuenta = '" & .Fields("CodigoCuenta").Value & "'", "Cuenta")
                            End If
                        Else
                            'Error: Aca no se que pasa
                        End If
                        
                        If vIDBancosCuentas(0) = "" Then
                            Call EjecutarScript("INSERT INTO BancosMovimientos (idBancos,Fecha,Debito,Credito,Comentario,NroAsiento) VALUES ('" & vidBancos(0) & "', '" & strfechaMySQL(dtpFecha.Value) & "'," & Val(vImporteD(0)) & "," & Val(vImporteH(0)) & ",'" & Trim(vleyenda(0)) & "'," & vnroasiento & ")")
                        Else
                            Call EjecutarScript("INSERT INTO BancosMovimientos (idBancos,idBancosCuentas,Fecha,Debito,Credito,Comentario,NroAsiento) VALUES ('" & vidBancos(0) & "'," & Val(vIDBancosCuentas(0)) & ",'" & strfechaMySQL(dtpFecha.Value) & "'," & Val(vImporteD(0)) & "," & Val(vImporteH(0)) & ",'" & Trim(vleyenda(0)) & "'," & vnroasiento & ")")
                        End If
                        
                    Case 2
                        'Caso 1-A o Caso 2-A o Caso 3-A

                        .Close
                        
                        sqlCajaBanco = "SELECT idAsientosDetalle,CodigoCuenta,Debe,Haber,idBancos,Descripcion,EsCaja FROM AsientosDetalle LEFT JOIN Bancos ON CodigoCuenta=CuentaContableAsociada WHERE (Numero = " & vnroasiento & ") ORDER BY EsCaja DESC,idAsientosDetalle ASC"
                        
                        Call .Open(sqlCajaBanco, ConnDDBB, adOpenStatic, adLockReadOnly)
                        
                        Dim i As Integer, j As Integer
                        
                        ReDim vEsCaja(1)
                        ReDim vidBancos(1)
                        
                        .Fields.Refresh
                        
                        .MoveFirst
                        
                        
                        For i = 0 To 1
                            
                            vEsCaja(i) = ""
                            vidBancos(i) = ""

                            vImporteD(i) = 0
                            vImporteH(i) = 0
                            vTipoDH(i) = ""
                            vleyenda(i) = ""
                            
                            vEsCaja(i) = EsNulo(.Fields("EsCaja").Value)
                            vidBancos(i) = EsNulo(.Fields("idBancos").Value)
                            
                            vImporteD(i) = Val(.Fields("Debe").Value)
                            vImporteH(i) = Val(.Fields("Haber").Value)
                            
                            vIDBancosCuentas(i) = EsNulo(TraerDato("BancosCuentas", "CuentaContableAsociada = '" & .Fields("CodigoCuenta").Value & "'", "idBancosCuentas"))
                            
                            vleyenda(i) = EsNulo(TraerDato("Cuentas", "CodigoCuenta = '" & .Fields("CodigoCuenta").Value & "'", "Cuenta"))
                            
                            If .Fields("Debe").Value > 0 Then
                                vTipoDH(i) = "D"
                            Else
                                vTipoDH(i) = "H"
                            End If
                            
                            .MoveNext
                        Next
                            
                        If vEsCaja(0) = "S" Then
                            If vEsCaja(1) = "S" Then
                                '1-A: Caja a Caja
                                For i = 0 To 1
                                    If i = 0 Then
                                        Call EjecutarScript("INSERT INTO BancosMovimientos (idBancos,Fecha,Debito,Credito,Comentario,NroAsiento) VALUES ('" & vidBancos(1) & "','" & strfechaMySQL(dtpFecha.Value) & "', " & Val(vImporteD(1)) & ", " & Val(vImporteH(1)) & ",'" & Trim(vleyenda(0)) & "'," & vnroasiento & ")")
                                    Else
                                        Call EjecutarScript("INSERT INTO BancosMovimientos (idBancos,Fecha,Debito,Credito,Comentario,NroAsiento) VALUES ('" & vidBancos(0) & "','" & strfechaMySQL(dtpFecha.Value) & "', " & Val(vImporteD(0)) & ", " & Val(vImporteH(0)) & ",'" & Trim(vleyenda(1)) & "'," & vnroasiento & ")")
                                    End If
                                Next
                            
                            Else
                                
                                '3-A: (Caja a Banco)
                                For i = 0 To 1
                                    If i = 0 Then
                                        Call EjecutarScript("INSERT INTO BancosMovimientos (idBancos,idBancosCuentas,Fecha,Debito,Credito,Comentario,NroAsiento) VALUES ('" & vidBancos(1) & "'," & Val(vIDBancosCuentas(1)) & ",'" & strfechaMySQL(dtpFecha.Value) & "'," & Val(vImporteD(0)) & "," & Val(vImporteH(0)) & ",'" & Trim(vleyenda(0)) & "'," & vnroasiento & ")")
                                    Else
                                        Call EjecutarScript("INSERT INTO BancosMovimientos (idBancos,Fecha,Debito,Credito,Comentario,NroAsiento) VALUES ('" & vidBancos(0) & "','" & strfechaMySQL(dtpFecha.Value) & "'," & Val(vImporteD(1)) & "," & Val(vImporteH(1)) & ",'" & Trim(vleyenda(1)) & "'," & vnroasiento & ")")
                                    End If
                                Next
                            
                            End If
                            
                        Else
                                
                            If vEsCaja(1) = "S" Then
                                '3-A: Banco a Caja
                            Else
                                '2-A: Banco a Banco
                                For i = 0 To 1
                                    If i = 0 Then
                                        Call EjecutarScript("INSERT INTO BancosMovimientos (idBancos,idBancosCuentas,Fecha,Debito,Credito,Comentario,NroAsiento) VALUES ('" & vidBancos(1) & "'," & Val(vIDBancosCuentas(1)) & ",'" & strfechaMySQL(dtpFecha.Value) & "'," & Val(vImporteD(0)) & "," & Val(vImporteH(0)) & ",'" & Trim(vleyenda(0)) & "'," & vnroasiento & ")")
                                    Else
                                        Call EjecutarScript("INSERT INTO BancosMovimientos (idBancos,idBancosCuentas,Fecha,Debito,Credito,Comentario,NroAsiento) VALUES ('" & vidBancos(0) & "'," & Val(vIDBancosCuentas(0)) & ",'" & strfechaMySQL(dtpFecha.Value) & "'," & Val(vImporteD(1)) & "," & Val(vImporteH(1)) & ",'" & Trim(vleyenda(1)) & "'," & vnroasiento & ")")
                                    End If
                                Next
                                
                            End If
                        
                        End If
                        
                    Case 3
                        'ReDim vCantidadLineas(2)
                        'ReDim vimporte(2)
                        ReDim vTipoDH(2)
                        ReDim vleyenda(2)
                        ReDim vidBancos(2)
                    
                        
                        .Fields.Refresh
                        .MoveFirst
                        
                        For i = 0 To 2
                            vEsCaja(i) = ""
                            vidBancos(i) = ""

                            vTipoDH(i) = ""
                            vleyenda(i) = ""
                            
                            vEsCaja(i) = EsNulo(.Fields("EsCaja").Value)
                            vidBancos(i) = EsNulo(.Fields("idBancos").Value)
                            
                            vImporteD(i) = Val(.Fields("Debe").Value)
                            vImporteH(i) = Val(.Fields("Haber").Value)

                            vleyenda(i) = EsNulo(TraerDato("Cuentas", "CodigoCuenta = '" & .Fields("CodigoCuenta").Value & "'", "Cuenta"))
                            
                            If .Fields("Debe").Value > 0 Then
                                vTipoDH(i) = "D"
                            Else
                                vTipoDH(i) = "H"
                            End If
                            
                            .MoveNext
                        Next
                    
                    Case 4
                    
                    Case 5
                    
                    Case 6
                    
                    Case Else
                        
                
                End Select
            
            Case Else
            
        
        End Select
        
        
    
    End With

    If rsCajaBanco.State = 1 Then
        rsCajaBanco.Close
        Set rsCajaBanco = Nothing
    End If

If Err Then GrabarLog "BancoYCaja", Err.Number & " " & Err.Description, Me.Caption
End Sub
Public Sub txtAlta_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
    
        txtAlta(Index).Text = UCase(txtAlta(Index).Text)
        
        Select Case Index
        
            Case 0
                
                txtAlta(Index + 1).Text = TraerDato("TipoMovimientos", "Codigo = '" & Trim(txtAlta(Index).Text) & "'", "TipoMovimiento")
                
                If Not Trim(txtAlta(Index + 1).Text) = "" Then
                    If TraerDato("TipoMovimientos", "Codigo = '" & Trim(txtAlta(Index).Text) & "'", "IngresoEgreso") = "I" Then
                        RBIngresoEgresoCaja(0).Value = True
                    Else
                        RBIngresoEgresoCaja(1).Value = True
                    End If
                    
                    dtpFecha.SetFocus
                Else
                    txtAlta(Index).Text = ""
                    txtAlta(Index + 1).Text = ""
                End If
    
            Case 2
                txtAlta(3).SetFocus
            
            Case 3
                txtAlta(Index + 1).Text = TraerDato("TipoValor", "idTipoValor = '" & Trim(txtAlta(Index).Text) & "'", "TipoValor")
                
                If Not Trim(txtAlta(Index + 1).Text) = "" Then
                    If Not Trim(txtAlta(Index + 1).Text) = "CH" Then
                        txtAlta6.Text = ""
                        'dtpValor.Text = ""
                        txtAlta6.SetFocus
                    Else
                        txtAlta(5).SetFocus
                    End If
                Else
                    txtAlta(Index).Text = ""
                    txtAlta(Index + 1).Text = ""
                End If
            
            Case 6
                 
                If Not (txtAlta(Index).Text) = "" Then
                    txtAlta(Index + 1).Text = TraerDato("Bancos", "idBancos = '" & Trim(txtAlta(Index).Text) & "'", "Descripcion")
                
                    txtAlta(8).Text = ""
                    txtAlta(9).Text = ""
                    txtAlta10.Text = ""
                    txtAlta11.Text = ""
                
                    If Not (txtAlta(Index + 1).Text) = "" Then
                        If TraerDato("Bancos", "idBancos = '" & Trim(txtAlta(Index).Text) & "'", "EsCaja") = "S" Then
                            pbCarga(3).Enabled = False
                            txtAlta(8).Enabled = False
                            txtAlta(9).Enabled = False
                            txtAlta10.Text = TraerDato("Bancos", "idBancos = '" & Trim(txtAlta(Index).Text) & "'", "CuentaContableAsociada")
                            txtAlta(12).SetFocus
                        Else
                            pbCarga(3).Enabled = True
                            txtAlta(8).Enabled = True
                            txtAlta(9).Enabled = True
                            txtAlta(8).SetFocus
                        End If
                    Else
                        txtAlta(5).SetFocus
                    End If
                Else
                    txtAlta(8).SetFocus
                End If
            Case 8
                txtAlta10.SetFocus
            Case 10
                
                If Not txtAlta(Index).Text = "" Then
                    txtAlta(Index + 1).Text = TraerDato("Cuentas", "CodigoCuenta = '" & Trim(txtAlta(Index).Text) & "'", "Cuenta")
                    If txtAlta(Index + 1).Text = "" Then
                        txtAlta(Index).Text = ""
                        txtAlta(Index).SetFocus
                    Else
                        txtAlta(Index + 2).SetFocus
                    End If
                Else
                    txtAlta(Index + 1).Text = ""
                    txtAlta(12).SetFocus
                End If
                

            Case 12
                If tab2.SelectedItem = 1 Then
                    pbCarga(4).SetFocus
                Else
                    txtAlta(Index + 1).SetFocus
                End If
            Case 13
                Me.cmdAgregar.SetFocus
        End Select
    
    End If

If Err Then GrabarLog "txtAlta_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtAlta_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
    If KeyCode = vbKeyF3 Then
        If Index = 8 Then
            pbCarga_Click (3)
        End If
    End If
    
If Err Then GrabarLog "txtAlta_KeyUp", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtAlta_LostFocus(Index As Integer)
On Error Resume Next
    
    
     'SendKeys "{tab}"
     
     txtAlta(Index).Text = UCase(txtAlta(Index))
     
    
    Select Case Index
    
        Case 0
        Case 1
        Case 2
            
            vEstaModificando = False
            
            If Not Val(TraerDato("BancosMovimientos", "NroInterno =  " & Val(txtAlta(Index).Text) & "", "idBancosMovimientos")) = 0 Then
                If MsgBox("El Nro Interno que acaba de ingresar se encuentra cargado previamente. Desea Modificarlo?", vbYesNo + vbInformation, "Mensaje ...") = vbYes Then
                
                    Modificar (txtAlta(Index).Text)
                
                    vEstaModificando = True
                    
                
                Else
                
                End If
                
                
            Else
            
            End If
        
            txtAlta(2).Alignment = xtpEditAlignRight
            
            Me.chkNroInternoFijo.Value = xtpUnchecked
            Me.txtAlta(2).Enabled = False
        Case 3
        Case 4
        Case 5
            ValidadNroChe (txtAlta(5))
        Case 6
        Case 7
        Case 8
    
    End Select
    

If Err Then GrabarLog "txtAlta_LostFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Modificar(vnrointerno As Long)
On Error Resume Next

    Dim rsMovimientos As New ADODB.Recordset, sqlMovimientos As String
    
    sqlMovimientos = "SELECT * FROM BancosMovimientos WHERE (NroInterno = " & Val(vnrointerno) & ")"
    
    With rsMovimientos
        .CursorLocation = adUseClient
        
        Call .Open(sqlMovimientos, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then
            FormatoGrillaCaja (.RecordCount)
            .MoveFirst
        Else
            FormatoGrillaCaja (1)
        End If
        
        Do Until .EOF = True
            KlexMovimientoCaja.TextMatrix(.AbsolutePosition, 1) = EsNulo(.Fields("idBancosMovimientos").Value)
            KlexMovimientoCaja.TextMatrix(.AbsolutePosition, 2) = EsNulo(.Fields("idTipoValor").Value)
            KlexMovimientoCaja.TextMatrix(.AbsolutePosition, 3) = EsNulo(.Fields("NroCheque").Value)
            KlexMovimientoCaja.TextMatrix(.AbsolutePosition, 4) = EsNulo(.Fields("FechaValor").Value)
           ' KlexMovimientoCaja.TextMatrix(.AbsolutePosition, 5) = "[" & EsNulo(.Fields("idBancos").Value) & "]"
            KlexMovimientoCaja.TextMatrix(.AbsolutePosition, 5) = EsNulo(.Fields("idBancos").Value)
            KlexMovimientoCaja.TextMatrix(.AbsolutePosition, 6) = EsNulo(.Fields("idBancosCuentas").Value)
            
            If TraerDato("Bancos", "idBancos = '" & EsNulo(.Fields("idBancos").Value) & "'", "EsCaja") = "S" Then
                KlexMovimientoCaja.TextMatrix(.AbsolutePosition, 7) = TraerDato("Bancos", "idBancos = " & EsNulo(.Fields("idBancos").Value) & "", "CuentaContableAsociada")
            Else
                KlexMovimientoCaja.TextMatrix(.AbsolutePosition, 7) = TraerDato("BancosCuentas", "idBancosCuentas = " & EsNulo(.Fields("idBancosCuentas").Value) & "", "CuentaContableAsociada")
            End If
            
            
            If Val(.Fields("Debito").Value) > 0 Then
                KlexMovimientoCaja.TextMatrix(.AbsolutePosition, 8) = "D"
                KlexMovimientoCaja.TextMatrix(.AbsolutePosition, 9) = EsNulo(.Fields("Debito").Value)
            Else
                KlexMovimientoCaja.TextMatrix(.AbsolutePosition, 8) = "H"
                KlexMovimientoCaja.TextMatrix(.AbsolutePosition, 9) = EsNulo(.Fields("Credito").Value)
            End If
            
            KlexMovimientoCaja.TextMatrix(.AbsolutePosition, 10) = EsNulo(.Fields("Comentario").Value)
            
            .MoveNext
        Loop
        
    End With
    
    sqlMovimientos = ""

    If rsMovimientos.State = 1 Then
        rsMovimientos.Close
        Set rsMovimientos = Nothing
    End If

If Err Then GrabarLog "Modificar", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtAlta11_Change()
txtAlta10.Text = txtAlta11.Tag
Me.lblCta.Caption = txtAlta11.Text
End Sub

Private Sub txtAlta6_Click()
Call fbuscarGrilla("(select * from bancos where not EsCaja ='B') as t", "Descripcion", "idBancos", Me.txtAlta7.Name, Me)     ' ema:
        frmConsultas.vbuscando.SetFocus
End Sub

Private Sub txtAlta7_Change()
    Me.txtAlta6.Text = Me.txtAlta7.Tag
    Me.txtAlta6.Tag = Me.txtAlta7.Tag
End Sub

Private Sub vcliprovee_Change()

If Not vcliprovee.Tag = "" Then
        If hayFacturasImpagas(vcliprovee.Tag) And Me.RBDebeHaber(1).Value Then PusBuscarDocumento_Click
        Me.lbsaldo.Caption = Format(getSaldoProveedor22(vcliprovee.Tag), "###,###,##0.00")
        
        ' cargar el dataset de personas
        
       ' vsql = "select localidad from proveedor where idproveedor = " + Str(vcliprovee.Tag)
       ' vlocalidad = traerDatos2(vsql, "c", pathDBMySQL)
        
       ' vsql = "select localidad from proveedor where idproveedor = " + Str(vcliprovee.Tag)
       ' vEmpresa = traerDatos2(vsql, "c", pathDBMySQL)
        
       ' vlocalidad = traerDatos2(vsql, "c", pathDBMySQL)
       ' vDomicilio = traerDatos2(vsql, "c", pathDBMySQL)
        
End If

End Sub

Private Sub vcliprovee_LostFocus()
If vcliprovee.Text = "" Then
    vcliprovee.Tag = ""
End If
End Sub

Private Sub vCodBanco_Change()
Dim vsql, vsn As String

  vsql = "select escaja from bancos where idbancos='" + Me.vCodBanco.Text + "'"
        vsn = traerDatos2(vsql, "escaja", pathDBMySQL)
        
        If vsn = "N" Then
        
            Me.VNuevaCustodiaNombre.Tag = vCodBanco.Text
            Me.VNuevaCustodiaNombre.Text = Me.vDesBanco.Text
            Me.vNuevaCustodiaCodigo.Text = vCodBanco.Text
            
           ' Me.txtAlta(3) = "CH"
           ' Me.txtAlta(4) = "Cheques"
           
        End If
        
        If Me.RBDebeHaber(1) Then
        
                If getNroCheque(Me.vCodBanco.Text) >= 0 Then
                    txtAlta(5).Text = getNroCheque(Me.vCodBanco.Text) + 1
                    bandNrocheque = True
                Else
                    'txtAlta(5).Text = ""
                End If
                
                End If
                
        If Not vCodBanco.Text = "" Then
            txtAlta7.Tag = ""
            txtAlta7.Text = ""
            txtAlta6.Text = ""
        End If
        
        
        
End Sub

Private Sub vconcepto_Change()
On Error Resume Next
    Call setConceptos(Me.vconcepto.Tag)
If Err Then Exit Sub
End Sub

Private Sub setConceptos(vid As Integer)
On Error Resume Next
Dim va As Integer

Me.txtAlta6 = traerDatos2("select * from conceptos2 where idconceptos=" + Str(Me.vconcepto.Tag), "idbancos", pathDBMySQL)
Me.txtAlta7 = traerDatos2("select * from bancos where idbancos=" + Str(Me.txtAlta6), "Descripcion", pathDBMySQL)
 
va = traerDatos2("select * from conceptos2 where idconceptos=" + Str(Me.vconcepto.Tag), "idcuentas", pathDBMySQL)

'Me.txtAlta(10) = traerDatos2("select * from cuentas where idcuentas=" + Str(va), "codigocuenta", pathDBMySQL)
'Me.txtAlta(11) = traerDatos2("select * from cuentas where idcuentas=" + Str(va), "cuenta", pathDBMySQL)

Me.txtAlta10 = traerDatos2("select * from cuentas where idcuentas=" + Str(va), "codigocuenta", pathDBMySQL)
Me.txtAlta11 = traerDatos2("select * from cuentas where idcuentas=" + Str(va), "cuenta", pathDBMySQL)


va = traerDatos2("select * from conceptos2 where idconceptos=" + Str(Me.vconcepto.Tag), "idcuentas2", pathDBMySQL)

Me.vcta2.Tag = traerDatos2("select * from cuentas where idcuentas=" + Str(va), "CodigoCuenta", pathDBMySQL)

Me.vcta2 = traerDatos2("select * from cuentas where idcuentas=" + Str(va), "cuenta", pathDBMySQL)


Me.vrendicion.Tag = traerDatos2("select * from conceptos2 where idconceptos=" + Str(Val(Me.vconcepto.Tag)), "idrendiciones", pathDBMySQL)

Me.vrendicion.Text = traerDatos2("select * from rendiciones where idrendiciones=" + Str(Val(Me.vrendicion.Tag)), "nombre", pathDBMySQL)

If Err Then Exit Sub
End Sub

Private Sub vcta2_Change()
    Me.vcodgocta2.Text = vcta2.Tag
End Sub

Private Sub vDesBanco_Change()
    Me.vCodBanco.Text = Me.vDesBanco.Tag
End Sub

Private Sub vfiltro_Change()
On Error Resume Next
Dim vf As String

If vfiltro = "" Then
    vf = ""
Else
    vf = "codigo = '" + vfiltro.Text + "' or fecha = '" + strfechaMySQL(vfiltro) + "'"
End If

    Call actualizarGrilla(vf)

If Err Then Exit Sub
End Sub

Private Sub vNuevaCustodiaCodigo_Change()
'Call verificarSaldoCaja(Me.vNuevaCustodiaCodigo.Text)
End Sub


Private Sub verificarSaldoCaja(vcodigo As String)
Dim vmensaje, vsql  As String

Dim vsaldo, vimporte  As Double

vsql = "select sum(t.Debito) - sum(t.Credito) as saldo from bancosmovimientos t" + _
" where t.idBancos = '" + vcodigo + "' Group By t.idBancos"

vsaldo = Val(traerDatos2(vsql, "saldo", pathDBMySQL))

vimporte = Val(Me.txtAlta(12).Text)


If vsaldo - vimporte < 0 Then

    MsgBox "Atención. La caja - banco no tiene fondo para esta operación" + Chr(13) + "Saldo disponible: " + Str(vsaldo), vbCritical
    
End If


If Err Then Exit Sub
End Sub

Private Sub VNuevaCustodiaNombre_Change()
Me.vNuevaCustodiaCodigo.Text = Me.VNuevaCustodiaNombre.Tag
End Sub


Private Sub ImprimirRecibo()
On Error Resume Next

    
    Unload Mantenimiento
    Load Mantenimiento
    
    If MsgBox("Imprime el recibo del cobro/pago ? ", vbYesNo, "Recibos ...") = vbYes Then
        llenarDrRecibo
        drRecibo.Show
    End If
    

If Err Then GrabarLog "ImprimirRecibo", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub llenarDrRecibo()
On Error Resume Next

Dim vvsaldo As Double
Dim vnro As String

'vvsaldo = Format(CalSaldoPersona(Me.txtCliente(0).Text, CP.TablaCtaCte), "$###,###,##0.00")
 
'vnro = Str(getNroRecibo)
 
ActualizarRecibo

Unload Mantenimiento
Load Mantenimiento
 
 With drRecibo
         
         
        If Val(Me.vnro_doc_ale.Text) > 0 Then
            .Sections(2).Controls("enrocomprobante").Caption = Me.vnro_doc_ale.Text
        Else
            .Sections(2).Controls("enrocomprobante").Caption = vnrocomprobante
        End If
        
    
        .Sections(2).Controls("etipo").Caption = Trim(Me.txtAlta(4))
        
        
        If Not Me.txtAlta(0).Text = "TR" Then
        
        
            If Me.RBDebeHaber(1) Then
            
                .Sections("TituloEmpresa").Controls("etiqueta1").Caption = "Órden de Pago"
            
            Else
            
                .Sections("TituloEmpresa").Controls("etiqueta1").Caption = "Recibo"
    
            
            End If
            
        Else
        
                .Sections("TituloEmpresa").Controls("etiqueta1").Caption = "Operación Interna"
                
        End If
        
        .Sections(2).Controls("econcepto").Caption = txtAlta(1).Text
        
        
        .Sections(2).Controls("etiqueta9").Caption = Date
        '.Sections(2).Controls("lbllugar").Caption = vDatosEmpresa.Localidad & ", "
        .Sections(2).Controls("lblfecha").Caption = Me.dtpFecha
        
        .Sections(2).Controls("lblCliente").Caption = "Por cuenta de: " + Trim(Me.vcliprovee)
        
        .Sections(5).Controls("lblconcepto").Caption = Trim(Me.vobservacion.Text)
        .Sections(5).Controls("lblconcepto2").Caption = Trim(Me.vobservacion2.Text)
        
        
        If Me.txtAlta(0).Text = "TR" Then
        
            .Sections(5).Controls("etitulototal").Caption = ""
            .Sections(5).Controls("lbltotal").Caption = ""
            .Sections(5).Controls("eletras").Caption = ""
        
        Else
            
            .Sections(5).Controls("eletras").Caption = EnLetras2(Str(vtotal))
            .Sections(5).Controls("lbltotal").Caption = Format(vtotal, "$###,###,##0.00")
        
        
        End If
        
        
      
        .Sections(5).Controls("esaldoTitulo").Caption = ""
        .Sections(5).Controls("esaldo").Caption = ""
   
        
        If Not LeerXml("eRecibo2") = "" Then .Sections(5).Controls("esaldo").Caption = LeerXml("eRecibo2")
        If Not LeerXml("eRecibo1") = "" Then .Sections(0).Controls("e1").Caption = LeerXml("eRecibo1")
        
        
       '.Sections(5).Controls("esaldo").Caption = Format(vvsaldo, "$ ###,###,##0.00")
        
    '    .Hide
        If Not vDraft Then
                
        Else
                .Sections(2).Controls("enrorecibo").Caption = vnro
        End If
        
    End With



If Err Then
    'MsgBox "Error al intentar hacer el recibo" + Str$(Err)
    Exit Sub
End If

End Sub



Function flineaCta(linea As String, vimporte As Double) As String
On Error Resume Next
Dim v, i As Integer
Dim vrelleno As String


v = Len(linea)

vrelleno = ""
'For i = 1 To (200 - v)

'    vrelleno = vrelleno + "."

'Next


    flineaCta = "                               > " + linea

If Err Then
    flineaCta = linea
    Exit Function
End If
End Function

Private Sub ActualizarRecibo()
On Error Resume Next
Dim vsql, vlinea As String
Dim vsql1, vc, vd As String

Dim i As Integer

vsql = "delete from recibo_temp"
Call EjecutarScript(vsql, pathDBMySQL)

vtotal = 0

With KlexMovimientoCaja
    For i = 1 To .Rows - 1
    
    
        vc = Replace(.TextMatrix(i, 5), "*", "")
        
        vsql1 = "select * from bancos where idbancos='" + vc + "'"
        
        vd = traerDatos2(vsql1, "Descripcion", pathDBMySQL)
        
        'vlinea = Trim(.TextMatrix(i, 10)) + " - Caja: " + vc + "." + vd + " - " + Trim(.TextMatrix(i, 13))
        
         vlinea = Trim(.TextMatrix(i, 10))
        
        If Trim(.TextMatrix(i, 5)) = "*" Then
            vlinea = flineaCta(vlinea, Val(.TextMatrix(i, 9)))
        
        Else
              vtotal = vtotal + Val(.TextMatrix(i, 9))
        End If
            
          
        
            vsql = "insert into recibo_temp (descripcion,monto) values ('" + vlinea + "'," + Trim(Val(.TextMatrix(i, 9))) + ") "
            
       
       If Not vlinea = "" And (Val(.TextMatrix(i, 9)) > 0) Then Call EjecutarScript(vsql, pathDBMySQL)
    
    Next
End With

If Me.txtAlta(0).Text = "VL" Then vtotal = vtotal / 2
        

If Err Then
    Exit Sub
    'MsgBox Err.Description
    'Exit Sub
End If

End Sub

Public Sub initIngreso()
    Me.Show
    'Call RBDebeHaber_Click(1)
    Me.f1.BackColor = vbRed
    RBDebeHaber(1).Value = True
    RBDebeHaber(1).BackColor = vbRed
     RBDebeHaber(0).BackColor = vbRed
     PushButton1.BackColor = vbRed
     PushButton1.Caption = "Buscar Cheque"
     
End Sub


Public Sub initEgreso()
    Me.Show
    'Call RBDebeHaber_Click(0)
    Me.f1.BackColor = vbGreen
    RBDebeHaber(0).Value = True
    RBDebeHaber(0).BackColor = vbGreen
    RBDebeHaber(1).BackColor = vbGreen
    PushButton1.BackColor = vbGreen
    PushButton1.Caption = "Ing.Cheque"
End Sub


Public Sub cargarEventuales()
Dim i As Integer

KlexMovimientoCaja.Rows = 2

With Me.KlexMovimientoCaja

For i = 1 To frmTrabEventuales.grilla.Rows - 1


   .TextMatrix(i, 14) = frmTrabEventuales.grilla.TextMatrix(i, 15) ' idproveedor
   
   .TextMatrix(i, 9) = frmTrabEventuales.grilla.TextMatrix(i, 14) ' importe
   
   .TextMatrix(i, 13) = Me.txtAlta6.Text ' d
   
   .TextMatrix(i, 5) = txtAlta6.Text
   
   
   .TextMatrix(i, 8) = "D"
   
   .TextMatrix(i, 10) = frmTrabEventuales.grilla.TextMatrix(i, 2) + ", " + frmTrabEventuales.grilla.TextMatrix(i, 3) + " - " + frmTrabEventuales.grilla.TextMatrix(i, 4) + ", hs: " + frmTrabEventuales.grilla.TextMatrix(i, 5) + ", V.hs: " + frmTrabEventuales.grilla.TextMatrix(i, 7) + ", hs.E: " + frmTrabEventuales.grilla.TextMatrix(i, 6) + ", V.hs.E: " + frmTrabEventuales.grilla.TextMatrix(i, 8)
   
   .TextMatrix(i, 2) = EsNulo(txtAlta(3).Text)             '(Tipo Valor)
    
    .TextMatrix(i, 3) = EsNulo(txtAlta(5).Text)             '(Nro Valor)
    
    .TextMatrix(i, 4) = EsNulo(dtpFecha.Text)               '(F. Valor)
   
    .TextMatrix(i, 7) = EsNulo(txtAlta10.Text)

    .TextMatrix(i, 13) = EsNulo(Me.vcodgocta2.Text)
   
   .Rows = .Rows + 1
  
Next
  
  ' .TextMatrix(i, 14) = frmTrabEventuales.grilla.TextMatrix(i, 15)

  ' .TextMatrix(i, 14) = frmTrabEventuales.grilla.TextMatrix(i, 15)


End With


Unload frmTrabEventuales
End Sub

Private Sub vobservacion_LostFocus()
vobservacion.Text = UCase(vobservacion.Text)
End Sub

Private Sub vobservacion2_LostFocus()
vobservacion2 = UCase(vobservacion2.Text)
End Sub

Private Sub vrendicion_Change()
Me.txtAlta(12).SetFocus
End Sub

Private Sub vsaldoDisponible_DblClick()
Me.vsaldoDisponible.Caption = Format(getSaldoDisponible, "###,###,###,##0.00")
End Sub


Public Sub setvsqlPago(v() As String)
    vsqlpago = v
End Sub

Public Sub setvsqlPagoAuto(v() As String)
    vsqlpagoAuto = v
End Sub

Private Sub MarcarDocumnetosPagosAuto()
On Error Resume Next
Dim i As Integer
Dim v As String

If Trim(vsqlpagoAuto(1)) = "" Then Exit Sub

If Not valMD Then Exit Sub ' valido la posibilidad

For i = 1 To 100
    v = vsqlpagoAuto(i)
    If Not v = "" Then Call EjecutarScript(v, pathDBMySQL)
Next

If Err Then Exit Sub
End Sub


Private Sub MarcarDocumnetosPagos()
On Error Resume Next
Dim i As Integer
Dim v As String

'If Trim(vsqlpago(1)) = "" Then Exit Sub

If Not valMD Then Exit Sub ' valido la posibilidad

For i = 1 To 100
    v = ""
    v = vsqlpago(i)
    If Not v = "" Then Call EjecutarScript(v, pathDBMySQL)
Next

If Err Then Exit Sub
End Sub


Private Sub PagarCtaCteDirecto(importe As Double, vfecha As Date, vcodigo As String, vnombre As String, vcomentario As String, vnrointerno As Long)
' todo: poner el control de la cuenta corriente de proveedores. ComunaWw
Dim vsaldo1, vsaldo2 As Double


vsaldo1 = CalSaldoPersona(Trim(vcodigo), "pcuentasCorrientes")


Dim sqlInsert As String
sqlInsert = "Insert Into pcuentasCorrientes ( nrointerno, Fecha, Codigo, Nombre,debito,Credito, comentario, TipoMovimiento)" & _
            "VALUES (" + Str(vnrointerno) + ",'" & strfechaMySQL(dtpFecha.Value) & "', '" & Trim(vcodigo) & "', '" & vnombre & "',0, " & Str(importe) & ",'" & vcomentario & "','RC')"
            'Cn.Execute sqlInsert
            
Call EjecutarScript(sqlInsert, pathDBMySQL)

vsaldo2 = CalSaldoPersona(Trim(vcodigo), "pcuentasCorrientes")

Dim vmensaje As String
vmensaje = " El pago fue imputado correctamente. " + Chr(13) + _
" > Saldo Anterior : " + Format(vsaldo1, "###,###,##0.00") + Chr(13) + Chr(13) + _
" > Saldo Actual : " + Format(vsaldo2, "###,###,##0.00")


If vsaldo1 = vsaldo2 Then
    MsgBox "Cuidado ! la cuenta corriente del proveedor " + Trim(vcodigo) + Chr(13) + "El movimiento no pudo ser imputado", vbCritical
Else
    MsgBox vmensaje, vbInformation
End If

End Sub

Function valMD() As Boolean
Dim i As Double
Dim vmen As String

i = Val(Me.vtotalseleccionado) - Val(ltotal.Tag)

If i = 0 Or Val(vtotalseleccionado) = 0 Then
    valMD = True
Else

    valMD = False
    
    vmen = "Ud. ha ingresado un importe de pago diferente al total de los documentos seleccionados. " + Chr(13) + _
    "Quiere marcar a los documentos como pagos de todas manera ?"
    
    
    If MsgBox(vmen, vbYesNo) = vbYes Then
        valMD = True
    Else
        valMD = False
    End If

End If

End Function

Public Function validarTransaccion() As Boolean
Dim i As Integer
Dim vtD, vtC, vtotal As Double

validarTransaccion = True

With KlexMovimientoCaja

    For i = 1 To .Rows - 1
        
        If UCase(.TextMatrix(i, 8)) = "D" Then
            vtD = vtD + Val(.TextMatrix(i, 9))
        End If
        
         If UCase(.TextMatrix(i, 8)) = "H" Then
            vtC = vtC + Val(.TextMatrix(i, 9))
        End If
        
    Next
    
  vtotal = vtD - vtC
    
    If Not vtotal = 0 Then
        MsgBox "Hay una diferencia entre entrada y salida de :" + Format(vtotal, "###,###,##0.00")
        validarTransaccion = False
    End If

End With

End Function


