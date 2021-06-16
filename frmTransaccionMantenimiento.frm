VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.ShortcutBar.v13.0.0.Demo.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "Copia de KlexGrid.ocx"
Begin VB.Form frmTransaccionMantenimiento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimeinto de Transacciones:"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   15585
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   8895
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   15555
      _Version        =   851968
      _ExtentX        =   27437
      _ExtentY        =   15690
      _StockProps     =   68
      ItemCount       =   2
      Item(0).Caption =   "Ver"
      Item(0).ControlCount=   21
      Item(0).Control(0)=   "GroupBox3"
      Item(0).Control(1)=   "Pus(0)"
      Item(0).Control(2)=   "txtvnrointerno"
      Item(0).Control(3)=   "GroupBox6"
      Item(0).Control(4)=   "GroupBox5"
      Item(0).Control(5)=   "GroupBox4"
      Item(0).Control(6)=   "GroupBox2"
      Item(0).Control(7)=   "GroupBox1"
      Item(0).Control(8)=   "TreeView3"
      Item(0).Control(9)=   "PusFiltrarMovimentos"
      Item(0).Control(10)=   "Pus(1)"
      Item(0).Control(11)=   "PusEjecutarOperación"
      Item(0).Control(12)=   "botonCerrar"
      Item(0).Control(13)=   "GroupBox8"
      Item(0).Control(14)=   "log"
      Item(0).Control(15)=   "GroupBox9"
      Item(0).Control(16)=   "buscarTransacciones"
      Item(0).Control(17)=   "GroupBox7"
      Item(0).Control(18)=   "GroupBox10"
      Item(0).Control(19)=   "lblIngUn"
      Item(0).Control(20)=   "vdisplay"
      Item(1).Caption =   "Buscar"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "GroupBox11"
      Begin VB.TextBox txtvnrointerno 
         Height          =   285
         Left            =   2790
         TabIndex        =   6
         Top             =   450
         Width           =   1425
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   1725
         Left            =   90
         TabIndex        =   1
         Top             =   780
         Width           =   7785
         _Version        =   851968
         _ExtentX        =   13732
         _ExtentY        =   3043
         _StockProps     =   79
         Caption         =   "Cuentas Corrientes:"
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
         Begin Grid.KlexGrid GrillaCtaCte 
            Height          =   1455
            Left            =   60
            TabIndex        =   2
            Top             =   240
            Visible         =   0   'False
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   2566
            EnterKeyBehaviour=   0
            BackColorAlternate=   14737632
            GridLines       =   0
            GridLinesFixed  =   0
            Appearance      =   0
            BackColor       =   16777215
            BackColorBkg    =   16777215
            BackColorFixed  =   -2147483626
            BorderStyle     =   0
            Cols            =   5
            FixedCols       =   0
            FixedRows       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   255
            GridColorFixed  =   8421504
            HighLight       =   0
            MouseIcon       =   "frmTransaccionMantenimiento.frx":0000
            Rows            =   10
         End
         Begin XtremeSuiteControls.PushButton PushButton10 
            Height          =   285
            Left            =   7320
            TabIndex        =   3
            Top             =   210
            Width           =   315
            _Version        =   851968
            _ExtentX        =   556
            _ExtentY        =   503
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmTransaccionMantenimiento.frx":001C
         End
         Begin XtremeSuiteControls.PushButton PushButton11 
            Height          =   255
            Left            =   7320
            TabIndex        =   4
            Top             =   540
            Width           =   285
            _Version        =   851968
            _ExtentX        =   503
            _ExtentY        =   450
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmTransaccionMantenimiento.frx":05B6
         End
      End
      Begin XtremeSuiteControls.PushButton Pus 
         Height          =   225
         Index           =   0
         Left            =   4650
         TabIndex        =   5
         Top             =   480
         Width           =   375
         _Version        =   851968
         _ExtentX        =   661
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "+"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox6 
         Height          =   2025
         Left            =   7980
         TabIndex        =   7
         Top             =   6600
         Width           =   7035
         _Version        =   851968
         _ExtentX        =   12409
         _ExtentY        =   3572
         _StockProps     =   79
         Caption         =   "Cheques:"
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
         Begin Grid.KlexGrid GrillaCheques 
            Height          =   1785
            Left            =   60
            TabIndex        =   8
            Top             =   210
            Visible         =   0   'False
            Width           =   6435
            _ExtentX        =   11351
            _ExtentY        =   3149
            EnterKeyBehaviour=   0
            BackColorAlternate=   14737632
            GridLinesFixed  =   2
            Appearance      =   0
            BackColor       =   16777215
            BackColorFixed  =   -2147483626
            BorderStyle     =   0
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
            ForeColor       =   16448
            GridColorFixed  =   8421504
            MouseIcon       =   "frmTransaccionMantenimiento.frx":0B50
            Rows            =   10
         End
         Begin XtremeSuiteControls.PushButton PushButton8 
            Height          =   285
            Left            =   6690
            TabIndex        =   9
            Top             =   210
            Width           =   285
            _Version        =   851968
            _ExtentX        =   503
            _ExtentY        =   503
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmTransaccionMantenimiento.frx":0B6C
         End
         Begin XtremeSuiteControls.PushButton PushButton9 
            Height          =   285
            Left            =   6690
            TabIndex        =   10
            Top             =   480
            Width           =   285
            _Version        =   851968
            _ExtentX        =   503
            _ExtentY        =   503
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmTransaccionMantenimiento.frx":1106
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox5 
         Height          =   2115
         Left            =   30
         TabIndex        =   11
         Top             =   6540
         Width           =   7155
         _Version        =   851968
         _ExtentX        =   12621
         _ExtentY        =   3731
         _StockProps     =   79
         Caption         =   "Bano / Cajas:"
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
         Begin Grid.KlexGrid GrillaMovimientosBanco 
            Height          =   1845
            Left            =   90
            TabIndex        =   12
            Top             =   240
            Visible         =   0   'False
            Width           =   6465
            _ExtentX        =   11404
            _ExtentY        =   3254
            EnterKeyBehaviour=   0
            BackColorAlternate=   14737632
            GridLinesFixed  =   2
            Appearance      =   0
            BackColor       =   16777215
            BackColorFixed  =   -2147483626
            BorderStyle     =   0
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
            ForeColor       =   16448
            GridColorFixed  =   8421504
            MouseIcon       =   "frmTransaccionMantenimiento.frx":16A0
            Rows            =   10
         End
         Begin XtremeSuiteControls.PushButton PushButton6 
            Height          =   285
            Left            =   6720
            TabIndex        =   13
            Top             =   270
            Width           =   285
            _Version        =   851968
            _ExtentX        =   503
            _ExtentY        =   503
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmTransaccionMantenimiento.frx":16BC
         End
         Begin XtremeSuiteControls.PushButton PushButton7 
            Height          =   255
            Left            =   6720
            TabIndex        =   14
            Top             =   570
            Width           =   285
            _Version        =   851968
            _ExtentX        =   503
            _ExtentY        =   450
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmTransaccionMantenimiento.frx":1C56
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox4 
         Height          =   1515
         Left            =   8040
         TabIndex        =   15
         Top             =   4890
         Width           =   7185
         _Version        =   851968
         _ExtentX        =   12674
         _ExtentY        =   2672
         _StockProps     =   79
         Caption         =   "Asientos:"
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
         Begin Grid.KlexGrid GrillaAsientos 
            Height          =   1275
            Left            =   60
            TabIndex        =   16
            Top             =   240
            Visible         =   0   'False
            Width           =   6525
            _ExtentX        =   11509
            _ExtentY        =   2249
            EnterKeyBehaviour=   0
            BackColorAlternate=   16777215
            GridLinesFixed  =   2
            Appearance      =   0
            BackColor       =   16777215
            BackColorFixed  =   -2147483626
            BorderStyle     =   0
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
            ForeColor       =   4210688
            GridColorFixed  =   8421504
            MouseIcon       =   "frmTransaccionMantenimiento.frx":21F0
            Rows            =   10
         End
         Begin XtremeSuiteControls.PushButton PushButton4 
            Height          =   285
            Left            =   6870
            TabIndex        =   17
            Top             =   210
            Width           =   285
            _Version        =   851968
            _ExtentX        =   503
            _ExtentY        =   503
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmTransaccionMantenimiento.frx":220C
         End
         Begin XtremeSuiteControls.PushButton PushButton5 
            Height          =   255
            Left            =   6870
            TabIndex        =   18
            Top             =   510
            Width           =   285
            _Version        =   851968
            _ExtentX        =   503
            _ExtentY        =   450
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmTransaccionMantenimiento.frx":27A6
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   1575
         Left            =   90
         TabIndex        =   19
         Top             =   4800
         Width           =   7245
         _Version        =   851968
         _ExtentX        =   12779
         _ExtentY        =   2778
         _StockProps     =   79
         Caption         =   "Libro Iva:"
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
         Begin Grid.KlexGrid GrillaLibroIva 
            Height          =   1275
            Left            =   60
            TabIndex        =   20
            Top             =   240
            Visible         =   0   'False
            Width           =   6525
            _ExtentX        =   11509
            _ExtentY        =   2249
            EnterKeyBehaviour=   0
            BackColorAlternate=   16777215
            GridLinesFixed  =   2
            Appearance      =   0
            BackColor       =   16777215
            BackColorFixed  =   -2147483626
            BorderStyle     =   0
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
            ForeColor       =   4194368
            GridColorFixed  =   8421504
            MouseIcon       =   "frmTransaccionMantenimiento.frx":2D40
            Rows            =   10
         End
         Begin XtremeSuiteControls.PushButton PushButton2 
            Height          =   285
            Left            =   6690
            TabIndex        =   21
            Top             =   240
            Width           =   315
            _Version        =   851968
            _ExtentX        =   556
            _ExtentY        =   503
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmTransaccionMantenimiento.frx":2D5C
         End
         Begin XtremeSuiteControls.PushButton PushButton3 
            Height          =   285
            Left            =   6690
            TabIndex        =   22
            Top             =   540
            Width           =   315
            _Version        =   851968
            _ExtentX        =   556
            _ExtentY        =   503
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmTransaccionMantenimiento.frx":32F6
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1665
         Left            =   8190
         TabIndex        =   23
         Top             =   960
         Width           =   7005
         _Version        =   851968
         _ExtentX        =   12356
         _ExtentY        =   2937
         _StockProps     =   79
         Caption         =   "Documentos:"
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
         Begin Grid.KlexGrid GrillaDocumentos 
            Height          =   1455
            Left            =   60
            TabIndex        =   24
            Top             =   180
            Visible         =   0   'False
            Width           =   6435
            _ExtentX        =   11351
            _ExtentY        =   2566
            EnterKeyBehaviour=   0
            BackColorAlternate=   16777215
            GridLines       =   0
            GridLinesFixed  =   0
            Appearance      =   0
            BackColor       =   16777215
            BackColorBkg    =   16777215
            BackColorFixed  =   -2147483626
            BorderStyle     =   0
            Cols            =   5
            FixedCols       =   0
            FixedRows       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   8388608
            GridColorFixed  =   8421504
            MouseIcon       =   "frmTransaccionMantenimiento.frx":3890
            Rows            =   10
         End
         Begin XtremeSuiteControls.PushButton PusVerDetalle 
            Height          =   285
            Left            =   6660
            TabIndex        =   25
            Top             =   150
            Width           =   285
            _Version        =   851968
            _ExtentX        =   503
            _ExtentY        =   503
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmTransaccionMantenimiento.frx":38AC
         End
         Begin XtremeSuiteControls.PushButton PushButton1 
            Height          =   285
            Left            =   6660
            TabIndex        =   26
            Top             =   450
            Width           =   285
            _Version        =   851968
            _ExtentX        =   503
            _ExtentY        =   503
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmTransaccionMantenimiento.frx":3E46
         End
      End
      Begin MSComctlLib.TreeView TreeView3 
         Height          =   345
         Left            =   990
         TabIndex        =   27
         Top             =   1200
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   609
         _Version        =   393217
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin XtremeSuiteControls.PushButton PusFiltrarMovimentos 
         Height          =   345
         Left            =   5160
         TabIndex        =   28
         Top             =   450
         Width           =   4215
         _Version        =   851968
         _ExtentX        =   7435
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Obtener la transacción correspondiente al nro interno"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmTransaccionMantenimiento.frx":43E0
      End
      Begin XtremeSuiteControls.PushButton Pus 
         Height          =   225
         Index           =   1
         Left            =   4260
         TabIndex        =   29
         Top             =   480
         Width           =   375
         _Version        =   851968
         _ExtentX        =   661
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "-"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PusEjecutarOperación 
         Height          =   345
         Left            =   9510
         TabIndex        =   30
         Top             =   420
         Width           =   1965
         _Version        =   851968
         _ExtentX        =   3466
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Borrar el Movimientos"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmTransaccionMantenimiento.frx":497A
      End
      Begin XtremeSuiteControls.PushButton botonCerrar 
         Height          =   375
         Left            =   14160
         TabIndex        =   31
         Top             =   390
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cerrar"
         Appearance      =   2
         Picture         =   "frmTransaccionMantenimiento.frx":4F14
      End
      Begin XtremeSuiteControls.GroupBox GroupBox8 
         Height          =   1875
         Left            =   60
         TabIndex        =   32
         Top             =   2610
         Width           =   7845
         _Version        =   851968
         _ExtentX        =   13838
         _ExtentY        =   3307
         _StockProps     =   79
         Caption         =   "Detalle del Documento"
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
         Begin Grid.KlexGrid GrillaFdetalle 
            Height          =   1605
            Left            =   60
            TabIndex        =   33
            Top             =   240
            Visible         =   0   'False
            Width           =   7185
            _ExtentX        =   12674
            _ExtentY        =   2831
            EnterKeyBehaviour=   0
            BackColorAlternate=   14737632
            GridLinesFixed  =   2
            Appearance      =   0
            BackColor       =   16777215
            BackColorFixed  =   -2147483626
            BorderStyle     =   0
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
            ForeColor       =   16744576
            GridColorFixed  =   8421504
            MouseIcon       =   "frmTransaccionMantenimiento.frx":5314
            Rows            =   10
         End
         Begin XtremeSuiteControls.PushButton PushButton12 
            Height          =   315
            Left            =   7350
            TabIndex        =   34
            Top             =   240
            Width           =   285
            _Version        =   851968
            _ExtentX        =   503
            _ExtentY        =   556
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmTransaccionMantenimiento.frx":5330
         End
         Begin XtremeSuiteControls.PushButton PushButton13 
            Height          =   285
            Left            =   7350
            TabIndex        =   35
            Top             =   570
            Width           =   285
            _Version        =   851968
            _ExtentX        =   503
            _ExtentY        =   503
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmTransaccionMantenimiento.frx":58CA
         End
      End
      Begin XtremeSuiteControls.ListBox log 
         Height          =   525
         Left            =   420
         TabIndex        =   36
         Top             =   6750
         Visible         =   0   'False
         Width           =   6975
         _Version        =   851968
         _ExtentX        =   12303
         _ExtentY        =   926
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.GroupBox GroupBox9 
         Height          =   2115
         Left            =   8190
         TabIndex        =   37
         Top             =   2700
         Width           =   6975
         _Version        =   851968
         _ExtentX        =   12303
         _ExtentY        =   3731
         _StockProps     =   79
         Caption         =   "Retenciones:"
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
         Begin Grid.KlexGrid GrillaRetenciones 
            Height          =   1665
            Left            =   60
            TabIndex        =   38
            Top             =   210
            Visible         =   0   'False
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   2937
            EnterKeyBehaviour=   0
            BackColorAlternate=   14737632
            GridLinesFixed  =   2
            Appearance      =   0
            BackColor       =   16777215
            BackColorFixed  =   -2147483626
            BorderStyle     =   0
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
            ForeColor       =   49152
            GridColorFixed  =   8421504
            MouseIcon       =   "frmTransaccionMantenimiento.frx":5E64
            Rows            =   10
         End
         Begin XtremeSuiteControls.PushButton PushButton14 
            Height          =   285
            Left            =   6630
            TabIndex        =   39
            Top             =   240
            Width           =   285
            _Version        =   851968
            _ExtentX        =   503
            _ExtentY        =   503
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmTransaccionMantenimiento.frx":5E80
         End
         Begin XtremeSuiteControls.PushButton PushButton15 
            Height          =   285
            Left            =   6630
            TabIndex        =   40
            Top             =   540
            Width           =   285
            _Version        =   851968
            _ExtentX        =   503
            _ExtentY        =   503
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Picture         =   "frmTransaccionMantenimiento.frx":641A
         End
      End
      Begin XtremeSuiteControls.PushButton buscarTransacciones 
         Height          =   345
         Left            =   11490
         TabIndex        =   41
         Top             =   420
         Width           =   1935
         _Version        =   851968
         _ExtentX        =   3413
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Buscar transacciones"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox7 
         Height          =   315
         Left            =   5730
         TabIndex        =   42
         Top             =   1020
         Visible         =   0   'False
         Width           =   2655
         _Version        =   851968
         _ExtentX        =   4683
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Datos Transaccionales:"
         ForeColor       =   14737632
         BackColor       =   12632256
         Appearance      =   6
         Begin VB.Label lblNroInterno 
            Height          =   555
            Left            =   180
            TabIndex        =   43
            Top             =   630
            Width           =   2295
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox10 
         Height          =   135
         Left            =   90
         TabIndex        =   44
         Top             =   780
         Width           =   15195
         _Version        =   851968
         _ExtentX        =   26802
         _ExtentY        =   238
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
      End
      Begin XtremeSuiteControls.GroupBox GroupBox11 
         Height          =   2835
         Left            =   -69610
         TabIndex        =   47
         Top             =   600
         Visible         =   0   'False
         Width           =   8985
         _Version        =   851968
         _ExtentX        =   15849
         _ExtentY        =   5001
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.FlatEdit FlatEdit1 
            Height          =   285
            Left            =   990
            TabIndex        =   48
            Top             =   1110
            Width           =   3615
            _Version        =   851968
            _ExtentX        =   6376
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            Text            =   "FlatEdit1"
         End
         Begin VB.Label lblPersonas 
            Caption         =   "Personas"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   1140
            Width           =   795
         End
      End
      Begin XtremeShortcutBar.ShortcutCaption vdisplay 
         Height          =   765
         Left            =   5310
         TabIndex        =   46
         Top             =   5310
         Visible         =   0   'False
         Width           =   7095
         _Version        =   851968
         _ExtentX        =   12515
         _ExtentY        =   1349
         _StockProps     =   14
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientColorLight=   12582912
         GradientColorDark=   16761024
         ForeColor       =   4210752
      End
      Begin VB.Label lblIngUn 
         Alignment       =   1  'Right Justify
         Caption         =   "Ing. un número interno para  buscar:"
         Height          =   225
         Left            =   120
         TabIndex        =   45
         Top             =   450
         Width           =   2595
      End
   End
End
Attribute VB_Name = "frmTransaccionMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public idCtacte, idAsientos, idDocumentos, idBancosMovimientos, idLibroIva, idCheques As Integer
Public vtablaCtaCte, vid, vtablaCP, vfacturaCP, vlibroivaCP, vidIvaFacturaCP, vidFacturaCP  As String
Public vViene As String
Dim vproblema As String
Public vasientoNumero, vnrointerno  As Long
Dim vcredito As Double
Public vfecha As Date
Public vnrointernocheque2 As Long
Dim vlog As String
Dim vfechaborrado, c1, c2, c3, c4, c5, c6, c7, c8, c9, c10 As String
Dim nopermitorBorrar As Boolean

Private Sub cargaId()
On Error Resume Next
Dim vsql As String
Dim vidctacteHumbral As Integer


If vtablaCtaCte = "cuentascorrientes" Then vidctacteHumbral = 4317 ' panic: para sacar

If vtablaCtaCte = "pcuentascorrientes" Then vidctacteHumbral = 30710  ' panic: para sacar



Select Case vViene

    Case "frmCtaCteC"
   
   
   'If idCtacte > vidctacteHumbral Then  ' los cuentascorrientes.id menores son migrados entonces tengo que conectarte con el nrointerno
   
   If False Then ' hago que el link entre las transacciones sea por el numero interno
   
   vsql = " SELECT " & _
          " " + vtablaCtaCte + "." + vid + ", " & _
          " " + vfacturaCP + "." + vidFacturaCP + ", " & _
          " asientos.idAsientos, asientos.numero, " & _
          " " + vlibroivaCP + "." + vidIvaFacturaCP + ", " & _
          " bancosmovimientos.idBancosMovimientos, asientos.numero, " + vtablaCtaCte + ".NroInterno " & _
          " From " & _
          " " + vtablaCtaCte + " " & _
          " LEFT OUTER JOIN " + vfacturaCP + " ON (" + vtablaCtaCte + ".Remito=" + vfacturaCP + ".Remito) " & _
          " left outer join " + vlibroivaCP + " ON (" + vfacturaCP + ".Remito=" + vlibroivaCP + ".Remito) " & _
          " LEFT OUTER JOIN asientos ON (" + vtablaCtaCte + ".NroAsiento=asientos.Numero) " & _
          " AND (" + vtablaCtaCte + ".Fecha=asientos.Fecha) " & _
          " LEFT OUTER JOIN bancosmovimientos ON (asientos.Numero=bancosmovimientos.NroAsiento) " & _
          " where " + vtablaCtaCte + "." + Trim(vid) + "=" + Trim(Val(EsNulo(idCtacte))) ' + "order by " + vtablaCtaCte + "." + Me.vid + " desc"
          
          
    Else
          
    vsql = " SELECT " & _
          " " + vtablaCtaCte + "." + vid + ", " & _
          " " + vfacturaCP + "." + vidFacturaCP + ", " & _
          " asientos.idAsientos, asientos.numero, " & _
          " " + vlibroivaCP + "." + vidIvaFacturaCP + ", " & _
          " bancosmovimientos.idBancosMovimientos, " + vtablaCtaCte + ".NroInterno " & _
          " From " & _
          " " + vtablaCtaCte + " " & _
          " LEFT OUTER JOIN " + vfacturaCP + " ON (" + vtablaCtaCte + ".Remito=" + vfacturaCP + ".Remito) " & _
          " left outer join " + vlibroivaCP + " ON (" + vfacturaCP + ".Remito=" + vlibroivaCP + ".Remito) " & _
          " LEFT OUTER JOIN asientos ON (" + vtablaCtaCte + ".NroInterno=asientos.NroInterno) " & _
          " AND (" + vtablaCtaCte + ".Fecha=asientos.Fecha) " & _
          " LEFT OUTER JOIN bancosmovimientos ON (asientos.NroInterno=bancosmovimientos.NroInterno) " & _
          " where " + vtablaCtaCte + "." + Trim(vid) + "=" + Trim(Val(EsNulo(idCtacte))) '+ "order by " + vtablaCtaCte + "." + Me.vid + " desc"
          
    End If
          
     Print "ff"
          
          
          
          
    Case "frmAsientos"
    
    'If vasientoNumero > 188 Then ' lo hago para compensar los asientos mal puesto los numeros porque el modulo de bancomovimiento no lo hacia bien
    If True Then
       vsql = " SELECT " & _
          " " + vtablaCtaCte + "." + vid + ", " & _
          " " + vfacturaCP + "." + vidFacturaCP + ", " & _
          " asientos.idAsientos, asientos.numero, " & _
          " " + vlibroivaCP + "." + vidIvaFacturaCP + ", " & _
          " bancosmovimientos.idBancosMovimientos, " + vtablaCtaCte + ".NroInterno " & _
          " From " & _
          " " + vtablaCtaCte + " " & _
          " LEFT OUTER JOIN " + vfacturaCP + " ON (" + vtablaCtaCte + ".Remito=" + vfacturaCP + ".Remito) " & _
          " left outer join " + vlibroivaCP + " ON (" + vfacturaCP + ".Remito=" + vlibroivaCP + ".Remito) " & _
          " right OUTER JOIN asientos ON (" + vtablaCtaCte + ".nrointerno=asientos.nrointerno) " & _
          " AND (" + vtablaCtaCte + ".Fecha=asientos.Fecha) " & _
          " LEFT OUTER JOIN bancosmovimientos ON (asientos.nrointerno=bancosmovimientos.nrointerno) " & _
          " where asientos.numero=" + Trim(EsNulo(vasientoNumero))
        
    Else
        
        vsql = " SELECT " & _
          " " + vtablaCtaCte + "." + vid + ", " & _
          " " + vfacturaCP + "." + vidFacturaCP + ", " & _
          " asientos.idAsientos, asientos.numero, " & _
          " " + vlibroivaCP + "." + vidIvaFacturaCP + ", " & _
          " bancosmovimientos.idBancosMovimientos, " + vtablaCtaCte + ".NroInterno " & _
          " From " & _
          " " + vtablaCtaCte + " " & _
          " LEFT OUTER JOIN " + vfacturaCP + " ON (" + vtablaCtaCte + ".Remito=" + vfacturaCP + ".Remito) " & _
          " left outer join " + vlibroivaCP + " ON (" + vfacturaCP + ".Remito=" + vlibroivaCP + ".Remito) " & _
          " right OUTER JOIN asientos ON (" + vtablaCtaCte + ".NroAsiento=asientos.Numero) " & _
          " AND (" + vtablaCtaCte + ".Fecha=asientos.Fecha) " & _
          " LEFT OUTER JOIN bancosmovimientos ON (asientos.NroInterno=bancosmovimientos.NroInterno) " & _
          " where asientos.numero=" + Trim(EsNulo(vasientoNumero)) + "order by " + vtablaCtaCte + "." + Me.vid + " desc"
        
        
    End If

End Select

'vnrointerno = Val(traerDatos2(vsql, "nrointerno", pathDBMySQL))
idCtacte = Val(traerDatos2(vsql, "" + vid + "", pathDBMySQL))
idAsientos = Val(traerDatos2(vsql, "idAsientos", pathDBMySQL))
If vasientoNumero = 0 Then vasientoNumero = Val(traerDatos2(vsql, "numero", pathDBMySQL))
idDocumentos = Val(traerDatos2(vsql, vidFacturaCP, pathDBMySQL))
idLibroIva = Val(traerDatos2(vsql, "" + vidIvaFacturaCP + "", pathDBMySQL))

' panic! hacer que funcione con este Id. Ahora no lo puedo ver por la consulta, por eso no está  esta tabla.
idBancosMovimientos = Val(traerDatos2(vsql, "idBancosMovimientos", pathDBMySQL))
idCheques = Val(traerDatos2(vsql, "" + vtablaCtaCte + ".Id", pathDBMySQL))

'log.Clear

mostrarIds

' completar: mensaje por no encontrar todos los datos para borrar correctamete

If Err Then
   ' MsgBox "Debe seleccionar un movimiento de cuenta corrientes", vbCritical, "Cuidado..."
    Exit Sub
End If
End Sub
Private Sub mostrarIds()
'-----------------------------------------------------------------------
'Me.log.AddItem ("NroInterno: " + Str(vnrointerno))
'Me.log.AddItem ("isCtaCte: " + Str(idCtacte))
'Me.log.AddItem ("idAsiento: " + Str(idAsientos))
'Me.log.AddItem ("NroAsiento: " + Str(vasientoNumero))
'Me.log.AddItem ("idDocumentos: " + Str(idDocumentos))
'Me.log.AddItem ("idLibroIva: " + Str(idLibroIva))
'Me.log.AddItem ("idBanosMovimientos: " + Str(idBancosMovimientos))
'Me.log.AddItem ("idCheque: " + Str(idCheques))
'------------------------------------------------------------------------
End Sub

Private Sub llenarGrillas()
Dim vsql As String

    vproblema = ""

    llenarGrillaCtaCte ' esto son los que hay acá adentro: lenarGrillaBanco   -  llenarGrillaFactura -   llenarGrillaLibroIva
    
    llenarGrillaBanco
    
    llenarGrillaFactura
    
    llenarGrillaFDetalle
    
    llenarGrillaLibroIva
    
    llenarGrillaAsientos
    
   llenarGrillaCheque
   
   llenarGrillaRetenciones


End Sub

Private Sub llenarGrillaFDetalle()
On Error Resume Next
Dim vremito As Long
Dim vsql As String

vremito = traerDatos2("select remito from factura where nrointerno=" + Trim(vnrointerno), "remito", pathDBMySQL)
vsql = "select * from fdetalle where remito=" + Trim(vremito)
c3 = "fdetalle"

If vremito = 0 Then
    vremito = traerDatos2("select remito from pfactura where nrointerno=" + Trim(vnrointerno), "remito", pathDBMySQL)
    vsql = "select * from pfdetalle where remito=" + Trim(vremito)
    c3 = "pfdetalle"
End If
        
        Call LlenarGrilla("fdetalle", Me.GrillaFdetalle, vsql, "")
If Err Then Exit Sub
End Sub

Private Sub llenarGrillaRetenciones()
Dim vsql As String

        vsql = "select * from retencionesmovimientos where nrointerno=" + Trim(vnrointerno)
        
        c8 = "retencionesmovimientos"
        'Me.idAsientos = Val(traerDatos2(vsql, "id", pathDBMySQL))
         
        Call LlenarGrilla("retencionesmovimientos", Me.GrillaRetenciones, vsql, "")
End Sub

Private Sub llenarGrillaCheque()
Dim vsql As String

        vsql = "select * from cheques where nrointerno=" + Trim(vnrointerno)
        c6 = "cheques"
        
        Me.idAsientos = Val(traerDatos2(vsql, "idAsientos", pathDBMySQL))
         
        Call LlenarGrilla("cheques", Me.GrillaCheques, vsql, "")
End Sub
Private Sub llenarGrillaAsientos()
Dim vsql As String

     vsql = "select fecha as c from asientos where nrointerno=" + Trim(Str(vnrointerno))
    
    
    c5 = "asientos"
    
    If Not vfechaborrado = traerDatos2(vsql, "c", pathDBMySQL) And Not vfechaborrado = "" And Not traerDatos2(vsql, "c", pathDBMySQL) = "" Then
        nopermitorBorrar = True
    End If
    


        vsql = "select * from asientos where asientos.nrointerno=" + Trim(vnrointerno)
        Me.idAsientos = Val(traerDatos2(vsql, "idAsientos", pathDBMySQL))
         
        Call LlenarGrilla("asientos", Me.GrillaAsientos, vsql, "")
End Sub

Private Sub llenarGrillaLibroIva()
Dim vsql As String

   
   



 vsql = "select * from ivafacturaventa where nrointerno = " + Str(vnrointerno)
    Call LlenarGrilla("ivafacturaventa", Me.GrillaLibroIva, vsql, "")
    c4 = "ivafacturaventa"
      
If Not Me.GrillaLibroIva.Tag = "visible" Then
 vsql = "select * from ivafacturacompra where nrointerno = " + Str(vnrointerno)
    Call LlenarGrilla("ivafacturacompra", Me.GrillaLibroIva, vsql, "")
    c4 = "ivafacturacompra"

End If


End Sub



Private Sub llenarGrillaFactura()
Dim vsql As String

vsql = "select * from factura where nrointerno =" + Str(vnrointerno)
Call LlenarGrilla("factura", Me.GrillaDocumentos, vsql, "")
c9 = "factura"
    
If Not Me.GrillaDocumentos.Tag = "visible" Then
    vsql = "select * from pfactura where nrointerno =" + Str(vnrointerno)
    c9 = "pfactura"
    Call LlenarGrilla("pfactura", Me.GrillaDocumentos, vsql, "")
End If


End Sub


'Private Sub llenarGrillaStock()
'Dim vsql As String
'
'Dim vIDFDetalle, vremito As Long
'
'' alguno de los dos tiene q
'vremito = Val(EsNulo(traerDatos2("select * from Factura where nrointerno=" + Str(vnrointerno), "remito", pathDBMySQL)))
'vremito = vremito + Val(EsNulo(traerDatos2("select * from PFactura where nrointerno=" + Str(vnrointerno), "remito", pathDBMySQL)))
'
'
'vIDFDetalle = Val(EsNulo(traerDatos2("select * from factura where remito=" + Str(vremito), "idPFDetalle", pathDBMySQL))) + Val(EsNulo(traerDatos2("select * from pfactura where remito=" + Str(vremito), "idPFDetalle", pathDBMySQL)))
'
'
'vsql = "select * from stock where (idFDetalle =" + Str(vIDFDetalle) + ") or (idPFDetalle =" + Str(vIDFDetalle) + ")"
'
'Call LlenarGrilla("stock", Me.GrillaCtaCte, vsql, "")
'
'End Sub


Private Sub llenarGrillaCtaCte()
Dim vsql As String


vsql = "select fecha as c from cuentascorrientes where nrointerno=" + Trim(Str(vnrointerno))
vfechaborrado = traerDatos2(vsql, "c", pathDBMySQL)


 vsql = "select * from cuentascorrientes where nrointerno=" + Trim(Str(vnrointerno))

Call LlenarGrilla("cuentascorrientes", Me.GrillaCtaCte, vsql, "")
c1 = "cuentascorrientes"


vsql = "select count() as  c  from cuentascorrientes"

If Val(traerDatos2(vsql, "c", pathDBMySQL)) > 1 Then
    nopermitorBorrar = True
End If

If Not Me.GrillaCtaCte.Tag = "visible" Then
     
    vsql = "select fecha as c from pcuentascorrientes where nrointerno=" + Trim(Str(vnrointerno))
    vfechaborrado = traerDatos2(vsql, "c", pathDBMySQL)
     
     
     vsql = "select * from pcuentascorrientes where nrointerno=" + Str(vnrointerno)
     Call LlenarGrilla("pcuentascorrientes", Me.GrillaCtaCte, vsql, "")
     
     c1 = "pcuentascorrientes"
     
     vsql = "select count() as  c  from pcuentascorrientes"

    If Val(traerDatos2(vsql, "c", pathDBMySQL)) > 1 Then
        nopermitorBorrar = True
    End If

     
     
End If


    
End Sub

Private Sub llenarGrillaBanco()
    Dim vsql  As String
    
    vsql = "select fecha as c from bancosmovimientos where nrointerno=" + Trim(Str(vnrointerno))
    
    
    c2 = "bancosmovimientos"
    
    
    If Not vfechaborrado = traerDatos2(vsql, "c", pathDBMySQL) And Not vfechaborrado = "" And Not traerDatos2(vsql, "c", pathDBMySQL) = "" Then
        nopermitorBorrar = True
    End If
    
    
    vsql = "select * from bancosmovimientos where bancosmovimientos.NroInterno=" + Str(vnrointerno)
    Call LlenarGrilla("bancomovimientos", Me.GrillaMovimientosBanco, vsql, "")
End Sub


Private Sub KlexFacturas_BeforeEdit(Cancel As Boolean)

End Sub
Private Sub ArmarArbol()
On Error Resume Next
Dim nodX As Node
Dim nodX2 As Node
Dim nodX3 As Node
Dim nodX4 As Node


Set nodX = TreeView3.Nodes.Add(, "P", , "Mantenimiento")


Set nodX = TreeView3.Nodes.Add(1, tvwChild, "D1", "Borrar")
Set nodX = TreeView3.Nodes.Add("D1", tvwChild, , "Todo")
Set nodX = TreeView3.Nodes.Add("D1", tvwChild, , "Por Módulos")
'Set nodX = TreeView3.Nodes.Add("D1", tvwChild, , "Liq. Final del Parcial")

Set nodX = TreeView3.Nodes.Add(1, tvwChild, "N1", "Ver Detalles")
Set nodX = TreeView3.Nodes.Add("N1", tvwChild, , "Por Módulos")
Set nodX = TreeView3.Nodes.Add("N1", tvwChild, , "Agrupados")
'Set nodX = TreeView3.Nodes.Add("N1", tvwChild, , "Liq. Final del Parcial")



'Set nodX2 = TreeView3.Nodes.Add(, "P", , "Errores")

'Set nodX2 = TreeView3.Nodes.Add(1, tvwChild, "DD1", "Inconsistencias")
'Set nodX2 = TreeView3.Nodes.Add("DD1", tvwChild, , "")
'Set nodX2 = TreeView3.Nodes.Add("DD1", tvwChild, , "Liq. Parcial")
'Set nodX2 = TreeView3.Nodes.Add("DD1", tvwChild, , "Liq. Final del Parcial")

'Set nodX2 = TreeView3.Nodes.Add(1, tvwChild, "NN1", "Flete en deducción")
'Set nodX2 = TreeView3.Nodes.Add("NN1", tvwChild, , "Liq. Final")
'Set nodX2 = TreeView3.Nodes.Add("NN1", tvwChild, , "Liq. Parcial")
'Set nodX2 = TreeView3.Nodes.Add("NN1", tvwChild, , "Liq. Final del Parcial")

If Err Then Exit Sub
End Sub


Private Sub Form_Load()
init
End Sub


Private Sub init()

nopermitorBorrar = False

vnrointernocheque2 = 0
Me.txtvnrointerno = vnrointerno

If vnrointerno = 0 Then
    MsgBox "Usted no ha seleccionado correctamente una transacción", vbInformation, "Cuidado..."
    Exit Sub
Else

llenarGrillas
 

End If
                
 Me.Top = 0
          
ArmarArbol ' armo el menu de la izquierda
End Sub

Private Sub botonCerrar_Click()
Unload Me
End Sub

Function validar() As Boolean
validar = True

If Me.txtvnrointerno = 1 Then
    MsgBox "Este movimiento no puede ser borrado por el usuario. Contacte al servicio tecnico"
    validar = False
End If

If nopermitorBorrar Then
    validar = False
    MsgBox "Este movimiento no puede ser borrardo. Consulte al soporte técnico."
End If

If Not ValidarCajaCerrada Then validar = False

If UCase(LeerXml("Puesto")) = "NOBORRA" Then
    validar = False
    MsgBox "No tiene permiso para efectuar esta operación."
End If

End Function


Function ValidarCajaCerrada() As Boolean
On Error Resume Next
Dim vsql As String
Dim vfecha As Date
Dim vidS As String

ValidarCajaCerrada = True

vsql = "select * from bancosmovimientos where nrointerno = " + Str(vnrointerno)
vfecha = CDate(traerDatos2(vsql, "fecha", pathDBMySQL))

vsql = "select * from t_cajacierre where fecha = '" + strfechaMySQL(vfecha) + "'"
vidS = traerDatos2(vsql, "idcajacierre", pathDBMySQL)

If Not Trim(vidS) = "" Then
    MsgBox "No se puede realizar la operación.  La caja está cerrada", vbInformation, "Caja Cerrada."
    ValidarCajaCerrada = False
End If

If Err Then Exit Function
End Function


Private Sub borrarTodosLosModulos()
'log.Clear


vlog = ""



    Call BorrarBase("pcuentascorrientes WHERE nrointerno=" + Str(vnrointerno), pathDBMySQL)
    Call BorrarBase("cuentascorrientes WHERE nrointerno=" + Str(vnrointerno), pathDBMySQL)

    Call BorrarFDetalle(vnrointerno, "idPFDetalle")
    Call BorrarFDetalle(vnrointerno, "idFDetalle")
    
    Call BorrarBase("factura WHERE nrointerno=" + Str(vnrointerno), pathDBMySQL)
    Call BorrarBase("pfactura WHERE nrointerno=" + Str(vnrointerno), pathDBMySQL)
    

    Call BorrarBase("ivafacturaventa WHERE nrointerno=" + Str(vnrointerno), pathDBMySQL)
    Call BorrarBase("ivafacturacompra WHERE nrointerno=" + Str(vnrointerno), pathDBMySQL)


    Call BorrarBase("bancosmovimientos WHERE nrointerno=" + Str(vnrointerno), pathDBMySQL)
    
    
    vasientoNumero = traerDatos2("select * from asientos where nrointerno=" + Str(vnrointerno), "numero", pathDBMySQL)
    
    Call BorrarBase("bancosmovimientos WHERE nrointerno=" + Str(vnrointerno), pathDBMySQL)
    
    
    Call BorrarBase("asientos" + " WHERE numero=" + Trim(Str(Val(Me.vasientoNumero))), pathDBMySQL)
    Call BorrarBase("asientosdetalle" + " WHERE Numero=" + Trim(Str(Val(Me.vasientoNumero))), pathDBMySQL)
    
    Call BorrarBase("retencionesmovimientos" + " WHERE nrointerno=" + Str(vnrointerno), pathDBMySQL)
    
    
    Call BorrarBase("t_logcaja" + " WHERE nrointerno=" + Str(vnrointerno), pathDBMySQL)
    
    MsgBox "Datos borrados"
    ' verifica si hay cheques que deben ser borrados o hay que cambiar la custodia
    verificarChequesBorrar (vnrointerno)
    'Call BorrarBase("cheques" + " WHERE nrointerno=" + Str(vnrointerno), pathDBMySQL)
    
End Sub
Private Sub verificarChequesBorrar(vnrointerno As Long)
Dim vcount As Integer
Dim vsql As String

vsql = "select count(idCheques) as c from cheques where nrointerno=" + Str(vnrointerno)
vcount = traerDatos2(vsql, "c", pathDBMySQL)

If Not vcount > 0 Then Exit Sub

vsql = "En esta transacción hay cheques involucrados. " + Chr(13) + "Debe evaluar si los borra o cambia la custodia y estado." + Chr(13) + "Desea borrar el/los cheques ?"
If MsgBox(vsql, vbYesNo, "Cheques") = vbYes Then
    Call BorrarBase("cheques" + " WHERE nrointerno=" + Str(vnrointerno), pathDBMySQL)
Else
    vsql = "Debe cambiarle la custodia a los siguientes los cheques que se mostrarán a continuación..."
    MsgBox vsql, vbInformation, "Cheques"
    
    
    Dim vniCheques As String
    Dim i As Integer
    
    vsql = "select * from cheques where nrointerno=" + Str(vnrointerno)
    
    
    Dim bcheque As New ADODB.Recordset
    Call bcheque.Open(vsql, ConnDDBB, adOpenStatic, adLockReadOnly)

    vniCheques = ""
    
    For i = 1 To vcount
    
        vniCheques = "1=2 or "
        vniCheques = vniCheques + "or nrointerno=" + Str(bcheque("nrointerno"))
        
        bcheque.NextRecordset
    
    Next
End If
  
End Sub


Private Sub verCreditoenVenta()
    Dim vidCredito As Long
    Dim vvsql As String
    vidCredito = 0
    vvsql = "select * from " + vtablaCtaCte + " where credito > 0 and NroAsiento=" + Str(vasientoNumero)
    
    vidCredito = Val(traerDatos2(vvsql, vid, pathDBMySQL))
    
            If Not vidCredito > 0 Then ' si no lo encuentro por el nrode asiento lo busco por nrointerno
                   vidCredito = Val(traerDatos2("select * from " + vtablaCtaCte + " where credito > 0 and NroInterno=" + Str(vnrointerno), vid, pathDBMySQL))
            End If
       
    If vidCredito > 0 Then
    
        ' verifico si no tiene asociado movimientos de bancos para este credito
        vvsql = "select * from bancosmovimientos where NroAsiento=" + Str(vasientoNumero)
        Me.idBancosMovimientos = Val(traerDatos2(vvsql, "idbancosmovimientos", pathDBMySQL))
        
        If Me.idBancosMovimientos = 0 Then ' en el caso que no lo encuentre con el nroasiento lo busco con el nrointerno
                vvsql = "select * from bancosmovimientos where NroInterno=" + Str(vnrointerno)
                Me.idBancosMovimientos = Val(traerDatos2(vvsql, "idbancosmovimientos", pathDBMySQL))
        End If
        
    
       ' log.AddItem ("(1) Borrando ctacte Credito... ")
        Call BorrarBase(vtablaCtaCte + " WHERE (" + vid + " = " & Str(vidCredito) & ")", pathDBMySQL)
    
    End If
    
    
    
    'log.AddItem ("(1) Borrando ctacte Debito ... ")
    Call BorrarBase(vtablaCtaCte + " WHERE (" + vid + " = " & Str(Me.idCtacte) & ")", pathDBMySQL)
    Me.GrillaCtaCte.Visible = False

End Sub
Private Sub verDebitoEnCompra()
    Dim vidDebito As Long
    Dim vvsql As String
    vidDebito = 0
    vvsql = "select * from " + vtablaCtaCte + " where debito > 0 and NroAsiento=" + Str(vasientoNumero)
    
    vidDebito = Val(traerDatos2(vvsql, vid, pathDBMySQL))
    
            If Not vidDebito > 0 Then ' si no lo encuentro por el nrode asiento lo busco por nrointerno
                   vidDebito = Val(traerDatos2("select * from " + vtablaCtaCte + " where debito > 0 and NroInterno=" + Str(vnrointerno), vid, pathDBMySQL))
            End If
       
    If vidDebito > 0 Then
    
        ' verifico si no tiene asociado movimientos de bancos para este debito
        vvsql = "select * from bancosmovimientos where NroAsiento=" + Str(vasientoNumero)
        Me.idBancosMovimientos = Val(traerDatos2(vvsql, "idbancosmovimientos", pathDBMySQL))
        
        If Me.idBancosMovimientos = 0 Then ' en el caso que no lo encuentre con el nroasiento lo busco con el nrointerno
                vvsql = "select * from bancosmovimientos where NroInterno=" + Str(vnrointerno)
                Me.idBancosMovimientos = Val(traerDatos2(vvsql, "idbancosmovimientos", pathDBMySQL))
        End If
        
    
        log.AddItem ("(1) Borrando ctacte Debito... ")
        Call BorrarBase(vtablaCtaCte + " WHERE (" + vid + " = " & Str(vidDebito) & ")", pathDBMySQL)
    
    End If
    
    
    
    log.AddItem ("(1) Borrando ctacte Credito ... ")
    Call BorrarBase(vtablaCtaCte + " WHERE (" + vid + " = " & Str(Me.idCtacte) & ")", pathDBMySQL)
    Me.GrillaCtaCte.Visible = False
End Sub


Private Sub GrillaMovimientosBanco_Click()
On Error Resume Next
Dim r, c As Integer
Dim vid As Long


r = GrillaMovimientosBanco.Row
c = GrillaMovimientosBanco.Cols - 2
vid = GrillaMovimientosBanco.TextMatrix(r, c)
Call llenarGrillaChequeMovBanco(vid)

If Err Then Exit Sub
End Sub

Private Sub Pus_Click(Index As Integer)
txtvnrointerno.Text = Val(txtvnrointerno.Text) + 1 - (Index * 2)
Call PusFiltrarMovimentos_Click
End Sub

Private Sub PusEjecutarOperación_Click()

If Not validar Then Exit Sub


If Not MsgBox("¿Esta seguro que desea borrar la transaccion con nro interno: " & vnrointerno & "?", vbInformation + vbYesNo) = vbYes Then Exit Sub
 
Dim vmotivo As String
vmotivo = InputBox("Ingrese la clave para poder realizar esta operación.", "Borrado ...")

If Not vmotivo = "dalas" Then Exit Sub


borrarTodosLosModulos

GrabarLog "Borrar Transacción (***)", vmotivo + "Asiento:" + Str(Val(vasientoNumero)) + " Interno: " + Str(vnrointerno) + " - " + vlog, Me.Caption


llenarGrillas ' lleno las grilla para que no muestre  los datos guardados

If vViene = "frmAsientos" Then
    Call frmAsientos.PbAcciones_Click(4)
Exit Sub
End If

If vViene = "frmCtaCteC" Then
    'Call frmCtaCteC.cmdFiltroMovimientos_Click
    Unload frmCtaCteC
End If
'Unload Me
End Sub

Private Sub PusFiltrarMovimentos_Click()
    vnrointerno = Val(Me.txtvnrointerno)
    Call init
End Sub

Private Sub PushButton1_Click()
Dim vsql As String
Dim v As String
Dim vcampo, vdato As String
vcampo = GrillaDocumentos.TextMatrix(0, 1)
vdato = GrillaDocumentos.TextMatrix(GrillaDocumentos.Row, 1)

vsql = "delete from " + c9 + " where " + vcampo + " = " + vdato

Call EjecutarScript(vsql, pathDBMySQL)

Call PusFiltrarMovimentos_Click
End Sub

Private Sub PushButton11_Click()
Dim vsql As String
Dim v As String
Dim vcampo, vdato As String
vcampo = GrillaCtaCte.TextMatrix(0, 1)
vdato = GrillaCtaCte.TextMatrix(GrillaCtaCte.Row, 1)

vsql = "delete from " + c1 + " where " + vcampo + " = " + vdato

Call EjecutarScript(vsql, pathDBMySQL)

Call PusFiltrarMovimentos_Click

End Sub

Private Sub PushButton13_Click()
Dim vsql As String
Dim v As String
Dim vcampo, vdato As String
vcampo = GrillaFdetalle.TextMatrix(0, 1)
vdato = GrillaFdetalle.TextMatrix(GrillaFdetalle.Row, 1)

vsql = "delete from " + c3 + " where " + vcampo + " = " + vdato

Call EjecutarScript(vsql, pathDBMySQL)

Call PusFiltrarMovimentos_Click
End Sub

Private Sub PushButton15_Click()
Dim vsql As String
Dim v As String
Dim vcampo, vdato As String
vcampo = GrillaRetenciones.TextMatrix(0, 1)
vdato = GrillaRetenciones.TextMatrix(GrillaRetenciones.Row, 1)

vsql = "delete from " + c8 + " where " + vcampo + " = " + vdato

Call EjecutarScript(vsql, pathDBMySQL)

Call PusFiltrarMovimentos_Click
End Sub

Private Sub PushButton3_Click()
Dim vsql As String
Dim v As String
Dim vcampo, vdato As String
vcampo = GrillaLibroIva.TextMatrix(0, 1)
vdato = GrillaLibroIva.TextMatrix(GrillaLibroIva.Row, 1)

vsql = "delete from " + c4 + " where " + vcampo + " = " + vdato

Call EjecutarScript(vsql, pathDBMySQL)

Call PusFiltrarMovimentos_Click
End Sub

Private Sub PushButton5_Click()
Dim vsql As String
Dim v As String
Dim vcampo, vdato As String
vcampo = GrillaAsientos.TextMatrix(0, 1)
vdato = GrillaAsientos.TextMatrix(GrillaAsientos.Row, 1)

vsql = "delete from " + c5 + " where " + vcampo + " = " + vdato

Call EjecutarScript(vsql, pathDBMySQL)

Call PusFiltrarMovimentos_Click
End Sub

Private Sub PushButton6_Click()
frmBancoCajaDetalle.vnrointerno = Me.txtvnrointerno
Call frmBancoCajaDetalle.cmdFiltrar_Click
frmBancoCajaDetalle.Show
End Sub

Private Sub PushButton7_Click()
Dim vsql As String
Dim v As String
Dim vcampo, vdato As String
vcampo = GrillaMovimientosBanco.TextMatrix(0, 1)
vdato = GrillaMovimientosBanco.TextMatrix(GrillaMovimientosBanco.Row, 1)

vsql = "delete from " + c2 + " where " + vcampo + " = " + vdato

Call EjecutarScript(vsql, pathDBMySQL)

Call PusFiltrarMovimentos_Click
End Sub

Private Sub PushButton8_Click()
If Not vnrointernocheque2 = 0 Then
    Me.txtvnrointerno = vnrointernocheque2
End If
frmCheques.txtFicha(4).Text = Me.txtvnrointerno
frmCheques.txtFicha(5).Text = Me.txtvnrointerno

Call frmCheques.PBFiltrar_Click

frmCheques.Show
frmCheques.WindowState = vmaximizar

End Sub


Private Sub llenarGrillaChequeMovBanco(vidcheque As Long)
Dim vsql As String

        vsql = "select * from cheques where idcheques=" + Str(vidcheque)
        Me.idAsientos = Val(traerDatos2(vsql, "idAsientos", pathDBMySQL))
        
        vnrointernocheque2 = Val(traerDatos2(vsql, "nrointerno", pathDBMySQL))
        Call LlenarGrilla("cheques", Me.GrillaCheques, vsql, "")
End Sub

Private Sub PushButton9_Click()
Dim vsql As String
Dim v As String
Dim vcampo, vdato As String
vcampo = GrillaCheques.TextMatrix(0, 1)
vdato = GrillaCheques.TextMatrix(GrillaCheques.Row, 1)

vsql = "delete from " + c6 + " where " + vcampo + " = " + vdato

Call EjecutarScript(vsql, pathDBMySQL)

Call PusFiltrarMovimentos_Click
End Sub
