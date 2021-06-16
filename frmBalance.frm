VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.ShortcutBar.v13.0.0.Demo.ocx"
Object = "{9746E3DA-06E1-4D26-9CE4-D9F6411A9C70}#1.0#0"; "SMGA_OcxTxt2008.ocx"
Begin VB.Form frmBalance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Balance general de Sumas y Saldos"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   16005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   16005
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   8235
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   15945
      _Version        =   851968
      _ExtentX        =   28125
      _ExtentY        =   14526
      _StockProps     =   68
      ItemCount       =   5
      SelectedItem    =   1
      Item(0).Caption =   "Generar Balance"
      Item(0).ControlCount=   9
      Item(0).Control(0)=   "GBParametros"
      Item(0).Control(1)=   "Barra"
      Item(0).Control(2)=   "GroupBox3"
      Item(0).Control(3)=   "GroupBox2"
      Item(0).Control(4)=   "GroupBox1"
      Item(0).Control(5)=   "GroupBox8"
      Item(0).Control(6)=   "GroupBox9"
      Item(0).Control(7)=   "Frame6"
      Item(0).Control(8)=   "vframetitulo2"
      Item(1).Caption =   "Cierre de Balance"
      Item(1).ControlCount=   34
      Item(1).Control(0)=   "Frame1"
      Item(1).Control(1)=   "Barra2"
      Item(1).Control(2)=   "lblIngresoDe"
      Item(1).Control(3)=   "vDescCca"
      Item(1).Control(4)=   "vCodCca"
      Item(1).Control(5)=   "Label1"
      Item(1).Control(6)=   "Label2"
      Item(1).Control(7)=   "vleyendaCierre"
      Item(1).Control(8)=   "vleyendaApertura"
      Item(1).Control(9)=   "bpca"
      Item(1).Control(10)=   "PushButton3"
      Item(1).Control(11)=   "vDnroBalanceCierre"
      Item(1).Control(12)=   "vCnroBalanceCierre"
      Item(1).Control(13)=   "Label3"
      Item(1).Control(14)=   "PushButton4"
      Item(1).Control(15)=   "vDnroBalanceApertura"
      Item(1).Control(16)=   "vCnroBalanceApertura"
      Item(1).Control(17)=   "Label4"
      Item(1).Control(18)=   "log2"
      Item(1).Control(19)=   "vfbdesde"
      Item(1).Control(20)=   "vfbHasta"
      Item(1).Control(21)=   "Label5"
      Item(1).Control(22)=   "Label6"
      Item(1).Control(23)=   "Label7"
      Item(1).Control(24)=   "vnrointernoC"
      Item(1).Control(25)=   "Label8"
      Item(1).Control(26)=   "vnrointernoA"
      Item(1).Control(27)=   "c"
      Item(1).Control(28)=   "ShortcutCaption2"
      Item(1).Control(29)=   "log3"
      Item(1).Control(30)=   "GroupBox5"
      Item(1).Control(31)=   "vfcierre"
      Item(1).Control(32)=   "vfapertura"
      Item(1).Control(33)=   "GroupBox6"
      Item(2).Caption =   "Ver cálculo de las cuentas"
      Item(2).ControlCount=   0
      Item(3).Caption =   "Balances Ctas Virtuales"
      Item(3).ControlCount=   8
      Item(3).Control(0)=   "Frame3"
      Item(3).Control(1)=   "Frame4"
      Item(3).Control(2)=   "Frame2"
      Item(3).Control(3)=   "b4"
      Item(3).Control(4)=   "b3"
      Item(3).Control(5)=   "Label16"
      Item(3).Control(6)=   "vcomentario"
      Item(3).Control(7)=   "GroupBox7"
      Item(4).Caption =   "Conciliar Módulos - CtasCtables"
      Item(4).ControlCount=   6
      Item(4).Control(0)=   "Label18"
      Item(4).Control(1)=   "Label19"
      Item(4).Control(2)=   "vfcdesde"
      Item(4).Control(3)=   "vfchasta"
      Item(4).Control(4)=   "GroupBox4"
      Item(4).Control(5)=   "vdisplay"
      Begin XtremeSuiteControls.GroupBox vframetitulo2 
         Height          =   675
         Left            =   -69880
         TabIndex        =   146
         Top             =   6330
         Visible         =   0   'False
         Width           =   10275
         _Version        =   851968
         _ExtentX        =   18124
         _ExtentY        =   1191
         _StockProps     =   79
         Caption         =   "Ingre un comentario para agregarle al título del listado "
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.FlatEdit vtitulo2 
            Height          =   285
            Left            =   60
            TabIndex        =   147
            Top             =   270
            Width           =   10095
            _Version        =   851968
            _ExtentX        =   17806
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Tipo de Listado 1:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   585
         Left            =   -69910
         TabIndex        =   142
         Top             =   3360
         Visible         =   0   'False
         Width           =   10305
         Begin XtremeSuiteControls.CheckBox chkTodasLasCtas 
            Height          =   255
            Left            =   8610
            TabIndex        =   149
            Top             =   210
            Width           =   1485
            _Version        =   851968
            _ExtentX        =   2619
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Todas las ctas"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.OptionButton opMTodas 
            Caption         =   "Sumas y Saldos"
            Height          =   285
            Left            =   1770
            TabIndex        =   145
            Top             =   210
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton opPatrimonio 
            Caption         =   "Estado de Situación  Patrimonial"
            Height          =   285
            Left            =   3540
            TabIndex        =   144
            Top             =   210
            Width           =   2595
         End
         Begin VB.OptionButton opResultados 
            Caption         =   "Estado de Resultado"
            Height          =   285
            Left            =   6300
            TabIndex        =   143
            Top             =   210
            Width           =   1965
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox9 
         Height          =   6555
         Left            =   -59230
         TabIndex        =   138
         Top             =   540
         Visible         =   0   'False
         Width           =   4935
         _Version        =   851968
         _ExtentX        =   8705
         _ExtentY        =   11562
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin VB.ListBox log 
            BackColor       =   &H00808080&
            ForeColor       =   &H00C0C0C0&
            Height          =   4740
            Left            =   90
            TabIndex        =   148
            Top             =   1710
            Width           =   4725
         End
         Begin VB.ListBox log22 
            Height          =   1425
            Left            =   90
            TabIndex        =   139
            Top             =   240
            Width           =   4695
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox7 
         Height          =   195
         Left            =   -69790
         TabIndex        =   128
         Top             =   7410
         Visible         =   0   'False
         Width           =   15645
         _Version        =   851968
         _ExtentX        =   27596
         _ExtentY        =   344
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
      End
      Begin XtremeSuiteControls.GroupBox GroupBox6 
         Height          =   195
         Left            =   30
         TabIndex        =   127
         Top             =   7500
         Width           =   15885
         _Version        =   851968
         _ExtentX        =   28019
         _ExtentY        =   344
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
      End
      Begin VB.Frame c 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   30
         TabIndex        =   118
         Top             =   7440
         Width           =   15795
         Begin XtremeSuiteControls.PushButton pb1 
            Height          =   345
            Left            =   60
            TabIndex        =   119
            Top             =   300
            Width           =   3135
            _Version        =   851968
            _ExtentX        =   5530
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "Confeccionar asientos Cierre"
            ForeColor       =   0
            UseVisualStyle  =   -1  'True
            Picture         =   "frmBalance.frx":0000
         End
         Begin XtremeSuiteControls.PushButton PushButton2 
            Height          =   345
            Left            =   3300
            TabIndex        =   120
            Top             =   300
            Width           =   3165
            _Version        =   851968
            _ExtentX        =   5583
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "Aplicar los asientos de cierre"
            ForeColor       =   0
            UseVisualStyle  =   -1  'True
            Picture         =   "frmBalance.frx":0B4A
         End
         Begin XtremeSuiteControls.PushButton PushButton19 
            Height          =   345
            Left            =   9630
            TabIndex        =   121
            Top             =   270
            Width           =   3135
            _Version        =   851968
            _ExtentX        =   5530
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "Confeccionar asientos Apertura"
            ForeColor       =   0
            UseVisualStyle  =   -1  'True
            Picture         =   "frmBalance.frx":1DD4
         End
         Begin XtremeSuiteControls.PushButton PushButton20 
            Height          =   345
            Left            =   12780
            TabIndex        =   122
            Top             =   270
            Width           =   2955
            _Version        =   851968
            _ExtentX        =   5212
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "Aplicar los asientos de apertura"
            ForeColor       =   0
            UseVisualStyle  =   -1  'True
            Picture         =   "frmBalance.frx":291E
         End
      End
      Begin MSComCtl2.DTPicker vfapertura 
         Height          =   285
         Left            =   8160
         TabIndex        =   112
         Top             =   2910
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   503
         _Version        =   393216
         Format          =   84082689
         CurrentDate     =   41237
      End
      Begin MSComCtl2.DTPicker vfcierre 
         Height          =   285
         Left            =   2610
         TabIndex        =   111
         Top             =   2880
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   503
         _Version        =   393216
         Format          =   84082689
         CurrentDate     =   41237
      End
      Begin XtremeSuiteControls.GroupBox GroupBox5 
         Height          =   1275
         Left            =   10500
         TabIndex        =   103
         Top             =   960
         Width           =   5235
         _Version        =   851968
         _ExtentX        =   9234
         _ExtentY        =   2249
         _StockProps     =   79
         Caption         =   "Calcular totales: "
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.PushButton PushButton13 
            Height          =   255
            Left            =   420
            TabIndex        =   104
            Top             =   900
            Width           =   4605
            _Version        =   851968
            _ExtentX        =   8123
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Volver a Calcular Totales"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit vtHaber 
            Height          =   285
            Left            =   1920
            TabIndex        =   105
            Top             =   510
            Width           =   1365
            _Version        =   851968
            _ExtentX        =   2408
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vtDebe 
            Height          =   285
            Left            =   450
            TabIndex        =   106
            Top             =   510
            Width           =   1395
            _Version        =   851968
            _ExtentX        =   2461
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vtSaldo 
            Height          =   285
            Left            =   3330
            TabIndex        =   107
            Top             =   510
            Width           =   1725
            _Version        =   851968
            _ExtentX        =   3043
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin VB.Label Label22 
            Caption         =   "Saldo"
            Height          =   225
            Left            =   3330
            TabIndex        =   110
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label21 
            Caption         =   "Haber"
            Height          =   225
            Left            =   1950
            TabIndex        =   109
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "Debe"
            Height          =   225
            Left            =   450
            TabIndex        =   108
            Top             =   270
            Width           =   1335
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox4 
         Height          =   735
         Left            =   -69880
         TabIndex        =   96
         Top             =   1980
         Visible         =   0   'False
         Width           =   15675
         _Version        =   851968
         _ExtentX        =   27649
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Ejecute la conciliación de las  ctas  contables contra los siguientes módulos:"
         ForeColor       =   0
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton PushButton15 
            Height          =   375
            Left            =   120
            TabIndex        =   97
            Top             =   270
            Width           =   1905
            _Version        =   851968
            _ExtentX        =   3360
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Banco/Caja"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PushButton16 
            Height          =   375
            Left            =   2040
            TabIndex        =   98
            Top             =   270
            Width           =   1815
            _Version        =   851968
            _ExtentX        =   3201
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Proveedores"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PushButton17 
            Height          =   375
            Left            =   3870
            TabIndex        =   99
            Top             =   270
            Width           =   1815
            _Version        =   851968
            _ExtentX        =   3201
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Clientes"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PushButton18 
            Height          =   375
            Left            =   5700
            TabIndex        =   100
            Top             =   270
            Width           =   1815
            _Version        =   851968
            _ExtentX        =   3201
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Cheques"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.FlatEdit vcomentario 
         Height          =   285
         Left            =   -66940
         TabIndex        =   91
         Top             =   1200
         Visible         =   0   'False
         Width           =   12675
         _Version        =   851968
         _ExtentX        =   22357
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin VB.ListBox log3 
         Height          =   1815
         Left            =   10560
         TabIndex        =   81
         Top             =   5520
         Width           =   5175
      End
      Begin VB.Frame Frame4 
         Height          =   2235
         Left            =   -69880
         TabIndex        =   59
         Top             =   5040
         Visible         =   0   'False
         Width           =   15795
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gVirtualDetalle 
            Height          =   1635
            Left            =   60
            TabIndex        =   60
            Top             =   510
            Width           =   15675
            _ExtentX        =   27649
            _ExtentY        =   2884
            _Version        =   393216
            Cols            =   3
            ForeColorSel    =   255
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).BandIndent=   2
            _Band(0).Cols   =   3
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin XtremeSuiteControls.PushButton PushButton9 
            Height          =   285
            Left            =   14670
            TabIndex        =   76
            Top             =   180
            Width           =   1035
            _Version        =   851968
            _ExtentX        =   1826
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Borrar"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PushButton12 
            Height          =   285
            Left            =   13620
            TabIndex        =   77
            Top             =   180
            Width           =   1035
            _Version        =   851968
            _ExtentX        =   1826
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Ordenar"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
            Height          =   315
            Left            =   60
            TabIndex        =   61
            Top             =   150
            Width           =   13485
            _Version        =   851968
            _ExtentX        =   23786
            _ExtentY        =   556
            _StockProps     =   14
            Caption         =   "Cuentas cotables reales asociadas a la cuenta viertual seleccionada"
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
      End
      Begin VB.Frame Frame3 
         Height          =   2985
         Left            =   -69880
         TabIndex        =   56
         Top             =   2010
         Visible         =   0   'False
         Width           =   15705
         Begin XtremeSuiteControls.FlatEdit vVirtual 
            Height          =   285
            Left            =   1800
            TabIndex        =   65
            Top             =   210
            Width           =   4305
            _Version        =   851968
            _ExtentX        =   7594
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.PushButton PushButton6 
            Height          =   285
            Left            =   6300
            TabIndex        =   57
            Top             =   210
            Width           =   705
            _Version        =   851968
            _ExtentX        =   1244
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Agregar"
            UseVisualStyle  =   -1  'True
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gVirtual 
            Height          =   1725
            Left            =   90
            TabIndex        =   58
            Top             =   660
            Width           =   12135
            _ExtentX        =   21405
            _ExtentY        =   3043
            _Version        =   393216
            Cols            =   5
            ForeColorSel    =   255
            SelectionMode   =   1
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).BandIndent=   2
            _Band(0).Cols   =   5
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin XtremeSuiteControls.PushButton PushButton8 
            Height          =   285
            Left            =   7050
            TabIndex        =   66
            Top             =   210
            Width           =   675
            _Version        =   851968
            _ExtentX        =   1191
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Borrar"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PushButton10 
            Height          =   345
            Left            =   12570
            TabIndex        =   71
            Top             =   2520
            Width           =   3045
            _Version        =   851968
            _ExtentX        =   5371
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "Agregar cuenta real"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PushButton11 
            Height          =   315
            Left            =   4350
            TabIndex        =   72
            Tag             =   "Banco"
            Top             =   2580
            Width           =   315
            _Version        =   851968
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit vDCtaReal 
            Height          =   315
            Left            =   4710
            TabIndex        =   73
            Top             =   2580
            Width           =   7485
            _Version        =   851968
            _ExtentX        =   13203
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vCctaReal 
            Height          =   315
            Left            =   2910
            TabIndex        =   74
            Top             =   2580
            Width           =   1395
            _Version        =   851968
            _ExtentX        =   2461
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
            Alignment       =   1
         End
         Begin VB.Line Line1 
            X1              =   12840
            X2              =   15630
            Y1              =   1590
            Y2              =   1590
         End
         Begin XtremeSuiteControls.Label Label17 
            Height          =   285
            Left            =   12900
            TabIndex        =   89
            Top             =   1650
            Width           =   885
            _Version        =   851968
            _ExtentX        =   1561
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Diferencia:"
         End
         Begin XtremeSuiteControls.Label vresta 
            Height          =   285
            Left            =   13980
            TabIndex        =   88
            Top             =   1680
            Width           =   945
            _Version        =   851968
            _ExtentX        =   1667
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "0"
            ForeColor       =   32768
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
         End
         Begin XtremeSuiteControls.Label vtthaber 
            Height          =   285
            Left            =   13980
            TabIndex        =   87
            Top             =   1170
            Width           =   945
            _Version        =   851968
            _ExtentX        =   1667
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "0"
            ForeColor       =   255
            Alignment       =   1
         End
         Begin XtremeSuiteControls.Label vttdebe 
            Height          =   285
            Left            =   13980
            TabIndex        =   86
            Top             =   780
            Width           =   945
            _Version        =   851968
            _ExtentX        =   1667
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "0"
            ForeColor       =   16711680
            Alignment       =   1
         End
         Begin XtremeSuiteControls.Label Label15 
            Height          =   285
            Left            =   12870
            TabIndex        =   85
            Top             =   1200
            Width           =   885
            _Version        =   851968
            _ExtentX        =   1561
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Total haber:"
         End
         Begin XtremeSuiteControls.Label Label14 
            Height          =   285
            Left            =   12870
            TabIndex        =   84
            Top             =   780
            Width           =   885
            _Version        =   851968
            _ExtentX        =   1561
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Total debe :"
         End
         Begin XtremeSuiteControls.Label Label13 
            Height          =   285
            Left            =   10860
            TabIndex        =   78
            Top             =   210
            Width           =   4275
            _Version        =   851968
            _ExtentX        =   7541
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Seleccionar cuenta virtual haciendo doble clic sobre la fila"
         End
         Begin XtremeSuiteControls.Label Label12 
            Height          =   345
            Left            =   150
            TabIndex        =   75
            Top             =   2550
            Width           =   2475
            _Version        =   851968
            _ExtentX        =   4366
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "> Seleccionar cuenta contable: "
         End
         Begin XtremeSuiteControls.Label Label9 
            Height          =   285
            Left            =   120
            TabIndex        =   64
            Top             =   210
            Width           =   1695
            _Version        =   851968
            _ExtentX        =   2990
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "> Ing. cuenta virtual: "
         End
      End
      Begin XtremeSuiteControls.FlatEdit vnrointernoA 
         Height          =   285
         Left            =   8130
         TabIndex        =   55
         Top             =   3540
         Width           =   2055
         _Version        =   851968
         _ExtentX        =   3625
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vnrointernoC 
         Height          =   285
         Left            =   2610
         TabIndex        =   53
         Top             =   3510
         Width           =   1965
         _Version        =   851968
         _ExtentX        =   3466
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin MSComCtl2.DTPicker vfbdesde 
         Height          =   285
         Left            =   2610
         TabIndex        =   48
         Top             =   3180
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   503
         _Version        =   393216
         Format          =   84082689
         CurrentDate     =   41237
      End
      Begin VB.ListBox log2 
         Height          =   1815
         Left            =   10560
         TabIndex        =   44
         Top             =   3480
         Width           =   5115
      End
      Begin XtremeSuiteControls.FlatEdit vleyendaCierre 
         Height          =   285
         Left            =   2550
         TabIndex        =   34
         Top             =   1350
         Width           =   4425
         _Version        =   851968
         _ExtentX        =   7805
         _ExtentY        =   503
         _StockProps     =   77
         ForeColor       =   255
         BackColor       =   -2147483643
         Text            =   "Cierre de ejercicio"
      End
      Begin VB.Frame Frame1 
         Caption         =   "Resultados:"
         Height          =   3165
         Left            =   0
         TabIndex        =   13
         Top             =   4110
         Width           =   10275
         Begin XtremeSuiteControls.PushButton PushButton5 
            Height          =   285
            Left            =   8130
            TabIndex        =   47
            Top             =   360
            Width           =   1785
            _Version        =   851968
            _ExtentX        =   3149
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Confirmar cambio"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit vcelda 
            Height          =   315
            Left            =   3540
            TabIndex        =   46
            Top             =   300
            Width           =   2985
            _Version        =   851968
            _ExtentX        =   5265
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid gBalance 
            Height          =   2025
            Left            =   120
            TabIndex        =   21
            Top             =   960
            Width           =   10005
            _ExtentX        =   17648
            _ExtentY        =   3572
            _Version        =   393216
            Cols            =   4
            ForeColorSel    =   255
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).BandIndent=   2
            _Band(0).Cols   =   4
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   555
            Left            =   60
            TabIndex        =   45
            Top             =   210
            Width           =   10095
            _Version        =   851968
            _ExtentX        =   17806
            _ExtentY        =   979
            _StockProps     =   14
            Caption         =   "Cambiar el valor de la selda seleccionada: "
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
      End
      Begin MSComctlLib.ProgressBar Barra 
         Height          =   375
         Left            =   -69910
         TabIndex        =   12
         Top             =   7200
         Visible         =   0   'False
         Width           =   15690
         _ExtentX        =   27675
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   645
         Left            =   -69880
         TabIndex        =   22
         Top             =   7530
         Visible         =   0   'False
         Width           =   15675
         _Version        =   851968
         _ExtentX        =   27649
         _ExtentY        =   1138
         _StockProps     =   79
         BackColor       =   -2147483644
         UseVisualStyle  =   -1  'True
         BorderStyle     =   1
         Begin VB.CommandButton Command2 
            Caption         =   "Imprimir Última Generación de Balance"
            Height          =   345
            Left            =   1500
            Picture         =   "frmBalance.frx":3BA8
            TabIndex        =   129
            Top             =   240
            Width           =   3135
         End
         Begin XtremeSuiteControls.PushButton PbAcciones 
            Height          =   345
            Index           =   0
            Left            =   60
            TabIndex        =   23
            Top             =   240
            Width           =   1095
            _Version        =   851968
            _ExtentX        =   1931
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "Generar"
            UseVisualStyle  =   -1  'True
            Picture         =   "frmBalance.frx":4132
         End
         Begin XtremeSuiteControls.PushButton PbAcciones 
            Height          =   345
            Index           =   1
            Left            =   14340
            TabIndex        =   24
            Top             =   -630
            Visible         =   0   'False
            Width           =   1095
            _Version        =   851968
            _ExtentX        =   1931
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "Cerrar"
            UseVisualStyle  =   -1  'True
            Picture         =   "frmBalance.frx":4C7C
         End
      End
      Begin XtremeSuiteControls.ProgressBar Barra2 
         Height          =   165
         Left            =   120
         TabIndex        =   25
         Top             =   3870
         Width           =   10095
         _Version        =   851968
         _ExtentX        =   17806
         _ExtentY        =   291
         _StockProps     =   93
         Text            =   "Barra2"
         BackColor       =   -2147483644
         BarColor        =   255
      End
      Begin XtremeSuiteControls.PushButton bpca 
         Height          =   315
         Left            =   4110
         TabIndex        =   28
         Tag             =   "Banco"
         Top             =   960
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit vDescCca 
         Height          =   315
         Left            =   4500
         TabIndex        =   29
         Top             =   960
         Width           =   2445
         _Version        =   851968
         _ExtentX        =   4313
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Text            =   "DEL EJERCICIO"
      End
      Begin XtremeSuiteControls.FlatEdit vCodCca 
         Height          =   315
         Left            =   2520
         TabIndex        =   30
         Top             =   960
         Width           =   1515
         _Version        =   851968
         _ExtentX        =   2672
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Text            =   "310402"
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit vleyendaApertura 
         Height          =   285
         Left            =   2550
         TabIndex        =   35
         Top             =   1710
         Width           =   4425
         _Version        =   851968
         _ExtentX        =   7805
         _ExtentY        =   503
         _StockProps     =   77
         ForeColor       =   16711680
         BackColor       =   -2147483643
         Text            =   "Apertura de ejercicio"
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   315
         Left            =   4140
         TabIndex        =   36
         Tag             =   "Banco"
         Top             =   2040
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit vDnroBalanceCierre 
         Height          =   315
         Left            =   4530
         TabIndex        =   37
         Top             =   2040
         Width           =   2445
         _Version        =   851968
         _ExtentX        =   4313
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vCnroBalanceCierre 
         Height          =   315
         Left            =   3210
         TabIndex        =   38
         Top             =   2070
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   315
         Left            =   4140
         TabIndex        =   40
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
      Begin XtremeSuiteControls.FlatEdit vDnroBalanceApertura 
         Height          =   315
         Left            =   4530
         TabIndex        =   41
         Top             =   2430
         Width           =   2445
         _Version        =   851968
         _ExtentX        =   4313
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vCnroBalanceApertura 
         Height          =   315
         Left            =   3210
         TabIndex        =   42
         Top             =   2430
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin MSComCtl2.DTPicker vfbHasta 
         Height          =   285
         Left            =   8160
         TabIndex        =   49
         Top             =   3240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   503
         _Version        =   393216
         Format          =   84082689
         CurrentDate     =   41237
      End
      Begin XtremeSuiteControls.ProgressBar b4 
         Height          =   495
         Left            =   -66340
         TabIndex        =   79
         Top             =   480
         Visible         =   0   'False
         Width           =   12105
         _Version        =   851968
         _ExtentX        =   21352
         _ExtentY        =   873
         _StockProps     =   93
         Text            =   "B4"
         BackColor       =   1375373
         Appearance      =   6
         UseVisualStyle  =   0   'False
         BarColor        =   49152
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   -69910
         TabIndex        =   62
         Top             =   7380
         Visible         =   0   'False
         Width           =   15585
         Begin MSComCtl2.DTPicker vfvdesde 
            Height          =   285
            Left            =   1530
            TabIndex        =   67
            Top             =   270
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   503
            _Version        =   393216
            Format          =   84082689
            CurrentDate     =   41239
         End
         Begin XtremeSuiteControls.PushButton PushButton7 
            Height          =   345
            Left            =   12210
            TabIndex        =   63
            Top             =   270
            Width           =   1695
            _Version        =   851968
            _ExtentX        =   2990
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "Generarl Balance"
            UseVisualStyle  =   -1  'True
            Picture         =   "frmBalance.frx":507C
         End
         Begin MSComCtl2.DTPicker vfvhasta 
            Height          =   285
            Left            =   4590
            TabIndex        =   68
            Top             =   300
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   503
            _Version        =   393216
            Format          =   84082689
            CurrentDate     =   41239
         End
         Begin XtremeSuiteControls.PushButton PushButton14 
            Height          =   345
            Left            =   13950
            TabIndex        =   83
            Top             =   270
            Width           =   1605
            _Version        =   851968
            _ExtentX        =   2831
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "Imprimir"
            UseVisualStyle  =   -1  'True
            Picture         =   "frmBalance.frx":5BC6
         End
         Begin XtremeSuiteControls.Label Label11 
            Height          =   225
            Left            =   3420
            TabIndex        =   70
            Top             =   330
            Width           =   1125
            _Version        =   851968
            _ExtentX        =   1984
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "> Fecha Hasta:"
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   225
            Left            =   210
            TabIndex        =   69
            Top             =   330
            Width           =   1155
            _Version        =   851968
            _ExtentX        =   2037
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "> Fecha desde:"
         End
      End
      Begin MSComCtl2.DTPicker vfcdesde 
         Height          =   285
         Left            =   -68350
         TabIndex        =   92
         Top             =   690
         Visible         =   0   'False
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   503
         _Version        =   393216
         Format          =   84082689
         CurrentDate     =   41239
      End
      Begin MSComCtl2.DTPicker vfchasta 
         Height          =   285
         Left            =   -68350
         TabIndex        =   93
         Top             =   1320
         Visible         =   0   'False
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   503
         _Version        =   393216
         Format          =   84082689
         CurrentDate     =   41239
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   405
         Left            =   -69880
         TabIndex        =   123
         Top             =   4200
         Visible         =   0   'False
         Width           =   10245
         _Version        =   851968
         _ExtentX        =   18071
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Tipos de Asientos: "
         ForeColor       =   4210752
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
            Left            =   3540
            TabIndex        =   124
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
            Left            =   7020
            TabIndex        =   125
            Top             =   150
            Width           =   1485
            _Version        =   851968
            _ExtentX        =   2619
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Asiento normal"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rdtodos 
            Height          =   165
            Left            =   1740
            TabIndex        =   126
            Top             =   210
            Width           =   1695
            _Version        =   851968
            _ExtentX        =   2990
            _ExtentY        =   291
            _StockProps     =   79
            Caption         =   "Todos los asientos"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   405
         Left            =   -69880
         TabIndex        =   130
         Top             =   4950
         Visible         =   0   'False
         Width           =   10245
         _Version        =   851968
         _ExtentX        =   18071
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Tipo de Listado 2:"
         ForeColor       =   4210752
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
         Begin XtremeSuiteControls.RadioButton rdPresupuestado 
            Height          =   195
            Left            =   3540
            TabIndex        =   131
            Top             =   180
            Width           =   1485
            _Version        =   851968
            _ExtentX        =   2619
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Presupuestado"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rdProyectado 
            Height          =   255
            Left            =   6990
            TabIndex        =   132
            Top             =   180
            Width           =   1485
            _Version        =   851968
            _ExtentX        =   2619
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Proyectado"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rdBalance 
            Height          =   165
            Left            =   1710
            TabIndex        =   133
            Top             =   180
            Width           =   1155
            _Version        =   851968
            _ExtentX        =   2037
            _ExtentY        =   291
            _StockProps     =   79
            Caption         =   "Balance"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox8 
         Height          =   405
         Left            =   -69880
         TabIndex        =   134
         Top             =   5670
         Visible         =   0   'False
         Width           =   10245
         _Version        =   851968
         _ExtentX        =   18071
         _ExtentY        =   714
         _StockProps     =   79
         Caption         =   "Filtros:"
         ForeColor       =   4210752
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
         Begin XtremeSuiteControls.RadioButton rbCtaCSaldo 
            Height          =   195
            Left            =   3540
            TabIndex        =   135
            Top             =   180
            Width           =   1815
            _Version        =   851968
            _ExtentX        =   3201
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Solo Ctas con Saldo"
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbCtaSSaldo 
            Height          =   255
            Left            =   6990
            TabIndex        =   136
            Top             =   150
            Width           =   1905
            _Version        =   851968
            _ExtentX        =   3360
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Solo Ctas sin Saldo"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbCtaTodo 
            Height          =   195
            Left            =   1710
            TabIndex        =   137
            Top             =   180
            Width           =   1155
            _Version        =   851968
            _ExtentX        =   2037
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Todo"
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GBParametros 
         Height          =   2775
         Left            =   -69880
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   10290
         _Version        =   851968
         _ExtentX        =   18150
         _ExtentY        =   4895
         _StockProps     =   79
         Caption         =   "Parametros para la generacion del Balance"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.PushButton PusEsteMes 
            Height          =   285
            Left            =   3450
            TabIndex        =   140
            Top             =   450
            Width           =   1425
            _Version        =   851968
            _ExtentX        =   2514
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Este mes"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PushButton21 
            Height          =   315
            Left            =   7950
            TabIndex        =   117
            Top             =   2850
            Visible         =   0   'False
            Width           =   2145
            _Version        =   851968
            _ExtentX        =   3784
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Prueba"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Balance Gral"
            Height          =   345
            Left            =   7920
            TabIndex        =   116
            Top             =   540
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Frame Frame5 
            Caption         =   "Tipo de Balace"
            ForeColor       =   &H00404040&
            Height          =   615
            Left            =   4890
            TabIndex        =   113
            Top             =   180
            Visible         =   0   'False
            Width           =   2865
            Begin VB.OptionButton rbgral 
               Caption         =   "General"
               Height          =   225
               Left            =   1740
               TabIndex        =   115
               Top             =   330
               Value           =   -1  'True
               Width           =   1515
            End
            Begin VB.OptionButton rbss 
               Caption         =   "Sumas y Saldos"
               Height          =   225
               Left            =   180
               TabIndex        =   114
               Top             =   330
               Width           =   1515
            End
         End
         Begin XtremeSuiteControls.CheckBox chkvarios 
            Height          =   405
            Left            =   7950
            TabIndex        =   102
            Top             =   120
            Visible         =   0   'False
            Width           =   2145
            _Version        =   851968
            _ExtentX        =   3784
            _ExtentY        =   714
            _StockProps     =   79
            Caption         =   "Permitir varios ejercicios"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PushButton1 
            Height          =   285
            Left            =   3360
            TabIndex        =   20
            Top             =   1950
            Width           =   315
            _Version        =   851968
            _ExtentX        =   556
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "..."
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit vcbalance 
            Height          =   285
            Left            =   2160
            TabIndex        =   19
            Top             =   1950
            Width           =   1065
            _Version        =   851968
            _ExtentX        =   1879
            _ExtentY        =   503
            _StockProps     =   77
            BackColor       =   -2147483643
            Enabled         =   0   'False
         End
         Begin MSComCtl2.DTPicker vvtimestamp 
            Height          =   285
            Left            =   3870
            TabIndex        =   16
            Top             =   2880
            Visible         =   0   'False
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   503
            _Version        =   393216
            Format          =   84082689
            CurrentDate     =   41180
         End
         Begin XtremeSuiteControls.PushButton pbCarga 
            Height          =   315
            Index           =   0
            Left            =   3360
            TabIndex        =   4
            Tag             =   "CodigoCuentaD"
            Top             =   1200
            Width           =   315
            _Version        =   851968
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaContable 
            Height          =   315
            Index           =   0
            Left            =   2175
            TabIndex        =   5
            Top             =   1200
            Width           =   1095
            _Version        =   851968
            _ExtentX        =   1940
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaContable 
            Height          =   315
            Index           =   1
            Left            =   3810
            TabIndex        =   6
            Top             =   1200
            Width           =   3615
            _Version        =   851968
            _ExtentX        =   6376
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaContable 
            Height          =   315
            Index           =   2
            Left            =   2175
            TabIndex        =   7
            Top             =   1560
            Width           =   1095
            _Version        =   851968
            _ExtentX        =   1940
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.PushButton pbCarga 
            Height          =   315
            Index           =   1
            Left            =   3360
            TabIndex        =   8
            Tag             =   "CodigoCuentaH"
            Top             =   1560
            Width           =   315
            _Version        =   851968
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaContable 
            Height          =   315
            Index           =   3
            Left            =   3810
            TabIndex        =   9
            Top             =   1560
            Width           =   3615
            _Version        =   851968
            _ExtentX        =   6376
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin Aplisoft_CajasDeTexto.TxF txtnrobalance 
            Height          =   300
            Left            =   5340
            TabIndex        =   14
            Top             =   1980
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
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
            Height          =   300
            Index           =   0
            Left            =   2160
            TabIndex        =   0
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackStyle       =   0
         End
         Begin Aplisoft_CajasDeTexto.TxF dtpFecha 
            Height          =   300
            Index           =   1
            Left            =   2175
            TabIndex        =   1
            Top             =   840
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
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
         Begin XtremeSuiteControls.PushButton PusMesAnterior 
            Height          =   285
            Left            =   3450
            TabIndex        =   141
            Top             =   750
            Width           =   1425
            _Version        =   851968
            _ExtentX        =   2514
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Mes anterior"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label lblBalance 
            BackStyle       =   0  'Transparent
            Caption         =   "> Fecha Inicio :"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   27
            Top             =   480
            Width           =   1995
         End
         Begin VB.Label lblBalance 
            BackStyle       =   0  'Transparent
            Caption         =   "> Fecha Fin :"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   26
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label lblBalance 
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre del Balance:"
            Height          =   195
            Index           =   6
            Left            =   180
            TabIndex        =   18
            Top             =   1980
            Width           =   1695
         End
         Begin VB.Label lblBalance 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "> Fecha del balance generado con el que quiere comparar este balance:"
            ForeColor       =   &H00404040&
            Height          =   405
            Index           =   5
            Left            =   180
            TabIndex        =   17
            Top             =   2790
            Visible         =   0   'False
            Width           =   3075
         End
         Begin VB.Label lblBalance 
            BackStyle       =   0  'Transparent
            Caption         =   "> Nro. Balance:"
            Height          =   195
            Index           =   4
            Left            =   3900
            TabIndex        =   15
            Top             =   2010
            Width           =   1275
         End
         Begin VB.Label lblBalance 
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Cuenta Hasta :"
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   11
            Top             =   1560
            Width           =   1755
         End
         Begin VB.Label lblBalance 
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Cuenta Desde :"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   10
            Top             =   1200
            Width           =   1995
         End
      End
      Begin XtremeSuiteControls.ProgressBar b3 
         Height          =   495
         Left            =   -69820
         TabIndex        =   80
         Top             =   480
         Visible         =   0   'False
         Width           =   3435
         _Version        =   851968
         _ExtentX        =   6059
         _ExtentY        =   873
         _StockProps     =   93
         Text            =   "B3"
         BackColor       =   16777215
         Scrolling       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
         BarColor        =   16777215
      End
      Begin XtremeSuiteControls.Label vdisplay 
         Height          =   1065
         Left            =   -66640
         TabIndex        =   101
         Top             =   3000
         Visible         =   0   'False
         Width           =   10365
         _Version        =   851968
         _ExtentX        =   18283
         _ExtentY        =   1879
         _StockProps     =   79
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label Label19 
         Height          =   225
         Left            =   -69730
         TabIndex        =   95
         Top             =   720
         Visible         =   0   'False
         Width           =   1155
         _Version        =   851968
         _ExtentX        =   2037
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "> Fecha desde:"
      End
      Begin XtremeSuiteControls.Label Label18 
         Height          =   225
         Left            =   -69730
         TabIndex        =   94
         Top             =   1350
         Visible         =   0   'False
         Width           =   1125
         _Version        =   851968
         _ExtentX        =   1984
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "> Fecha Hasta:"
      End
      Begin VB.Label Label16 
         Caption         =   "Ing. comentario para el listado impreso:"
         Height          =   255
         Left            =   -69820
         TabIndex        =   90
         Top             =   1260
         Visible         =   0   'False
         Width           =   2835
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   285
         Left            =   10530
         TabIndex        =   82
         Top             =   3120
         Width           =   5175
         _Version        =   851968
         _ExtentX        =   9128
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "Resultados de control"
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
      Begin VB.Label Label8 
         Caption         =   "Nro interno para asiento de apertura:"
         Height          =   165
         Left            =   5130
         TabIndex        =   54
         Top             =   3540
         Width           =   2685
      End
      Begin VB.Label Label7 
         Caption         =   "Nro interno para asiento de cierre:"
         Height          =   165
         Left            =   90
         TabIndex        =   52
         Top             =   3540
         Width           =   2445
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha para el asiente de apertura:"
         Height          =   255
         Left            =   4830
         TabIndex        =   51
         Top             =   2940
         Width           =   2985
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha para el asiente de cierre:"
         Height          =   255
         Left            =   90
         TabIndex        =   50
         Top             =   2910
         Width           =   2265
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Nro balance para el asiento de apertura:"
         Height          =   255
         Left            =   60
         TabIndex        =   43
         Top             =   2460
         Width           =   2985
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Nro balance para el asiento del cierre:"
         Height          =   255
         Left            =   270
         TabIndex        =   39
         Top             =   2130
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Leyenda asiento de apertura:"
         Height          =   255
         Left            =   300
         TabIndex        =   33
         Top             =   1770
         Width           =   2085
      End
      Begin VB.Label Label1 
         Caption         =   "Leyenda asiento de cierre:"
         Height          =   255
         Left            =   480
         TabIndex        =   32
         Top             =   1410
         Width           =   1935
      End
      Begin VB.Label lblIngresoDe 
         Caption         =   "Cuenta para el Cierre/Apertura: "
         Height          =   255
         Left            =   150
         TabIndex        =   31
         Top             =   990
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vnrobalance As Integer
Dim vfhasta, vfdesde As Date
Dim vcodigoBalance, vmarca, vlbltipolistado As String
Dim vnroasiento As Long
Dim vanomescierre As String
Public vmesDelBalance As String


Private Sub MostrarReporteBalance(vTipoBalance As Boolean)
On Error Resume Next
Dim vsql3, vsql2, vtinicial, vtDebe, vtHaber, vtperiodo, vtfinal, vcondi, vcondi2, vtacumulado, vtpresupuestado, vtdiferencia As String
Dim vns As String


    Unload Mantenimiento
    Load Mantenimiento
    
    vcondi = ""
    vcondi2 = ""

    MsgBox "Prepare la Impresora!!!", vbInformation, "Mensaje ..."
    
    If Me.chkTodasLasCtas.Value = xtpChecked Then
        vns = " 1=1 "
    Else
        vns = " (c09='N') "
    End If
    
    
    
    If vTipoBalance = True Then
        
        With Mantenimiento.rsBalance
            If Not .State = 0 Then .Close
            
            
            If Me.opMTodas Then
                .Source = "SELECT * FROM Temp2 ORDER BY C02 ASC"
                vlbltipolistado = " - SUMAS Y SALDOS"
            End If
            
            
            If Me.opPatrimonio Then
                vcondi = vcondi + " and " + vns + " and ( (c02 like '1%') or (c02 like '2%') or (c02 like '3%')) "
                vcondi2 = vcondi2 + "  and ( (c02 like '1%') or (c02 like '2%') or (c02 like '3%')) "
                
                vlbltipolistado = " - Estado de SITUACIÓN PATRIMONIAL"
                
                'vcondi = vcondi + " (and c09='S')"
            '    .Source = "SELECT * FROM Temp2 where c10='S' ORDER BY C06 ASC"
            End If
            
            If Me.opResultados Then
                vcondi = vcondi + " and " + vns + " and ( (c02 like '4%') or (c02 like '5%') or (c02 like '6%')) "
                vcondi2 = vcondi2 + " and ( (c02 like '4%') or (c02 like '5%') or (c02 like '6%')) "
                
                vlbltipolistado = " - Estado de RESULTADO"
             '  .Source = "SELECT * FROM Temp2 where c10='N' ORDER BY C02 ASC"
            End If
            
            
            
            If rbCtaCSaldo.Value And Me.rdPresupuestado.Value Then
                vcondi = vcondi + " and (CAST(t.C14 AS UNSIGNED))>0 "
                 'vcondi = vcondi + " and (c13) >0 "   ' si tiene saldo
               
            Else
            
              ' If Me.rbCtaCSaldo.Value Then vcondi = vcondi + " and (CAST(t.C04 AS UNSIGNED)>0  or CAST(t.C03 AS UNSIGNED)>0) "
            
               If Me.rbCtaCSaldo.Value Then vcondi = vcondi + " and  abs(c13) > 0  "
  
            End If
            
            
            .Source = "SELECT * FROM Temp2 t where 1=1 " + vcondi + " ORDER BY C02 ASC"
            
                
            If Not .State = 1 Then .Open
            .Close
            .Open
        
            If .RecordCount = 0 Then Exit Sub
            
        
        End With
    
    
    
     
            vsql2 = "select sum(c07) as c  from temp2 t where t.C09 = 's' " + vcondi2
            vtinicial = Trim(Format(traerDatos2(vsql2, "c", pathDBMySQL), "###,###,##0.00"))
            
            vsql2 = "select sum(c03) as c  from temp2 t where t.C09 = 's' " + vcondi2
            vtDebe = Trim(Format(traerDatos2(vsql2, "c", pathDBMySQL), "###,###,##0.00"))
            
             vsql2 = "select sum(c04) as c  from temp2 t where t.C09 = 's' " + vcondi2
            vtHaber = Trim(Format(traerDatos2(vsql2, "c", pathDBMySQL), "###,###,##0.00"))
            
            vsql2 = "select sum(c11) as c  from temp2 t where t.C09 = 's' " + vcondi2
            vtperiodo = Trim(Format(traerDatos2(vsql2, "c", pathDBMySQL), "###,###,##0.00"))
            
             vsql2 = "select sum(c13) as c  from temp2 t where t.C09 = 's' " + vcondi2
            vtfinal = Trim(Format(traerDatos2(vsql2, "c", pathDBMySQL), "###,###,##0.00"))
            
            vsql2 = "select sum(c14) as c  from temp2 t where t.C09 = 's' " + vcondi2
            vtpresupuestado = Trim(Format(traerDatos2(vsql2, "c", pathDBMySQL), "###,###,##0.00"))
            
            vsql2 = "select sum(c15) as c  from temp2 t where t.C09 = 's' " + vcondi2
            vtdiferencia = Trim(Format(traerDatos2(vsql2, "c", pathDBMySQL), "###,###,##0.00"))
            
            vsql2 = "select sum(c13) as c  from temp2 t where t.C09 = 's' " + vcondi2
            vtacumulado = Trim(Format(traerDatos2(vsql2, "c", pathDBMySQL), "###,###,##0.00"))
            
               
    If rdPresupuestado.Value Then
        With drProyectado
        
           .Sections("totales").Controls("vtinicial").Caption = Format(vtinicial, "###,###,##0.00")
            
            .Sections("totales").Controls("vtperiodo").Caption = Format(vtperiodo, "###,###,##0.00")
    
    
            .Sections("totales").Controls("vtacumulado").Caption = Format(vtacumulado, "###,###,##0.00")
            
       
            .Sections("totales").Controls("vtpresupuestado").Caption = Format(vtpresupuestado, "###,###,##0.00")
            
            .Sections("totales").Controls("vtdiferencia").Caption = Format(vtdiferencia, "###,###,##0.00")
            
          
            
            .Sections("TituloEmpresa").Controls("lblTitulo").Caption = "[ PRESUPUESTADO DESDE: " & dtpFecha(0).Value & " HASTA: " & dtpFecha(1).Value & " ] " + Trim(vtitulo2.Text)
            .Sections("TituloEmpresa").Controls("snombre").Caption = vDatosEmpresa.Nombre
            .Sections("TituloEmpresa").Controls("sdirtel").Caption = vDatosEmpresa.Direccion & "  /  " & vDatosEmpresa.Telefono
            .Sections("TituloEmpresa").Controls("slocalidad").Caption = vDatosEmpresa.Localidad
            .Sections("TituloEmpresa").Controls("semail").Caption = vDatosEmpresa.Email
            
            
        
            
            .Show
        End With
    Else
        With drBalance
 
                
            
            .Sections("totales").Controls("vtinicial").Caption = vtinicial
            
            
            .Sections("totales").Controls("vtdebe").Caption = vtDebe
            
            
            .Sections("totales").Controls("vthaber").Caption = vtHaber
            
           
            .Sections("totales").Controls("vtperiodo").Caption = vtperiodo
            
            
            .Sections("totales").Controls("vtfinal").Caption = vtfinal
            
            
            If vmesDelBalance = "" Then
                .Sections("TituloEmpresa").Controls("lblTitulo").Caption = vlbltipolistado + " [Fecha desde: " & dtpFecha(0).Value & " Hasta: " & dtpFecha(1).Value & " ] " + Trim(vtitulo2.Text)
            Else
                .Sections("TituloEmpresa").Controls("lblTitulo").Caption = vlbltipolistado + vmesDelBalance
            End If
            
            .Sections("TituloEmpresa").Controls("snombre").Caption = vDatosEmpresa.Nombre
            .Sections("TituloEmpresa").Controls("sdirtel").Caption = vDatosEmpresa.Direccion & "  /  " & vDatosEmpresa.Telefono
            .Sections("TituloEmpresa").Controls("slocalidad").Caption = vDatosEmpresa.Localidad
            .Sections("TituloEmpresa").Controls("semail").Caption = vDatosEmpresa.Email
            
            
            If LeerXml("Puesto") = "Caja" Then
                .Sections("TituloEmpresa").Controls("labeldebe").Caption = ""
                .Sections("TituloEmpresa").Controls("labelhaber").Caption = ""
                .Sections(3).Controls("txtdebe").DataField = "c01"
                .Sections(3).Controls("txthaber").DataField = "c01"
            End If
            
            
            .Show
        End With
       End If
    Else
    
    End If

    
If Err Then GrabarLog "MostrarReporte", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub bpca_Click()
Call fbuscarGrilla("Cuentas", "Cuenta", "CodigoCuenta", Me.vDescCca.Name, Me)  ' ema:
End Sub

Private Sub Command1_Click()
Call BalanceGral
End Sub

Private Sub Command2_Click()
Call MostrarReporteBalance(True)
End Sub

Public Sub dtpFecha_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index = 0 Then
        dtpFecha(1).SetFocus
        dtpFecha(1).Value = DiasDelMes(dtpFecha(0).Value) & "/" & AjustarMes(Month(dtpFecha(0).Value)) & "/" & Year(dtpFecha(0).Value)
        'dtpFecha(1).Value = dtpFecha(1).MaxValor
    End If
    
    
    If Index = 1 Then PbAcciones(0).SetFocus
    
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
    
  '  With Me
  '      .Show
  '      .Width = 10605
  '      .Height = 4785
  '      .KeyPreview = True
  '  End With
    
   ' dtpFecha(0).Value = "01/01/" & Year(Date)
   ' dtpFecha(1).Value = "31/12/" & Year(Date)
    

    
    finit
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub finit()

    TabControl1.SelectedItem = 0
    
    vlbltipolistado = ""

    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2 - 1000



    Me.TabControl1.TabIndex = 0


    Me.vfvdesde.Value = Date
    Me.vfvhasta.Value = Date
    
    Me.vfcdesde.Value = Date - 30
    Me.vfchasta.Value = Date
    
   
    
    vvtimestamp.Value = "08-08-2012"
    
     vnrobalance = TraerDato("balances", " Activo='S' order by NroBalance Desc", "NroBalance", pathDBMySQL)
    'vnrobalance = TraerDato("FechaInicio", " Activo='S' order by NroBalance Desc", "NroBalance", pathDBMySQL)
   ' vnrobalance = Val(TraerDato("balances", " Activo='S' order by NroBalance Desc", "NroBalance", pathDBMySQL))
    
    
   
    
    If Me.vcbalance = "" Then
        vcodigoBalance = traerDatos2("select * from balances where Activo='S' order by idBalances desc", "codigo", pathDBMySQL)
        Me.vcbalance.Text = vcodigoBalance
    
    Else
    
       vcodigoBalance = Me.vcbalance.Text
    
    End If
    
    Me.gBalance.TextMatrix(0, 0) = "Cuentas"
    Me.gBalance.TextMatrix(0, 1) = "Debe"
    Me.gBalance.TextMatrix(0, 2) = "Haber"
    Me.gBalance.TextMatrix(0, 3) = "Comentario"
    
    
    
    Me.gBalance.ColWidth(1) = 2000
    Me.gBalance.ColWidth(2) = 2000
    Me.gBalance.ColWidth(3) = 5000
    
    Me.vCnroBalanceCierre.Text = vnrobalance
    Me.vDnroBalanceCierre.Tag = vnrobalance
    Me.vDnroBalanceCierre.Text = vcodigoBalance
    
    'dtpFecha(0).SetFocus
    
    vnrointernoC = traerDatos2("select max(nrointerno) as c from asientos", "c", pathDBMySQL) + 1
    vnrointernoA = traerDatos2("select max(nrointerno) as c from asientos", "c", pathDBMySQL) + 2
    
    
    
    
    Me.gVirtual.TextMatrix(0, 0) = "Nombre Cta virtual"
    Me.gVirtual.TextMatrix(0, 1) = "Seleccionada"
    Me.gVirtual.TextMatrix(0, 2) = "Debe"
    Me.gVirtual.TextMatrix(0, 3) = "Haber"
    Me.gVirtual.TextMatrix(0, 4) = "Comentario"
    
    
    Me.gVirtual.ColWidth(0) = 2000
    Me.gVirtual.ColWidth(1) = 1000
    Me.gVirtual.ColWidth(2) = 2000
    Me.gVirtual.ColWidth(3) = 2000
    Me.gVirtual.ColWidth(4) = 4000

    
    Me.gVirtualDetalle.TextMatrix(0, 0) = "Nombre Cta virtual"
    Me.gVirtualDetalle.TextMatrix(0, 1) = "Cta Real"
    Me.gVirtualDetalle.TextMatrix(0, 2) = "Descripcion"
    
    Me.gVirtualDetalle.ColWidth(2) = 6000
    Me.gVirtualDetalle.ColWidth(1) = 2000
    
    If LeerXml("Puesto") = "CAJA" Then
    
        vtitulo2.Text = " ---- Resumen de Ejecución de Ingesos/Egresos ---- "
    End If
   
    
End Sub
Private Function ControlTemporal() As Boolean
On Error Resume Next

    Dim rstemp As New ADODB.Recordset
    Dim sqlTemp As String
    
    sqlTemp = "SELECT * FROM Temp2"
    
    With rstemp
        .Open sqlTemp, ConnDDBB, adOpenStatic, adLockReadOnly
        
        ControlTemporal = Not .EOF
    
    End With
    
    sqlTemp = ""
    
    rstemp.Close
    Set rstemp = Nothing

If Err Then GrabarLog "ControlTemporal", Err.Number & " " & Err.Description, Me.Caption
End Function

Private Sub GrillaBalances_BeforeEdit(Cancel As Boolean)

End Sub

Private Sub gBalance_Click()
vcelda.Text = Me.gBalance.TextMatrix(Me.gBalance.RowSel, Me.gBalance.ColSel)
vcelda.SetFocus
End Sub

Private Sub gVirtual_DblClick()

If (Me.gVirtual.TextMatrix(Me.gVirtual.RowSel, 1) = "X") Then
    Me.gVirtual.TextMatrix(Me.gVirtual.RowSel, 1) = ""
Else
    Me.gVirtual.TextMatrix(Me.gVirtual.RowSel, 1) = "X"
End If


End Sub

Private Sub pb1_Click()
Dim vsql As String

Me.gBalance.Clear
gBalance.Rows = 1

vtimestampAjuste = strfechaMySQL(vvtimestamp)
log.Clear
vfdesde = vfbdesde.Value
vfhasta = vfbhasta.Value


log2.Clear
log3.Clear

CalculosCierreAperturaBalance

Call PushButton13_Click

vsql = "Ahora puede cambiar los importes de las cuentas. " + Chr(13) + "Debe hacer clic sobre la selda y cambiar el valor en la caja de texto."

MsgBox vsql, vbInformation

End Sub
Private Sub CalculosCierreAperturaBalance()
Dim fd, fh As Date

If Not validarCalculosCA("cierre") Then Exit Sub



log2.AddItem ("Fecha cierre: " + Str$(CDate(vfhasta)))
log2.AddItem ("Fecha apertura: " + Str$(CDate(vfhasta) + 1))


vnrobalance = TraerDato("balances", " Activo='S' order by NroBalance Desc", "NroBalance", pathDBMySQL)
vnroasiento = Val(GenerarDato("SELECT MAX(Numero) AS UAsiento FROM Asientos", "UAsiento")) + 1      ' los numeros absolutos
'vnrointerno = UltimoNroInterno2

'Call FechasDelBalance(fd, fh) ' calcula la fecha del balance

Call CalSaldosPerdidasGanancias(ByVal vfdesde, vfhasta, "Perdidas")
Call CalSaldosPerdidasGanancias(vfdesde, vfhasta, "Ganancias")


End Sub
Private Function validarCalculosCA(vca As String) As Boolean
Dim vsql As String
Dim vnroA, vnroC As Long
validarCalculosCA = True

If Not Val(Me.vCnroBalanceCierre.Text) > 0 Then
    MsgBox "Debe seleccionar nro de balance para el cierre", vbInformation
    validarCalculosCA = False
    Exit Function
End If


If Not Val(Me.vCnroBalanceApertura.Text) > 0 Then
    MsgBox "Debe seleccionar nro de balance para el Aperura", vbInformation
    validarCalculosCA = False
    Exit Function
End If


If Not Val(Me.vCnroBalanceApertura.Text) = Val(Me.vCnroBalanceApertura.Text) Then
    MsgBox "Seleccione distintos balances", vbInformation
    validarCalculosCA = False
    Exit Function
End If


' verifico si el asiento de cierre y apertura están hechos

vsql = "select nrocierre as c from balances where nrobalance=" + Str(Me.vCnroBalanceCierre)
vnroC = Val(EsNulo(traerDatos2(vsql, "c", pathDBMySQL)))


If vnroC > 0 And vca = "cierre" Then
    MsgBox "Ya existe el asiento de cierre con nro: " + Str$(vnroC)
    validarCalculosCA = False
    Exit Function
End If


vsql = "select nroApertura as c from balances where nrobalance=" + Str(Me.vCnroBalanceApertura)
vnroA = Val(EsNulo(traerDatos2(vsql, "c", pathDBMySQL)))

If vnroA > 0 And vca = "apertura" Then
    MsgBox "Ya existe el asiento de apertura con nro: " + Str$(vnroA)
    validarCalculosCA = False
    Exit Function
End If


If Not Val(Me.vCnroBalanceApertura) > Val(Me.vCnroBalanceCierre) Then
    MsgBox "Los números de balances no son consecutivos"
    validarCalculosCA = False
    Exit Function
End If


End Function

Private Sub DoAsientosCierre(ByVal vfh As Date, vnroasiento As Long, vnrobalance As Integer)
Dim i As Integer
Dim vtDebe, vtHaber, vimporte, vdebe, vhaber  As Double
Dim vcuenta, vsql  As String
Dim vcant As Integer




vtDebe = 0
vtHaber = 0
vimporte = 0

barra2.Max = Me.gBalance.Rows - 1
barra2.Value = 0

For i = 1 To Me.gBalance.Rows - 1

    vdebe = 0
    vhaber = 0
    
    ' para el cierre se debe dar vuelta
    
    vdebe = Abs(Val(Me.gBalance.TextMatrix(i, 2)))
    vhaber = Abs(Val(Me.gBalance.TextMatrix(i, 1)))
    vcuenta = Me.gBalance.TextMatrix(i, 0)
    
    
    vtDebe = vtDebe + vdebe
    vtHaber = vtHaber + vhaber
    
    If vdebe + vhaber > 0 Then
        vsql = "insert into asientosdetalle (numero,nrobalance,codigocuenta,debe,haber) values (" + Str(vnroasiento) + "," + Str(vCnroBalanceCierre) + ",'" + vcuenta + "'," + Str(vdebe) + "," + Str(vhaber) + ")"
        Call EjecutarScript(vsql, pathDBMySQL)
        vcant = vcant + 1
    End If
    Me.log2.AddItem vsql
    barra2.Value = barra2.Value + 1
Next
    
    Me.log3.AddItem "----------------------------"
    Me.log3.AddItem "Cant. Cierre: " + Str$(vcant)

If Val(Me.vtSaldo) < 0 Then

    vtDebe = vtSaldo
    vtHaber = 0

Else

    vtDebe = 0
    vtHaber = vtSaldo

End If



' graba en asiento detalle DEBE de la cuenta  CierreApertura
vsql = "insert into asientosdetalle (numero,nrobalance,codigocuenta,debe,haber) values (" + Str(vnroasiento) + "," + Str(vnrobalance) + ",'" + Me.vCodCca.Text + "'," + Str(vtDebe) + "," + Str(vtHaber) + ")"
Call EjecutarScript(vsql, pathDBMySQL)

' graba en asiento detalle HABER de la cuenta  CierreApertura
'vsql = "insert into asientosdetalle (numero,nrobalance,codigocuenta,debe,haber) values (" + Str(vnroasiento) + ",'" + Str(vnrobalance) + "','" + Me.vCodCca.Text + "'," + Str(0) + "," + Str(vtHaber) + ")"
'Call EjecutarScript(vsql, pathDBMySQL)

vsql = "insert into asientos (numero,nrointerno,nrobalance,fecha,leyenda)  values (" + Str(vnroasiento) + "," + vnrointernoC + "," + Str(vnrobalance) + ",'" + strfechaMySQL(vfh) + "','" + Me.vleyendaCierre + "')"
Call EjecutarScript(vsql, pathDBMySQL)


vsql = "update balances set nroCierre=" + Str(vnroasiento) + " where nrobalance=" + Str(vnrobalance)
Call EjecutarScript(vsql, pathDBMySQL)

End Sub


Private Sub DoAsientosApertura(ByVal vfh As Date, vnroasiento As Long, vnrobalance As Integer)
Dim i As Integer
Dim vtDebe, vtHaber, vimporte, vdebe, vhaber  As Double
Dim vcuenta, vsql  As String
Dim vcant As Integer

'vfh = vfh + 1

vtDebe = 0
vtHaber = 0
vimporte = 0

barra2.Max = Me.gBalance.Rows - 1
barra2.Value = 0


For i = 1 To Me.gBalance.Rows - 1

    vdebe = 0
    vhaber = 0
    
    vdebe = Val(Me.gBalance.TextMatrix(i, 1))
    vhaber = Val(Me.gBalance.TextMatrix(i, 2))
    vcuenta = Me.gBalance.TextMatrix(i, 0)
    
    vtDebe = vtDebe + vdebe
    vtHaber = vtHaber + vhaber
    
    If vdebe + vhaber > 0 Then
    vsql = "insert into asientosdetalle (numero,nrobalance,codigocuenta,debe,haber) values (" + Str(vnroasiento) + "," + Str(vnrobalance) + ",'" + vcuenta + "'," + Str(vdebe) + "," + Str(vhaber) + ")"
    Call EjecutarScript(vsql, pathDBMySQL)
    vcant = vcant + 1
    End If

    barra2.Value = barra2.Value + 1
    Me.log2.AddItem vsql

Next

Me.log3.AddItem "----------------------------"
Me.log3.AddItem "Cant. Apertura : " + Str$(vcant)


If Val(Me.vtSaldo) > 0 Then

    vtDebe = Val(vtSaldo)
    vtHaber = 0

Else

    vtDebe = 0
    vtHaber = Val(vtSaldo)

End If


   
vsql = "insert into asientosdetalle (numero,nrobalance,codigocuenta,debe,haber) values (" + Str(vnroasiento) + "," + Str(vnrobalance) + ",'" + Me.vCodCca.Text + "'," + Str(vtHaber) + "," + Str(vtDebe) + ")"
Call EjecutarScript(vsql, pathDBMySQL)


'vsql = "insert into asientosdetalle (numero,nrobalance,codigocuenta,debe,haber) values (" + Str(vnroasiento) + "," + Str(vnrobalance) + ",'" + Me.vCodCca.Text + "'," + Str(vtHaber) + "," + Str(0) + ")"
'Call EjecutarScript(vsql, pathDBMySQL)

vsql = "insert into asientos (numero,nrointerno,nrobalance,fecha,leyenda)  values (" + Str(vnroasiento) + "," + vnrointernoA + "," + Str(vnrobalance) + ",'" + strfechaMySQL(vfh) + "','" + Me.vleyendaApertura + "')"
Call EjecutarScript(vsql, pathDBMySQL)


vsql = "update balances set nroApertura=" + Str(vnroasiento) + " where nrobalance=" + Str(vnrobalance)
Call EjecutarScript(vsql, pathDBMySQL)
End Sub



Private Sub CalSaldosPerdidasGanancias(ByVal fd As Date, ByVal fh As Date, vpg As String)
Dim bcuentas As New ADODB.Recordset, sqlCuentas As String
Dim vsaldoG, vsaldoP, vtsaldoG, vtsaldoP As Double ' saldos de ganancias y perdidas
Dim vrow, vcantcero, vcant As Integer

vcantcero = 0
vcant = 0


fh = fh + 1

If vpg = "Perdidas" Then
    sqlCuentas = "SELECT * FROM cuentas WHERE CodigoCuenta >= '5' and CodigoCuenta <= '555555' order by CodigoCuenta asc"
End If

If vpg = "Ganancias" Then
    sqlCuentas = "SELECT * FROM cuentas WHERE CodigoCuenta >= '4' and CodigoCuenta <= '444444' order by CodigoCuenta asc"
End If


If vpg = "123" Then
    sqlCuentas = "SELECT * FROM cuentas WHERE CodigoCuenta >= '1' and CodigoCuenta <= '333333' order by CodigoCuenta asc"
End If


With bcuentas
        Call .Open(sqlCuentas, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        
        If Not .EOF = True Then
            barra2.Value = 0
            barra2.Max = .RecordCount
            .MoveFirst
        End If
    
        ' ----------------
            
            vtsaldoP = 0
            vtsaldoG = 0
           
        
        Do Until .EOF = True
            DoEvents
            
            vsaldoP = 0
            vsaldoG = 0
            
            If .Fields("imputable") = "S" Then
                     
                     If vpg = "Perdidas" Then
                     '  vsaldoP = CalSaldoAnteriorCtaContable("TODOS", .Fields("CodigoCuenta").Value, Val(Me.vCnroBalanceCierre), fh, Me.vDnroBalanceCierre, fd)
                        vsaldoP = CalSaldoAnteriorCtaContable("TODOS", .Fields("CodigoCuenta").Value, Val(Me.vCnroBalanceCierre), fh, Me.vDnroBalanceCierre, fd)
                       
                       vsaldoP = vsaldoP
                       
                       ' si es negativo lo tengo que pasar a ganancias
                                             
                       If vsaldoP > 0 Then
                        vsaldoG = Abs(vsaldoP)
                        vsaldoP = 0
                        Else
                        vsaldoP = Abs(vsaldoP)
                       End If
                       
                       vtsaldoP = vtsaldoP + vsaldoP
                     End If
                     
                     If vpg = "Ganancias" Then
                        vsaldoG = CalSaldoAnteriorCtaContable("TODOS", .Fields("CodigoCuenta").Value, Val(Me.vCnroBalanceCierre), fh, Me.vDnroBalanceCierre, fd)
                        vsaldoG = vsaldoG
                        
                       ' si es negativo lo tengo que pasar a perdida
                       If vsaldoG < 0 Then
                        vsaldoP = Abs(vsaldoG)
                        vsaldoG = 0
                       
                       Else
                        vsaldoG = Abs(vsaldoG)
                       End If
                       
                        vtsaldoG = vtsaldoG + vsaldoG
                    End If
                     
                     
                    If vpg = "123" Then
                        vsaldoG = CalSaldoAnteriorCtaContable("TODOS", .Fields("CodigoCuenta").Value, Val(Me.vCnroBalanceCierre), fh, Me.vDnroBalanceCierre, fd)
                        vsaldoG = vsaldoG
                        
                       ' si es negativo lo tengo que pasar a perdida
                       If vsaldoG < 0 Then
                        vsaldoP = Abs(vsaldoG)
                        vsaldoG = 0
                       
                       Else
                        vsaldoG = Abs(vsaldoG)
                       End If
                       
                        vtsaldoG = vtsaldoG + vsaldoG
                    End If
                      
                     
                     
                     
                     'vsaldo = CalSaldoAnteriorCtaContable("TODOS", .Fields("CodigoCuenta").Value, vnrobalance, fh, "2011-2012", fd)
                    
                    If vsaldoG + vsaldoP = 0 Then
                        vcantcero = vcantcero + 1
                    Else
                        vcant = vcant + 1
                    End If
                    
                    
                    ' gBalance.Rows = gBalance.Rows + 1
                     vrow = gBalance.Rows
                     Call Me.gBalance.AddItem(.Fields("CodigoCuenta").Value + vbTab + Format(vsaldoG, "########0.00") + vbTab + Format(vsaldoP, "########0.00") + vbTab + .Fields("Cuenta").Value, vrow)
                     
            End If
            
            'Call CalcularMovimientos(vmarca, Index, .Fields("CodigoCuenta").Value, .Fields("Cuenta").Value, vfdesde, vfhasta, vfbdesde, vfbhasta, vcb, vnb)
            
            barra2.Value = barra2.Value + 1
            .MoveNext
           ' Call LastKlexRow(gBalance)
        Loop
End With

        log3.AddItem "--------------------------------"
        log3.AddItem vpg
        log3.AddItem "Cant. cero : " + Str(vcantcero)
        log3.AddItem "Cant. valor : " + Str(vcant)
        log3.AddItem "Total Ganancia : " + Str(vtsaldoG)
        log3.AddItem "Total Perdida : " + Str(vtsaldoP)
        
        
 
End Sub


Public Sub PbAcciones_Click(Index As Integer)
On Error Resume Next
Dim vtipo As String

vtimestampAjuste = strfechaMySQL(vvtimestamp)

    Select Case Index
    
        Case 0
            
            Dim vmarca As String
            
            If rdinterno Then vmarca = "INTERNO"
            If rdNormal Then vmarca = "NORMAL"
            If rdtodos Then vmarca = "TODOS"
            
            If viddesde = 0 Then
                        vnrobalance = selectNrobalance(dtpFecha(0).Value, dtpFecha(1).Value, vnrobalance)
                        If vnrobalance = 0 And Me.chkvarios.Value = xtpUnchecked Then Exit Sub
                        
                        vcodigoBalance = traerDatos2("select * from balances where nrobalance=" + Str(vnrobalance), "codigo", pathDBMySQL)
                        Me.vcbalance = vcodigoBalance
            End If
            
            If Me.rbgral.Value Then vtipo = "general"
            
            Call GenerarBalance(vmarca, 1, barra, dtpFecha(0).Value, dtpFecha(1).Value, vfdesde, ByVal vfhasta, vcodigoBalance, vnrobalance, txtCuentaContable(0).Text, txtCuentaContable(2).Text, vtipo, viddesde, vidhasta)
            
            If Me.rbgral Then BalanceGral  ' calcula los datos para el balance gral
            
            Call MostrarReporteBalance(1)
            
            finit   'inicia todo de nuevo
            
            
        Case 1
            Unload Me
    End Select
    
If Err Then GrabarLog "PbAcciones_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub BalanceGral()
On Error Resume Next
    Dim vsinicial, vSPeriodo, vSFinal, vtDebe, vtHaber, vtPres, vtSaldo2  As Double
    
    Dim btemp3 As New ADODB.Recordset, sqlGuardar, vsql As String
    
' --- init ---
vsinicial = 0
vSPeriodo = 0
vSFinal = 0
vtDebe = 0
vtHaber = 0
vtPres = 0
vtSaldo2 = 0

' --- init ---


    sqlGuardar = "SELECT * FROM Temp2 order by c02 asc"
 
        With btemp3
        Call .Open(sqlGuardar, ConnDDBB, adOpenDynamic, adLockOptimistic)
        
        .MoveFirst
        
        Me.barra.Max = .RecordCount + 1
        Me.barra.Value = 0
        Do Until .EOF
            
            'vsinicial = bgtotalizar(.Fields("c02"), "SInicial")
            'vSPeriodo = bgtotalizar(.Fields("c02"), "SPeriodo")
            'vSFinal = bgtotalizar(.Fields("c02"), "SFinal")
            'vtDebe = bgtotalizar(.Fields("c02"), "tdebe")
            'vtHaber = bgtotalizar(.Fields("c02"), "thaber")
            
           vsql = "select t.Imputable as c from cuentas t where t.codigocuenta = '" + Trim(.Fields("c02")) + "'"
            
            If traerDatos2(vsql, "c", pathDBMySQL) = "N" Then
                Call bgtotalizar(.Fields("c02"), vsinicial, vSPeriodo, vSFinal, vtDebe, vtHaber, vtPres)
             '   Call bgtotalizar2(.Fields("c06"), vsinicial, vSPeriodo, vSFinal, vtDebe, vtHaber, vtPres)
                
                vtSaldo2 = vtPres - vSFinal

                vsql = "update  temp2 set c07='" + Str(vsinicial) + "', c11='" + Str(vSPeriodo) + "',c13='" + Str(vSFinal) + "', c03='" + Str(vtDebe) + "',c04='" + Str(vtHaber) + "',c14='" + Str(vtPres) + "',c15='" + Str(vtSaldo2) + "' where c02='" + Trim(.Fields("c02")) + "'"
                Call EjecutarScript(vsql, pathDBMySQL)
            
            Else
               Debug.Print ">>>>> " + .Fields("c02")
            End If
            
          '  If vsinicial > -1 Then
                
'             vsql = "update  temp2 set c07='" + Str(vsinicial) + "', c11='" + Str(vSPeriodo) + "',c13='" + Str(vSFinal) + "', c03='" + Str(vtDebe) + "',c04='" + Str(vtHaber) + "' where c02='" + Trim(.Fields("c02")) + "'"
'             Call EjecutarScript(vsql, pathDBMySQL)
            '.Fields("c07") = vsinicial
            '.Fields("c11") = vSPeriodo
            '.Fields("c13") = vSFinal
                
             Debug.Print .Fields("c02")
           ' End If
        .MoveNext
        '.MoveNext
        Me.barra.Value = barra.Value + 1
        Loop
        .Update
       End With
       
       log.AddItem "Sale de Balance General"

End Sub

'Function bgtotalizar(vcodigo As String, vcampo As String) As Double
Function bgtotalizar(vcodigo As String, ByRef vsinicial, ByRef vSPeriodo, ByRef vSFinal, ByRef vtDebe, ByRef vtHaber, ByRef vtPres) As Double

On Error Resume Next

Dim vsql, vimputable As String
Dim vca, vc As String

Dim td, th, tsi, tsp, tsf, tpres As Double

Dim rs  As New ADODB.Recordset

tsi = 0
tsp = 0
tsf = 0
td = 0
th = 0

vsinicial = 0
vSPeriodo = 0
vSFinal = 0
vtDebe = 0
vtHaber = 0
vtPres = 0




vsql = "select * from cuentas where codigoCuenta='" + Trim(vcodigo) + "'"
vimputable = traerDatos2(vsql, "imputable", pathDBMySQL)


'If Not vimputable = "N" Then   '
'    bgtotalizar = -1  ' si es negativo no aplica porque es cta imputable
'    Exit Function
'End If


'If Not vimputable = "N" Then   '
'    bgtotalizar = -1  ' si es negativo no aplica porque es cta imputable
'    Exit Function
'End If




'vca = " replace(replace(c02,'0',''),'.','') "
'vc = Replace(Replace(vcodigo, "0", ""), ".", "")


vca = " replace(replace(c02,'0',''),'.','') "
vc = Replace(Replace(vcodigo, "0", ""), ".", "")


'vsql = "select sum(cast(c03 as decimal)) as tdebe, sum(cast(c04 as decimal)) as thaber, sum(cast(c07 as decimal)) as Sinicial,   sum(cast(c11 as decimal)) as SPeriodo,   sum(cast(c13 as decimal)) as SFinal from temp2 t where c10 = 'S' and c02 like '" + Trim(vcodigo) + "%' Group By c10"
'vsql = "select sum(cast(c03 as decimal)) as tdebe, sum(cast(c04 as decimal)) as thaber, sum(cast(c07 as decimal)) as Sinicial,   sum(cast(c11 as decimal)) as SPeriodo,   sum(cast(c13 as decimal)) as SFinal from temp2 t where " + vca + " like '" + vc + "%' and C09 = 'S' Group By c02 "
vsql = "select sum(cast(c14 as decimal)) as tpres, sum(cast(c03 as decimal)) as tdebe, sum(cast(c04 as decimal)) as thaber, sum(cast(c07 as decimal)) as Sinicial,   sum(cast(c11 as decimal)) as SPeriodo,   sum(cast(c13 as decimal)) as SFinal from temp2 t where " + vca + " like '" + vc + "%' and C09 = 'S'  "


Call rs.Open(vsql, ConnDDBB, adOpenStatic, adLockReadOnly)

rs.MoveFirst
Do Until rs.EOF

    td = td + rs.Fields("tdebe")
    
    th = th + rs.Fields("thaber")
    
    tsi = tsi + rs.Fields("sinicial")
    
    tsp = tsp + rs.Fields("speriodo")
    
    tsf = tsf + rs.Fields("sfinal")

    tpres = tpres + rs.Fields("tpres")
    
    rs.MoveNext
Loop

vsinicial = tsi
vSPeriodo = tsp
vSFinal = tsf
vtDebe = td
vtHaber = th
vtPres = tpres

'bgtotalizar = traerDatos2(vsql, vcampo, pathDBMySQL)

If Err Then
    bgtotalizar = 0
    Exit Function
End If

End Function


Function bgtotalizar2(vcodigo As String, ByRef vsinicial, ByRef vSPeriodo, ByRef vSFinal, ByRef vtDebe, ByRef vtHaber, ByRef vtPres) As Double

On Error Resume Next

Dim vsql, vimputable As String
Dim vca, vc As String

Dim td, th, tsi, tsp, tsf, tpres As Double

Dim rs  As New ADODB.Recordset

tsi = 0
tsp = 0
tsf = 0
td = 0
th = 0

vsinicial = 0
vSPeriodo = 0
vSFinal = 0
vtDebe = 0
vtHaber = 0
vtPres = 0




vsql = "select * from cuentas where codigoCuenta='" + Trim(vcodigo) + "'"
vimputable = traerDatos2(vsql, "imputable", pathDBMySQL)


vca = "c06"
vc = vcodigo

vsql = "select sum(cast(c14 as decimal)) as tpres, sum(cast(c03 as decimal)) as tdebe, " + _
" sum(cast(c04 as decimal)) as thaber, sum(cast(c07 as decimal)) as Sinicial, " + _
" sum(cast(c11 as decimal)) as SPeriodo,   sum(cast(c13 as decimal)) as SFinal " + _
" from temp2 t " + _
" where " + vca + " like '" + vc + "%' and C09 = 'S'  "


vsql = "select sum(c14) as tpres, sum(c03) as tdebe, " + _
" sum(c04) as thaber, sum(c07) as Sinicial, " + _
" sum(c11) as SPeriodo,   sum(c13) as SFinal " + _
" from temp2 t " + _
" where " + vca + " like '" + vc + "%' and C09 = 'S'  "


Call rs.Open(vsql, ConnDDBB, adOpenStatic, adLockReadOnly)

rs.MoveFirst
Do Until rs.EOF

    td = td + rs.Fields("tdebe")
    
    th = th + rs.Fields("thaber")
    
    tsi = tsi + rs.Fields("sinicial")
    
    tsp = tsp + rs.Fields("speriodo")
    
    tsf = tsf + rs.Fields("sfinal")

    tpres = tpres + rs.Fields("tpres")
    
    rs.MoveNext
Loop

vsinicial = tsi
vSPeriodo = tsp
vSFinal = tsf
vtDebe = td
vtHaber = th
vtPres = tpres

'bgtotalizar = traerDatos2(vsql, vcampo, pathDBMySQL)

If Err Then
    bgtotalizar2 = 0
    Exit Function
End If

End Function










Private Sub pbCarga_Click(Index As Integer)
On Error Resume Next

    vVuelveBusqueda = Me.Name
    vVieneBusqueda = pbCarga(Index).Tag

    Select Case Index
        
        Case 0 To 6
            frmBusqueda.Show
    
    End Select

If Err Then GrabarLog "pbCarga_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub PushButton1_Click()
Call fbuscarGrilla("balances", "codigo", "idBalances", Me.vcbalance.Name, Me)
End Sub

Private Sub TxF1_Click()

End Sub

Private Sub PushButton10_Click()
Dim vlinea As String

If Not validarVirtual Then Exit Sub

vlinea = Me.gVirtual.TextMatrix(Me.gVirtual.RowSel, 0) + vbTab + Me.vCctaReal + vbTab + Me.vDCtaReal
Me.gVirtualDetalle.AddItem (vlinea)

vCctaReal.Text = ""
vDCtaReal = ""
vDCtaReal.Tag = ""

End Sub

Private Sub PushButton11_Click()
Call fbuscarGrilla("Cuentas", "Cuenta", "CodigoCuenta", vDCtaReal.Name, Me)  ' ema:
End Sub
Function validarVirtual() As Boolean
validarVirtual = True
  
If Me.gVirtual.TextMatrix(Me.gVirtual.RowSel, 0) = "" Then
    MsgBox "Debe seleccionar una cuanta virtual", vbInformation
    validarVirtual = False
End If
End Function

Private Sub PushButton12_Click()
 Me.gVirtualDetalle.ColSel = 0
Me.gVirtualDetalle.Col = 0
Me.gVirtualDetalle.Sort = flexSortGenericAscending
End Sub

Private Sub PushButton13_Click()
Dim i As Integer
Dim vtD, vth As Double

vtD = 0
vth = 0

For i = 0 To gBalance.Rows - 1
    vtD = vtD + Val(gBalance.TextMatrix(i, 1))
    vth = vth + Val(gBalance.TextMatrix(i, 2))
    log3.AddItem Str(vtD) + " " + Str(vth)
Next

Me.vtDebe = Format(vtD, "###,###,##0.00")
Me.vtHaber = Format(vth, "###,###,##0.00")

Me.vtSaldo = vtD - vth

End Sub

Private Sub PushButton14_Click()
    Unload Mantenimiento
    Load Mantenimiento
    Dim vcriterio As String
    
    vcriterio = "Desde la fecha: " + Str$(Me.vfvdesde) + " hasta: " + Str$(Me.vfvhasta)
    
    With drBalanceVirtual
    
    .Sections("titulo").Controls("ecriterio").Caption = vcriterio
    .Sections("titulo").Controls("comentario").Caption = vcomentario
    
    
    End With
    drBalanceVirtual.Show
End Sub

Private Sub PushButton15_Click()
MousePointer = vbHourglass
vdisplay.Caption = "Espere un momento"
Call fMostrarGrilla(fSQLConciliaBancoCtas2(Me.vfcdesde, Me.vfchasta, ""))
vdisplay.Caption = ""
MousePointer = vbDefault


End Sub

Private Sub PushButton19_Click()
Dim fd, fh As Date

gBalance.Rows = 1

If Not validarCalculosCA("apertura") Then Exit Sub


log2.AddItem ("Fecha cierre: " + Str$(CDate(vfhasta)))
log2.AddItem ("Fecha apertura: " + Str$(CDate(vfhasta) + 1))


vnrobalance = TraerDato("balances", " Activo='S' order by NroBalance Desc", "NroBalance", pathDBMySQL)
vnroasiento = Val(GenerarDato("SELECT MAX(Numero) AS UAsiento FROM Asientos", "UAsiento")) + 1      ' los numeros absolutos
'vnrointerno = UltimoNroInterno2

'Call FechasDelBalance(fd, fh) ' calcula la fecha del balance

Call CalSaldosPerdidasGanancias(ByVal vfdesde, vfhasta, "123")
'Call CalSaldosPerdidasGanancias(vfdesde, vfhasta, "Ganancias")

End Sub

Private Sub PushButton2_Click()
    Call PushButton13_Click
    Call aplicarAsientos("cierre")
End Sub


Private Sub aplicarAsientos(vca As String)
'Call CalSaldosGanancias(fd, fh)
On Error Resume Next

Dim fd, fh As Date

If Not validarCalculosCA(vca) Then Exit Sub



'vfdesde = traerDatos2("select * from balances where NroBalance=" + Str$(Me.vCnroBalanceCierre.Text) + " order by idBalances desc", "FechaInicio", pathDBMySQL)
'vfhasta = traerDatos2("select * from balances where NroBalance=" + Str$(Me.vCnroBalanceCierre.Text) + " order by idBalances desc", "FechaFin", pathDBMySQL)


log2.AddItem ("Fecha cierre: " + Str$(CDate(vfhasta)))
log2.AddItem ("Fecha apertura: " + Str$(CDate(vfhasta) + 1))


vnrobalance = TraerDato("balances", " Activo='S' order by NroBalance Desc", "NroBalance", pathDBMySQL)
vnroasiento = Val(GenerarDato("SELECT MAX(Numero) AS UAsiento FROM Asientos", "UAsiento")) + 1      ' los numeros absolutos


'Call FechasDelBalance(fd, fh) ' calcula la fecha del balance

If vca = "cierre" Then
    Call DoAsientosCierre(ByVal Me.vfcierre, vnroasiento, Me.vCnroBalanceCierre)
End If

If vca = "apertura" Then
    Call DoAsientosApertura(ByVal Me.vfapertura, vnroasiento + 1, Me.vCnroBalanceApertura)
End If

MsgBox "Los asientos de cierre y apertura fueron realizados." + Chr(13) + "Verifique con los nros de asientos : " + Str$(vnroasiento) + " y " + Str$(vnroasiento + 1)

If Err < 0 Then
    MsgBox "Acurrió un error inesperado al realizar los asientos de cierre / apertura." + Chr(13) + "Verifique con los nros de asientos : " + Str$(vnroasiento) + " y " + Str$(vnroasiento + 1)
    Exit Sub
End If


End Sub

Private Sub PushButton20_Click()
    Call PushButton13_Click
    Call aplicarAsientos("apertura")
End Sub

Private Sub PushButton21_Click()
drBalance.Show
End Sub

Private Sub PushButton3_Click()
Call fbuscarGrilla("balances", "codigo", "NroBalance", Me.vDnroBalanceCierre.Name, Me)   ' ema:
End Sub

Private Sub PushButton4_Click()
Call fbuscarGrilla("balances", "codigo", "NroBalance", Me.vDnroBalanceApertura.Name, Me)
End Sub

Private Sub PushButton5_Click()
Me.gBalance.TextMatrix(Me.gBalance.RowSel, Me.gBalance.ColSel) = vcelda.Text
End Sub

Private Sub PushButton6_Click()
Me.gVirtual.AddItem (Me.vVirtual.Text)
End Sub

Private Sub PushButton7_Click()
Call PushButton12_Click

If rdinterno Then vmarca = "INTERNO"
If rdNormal Then vmarca = "NORMAL"
If rdtodos Then vmarca = "TODOS"

If Not validarVirtual2 = True Then Exit Sub

initVirtual
CalcularBalanceVirtual

pasarGrillaATabla

End Sub
Function validarVirtual2() As Boolean
If gVirtual.Rows = 2 Then
validarVirtual2 = False
Else
validarVirtual2 = True

End If
End Function

Private Sub pasarGrillaATabla()
Dim vsql, vvalores As String
Dim vsaldo As Double

Dim i As Integer

Call EjecutarScript("delete from ctavirtual", pathDBMySQL)


For i = 1 To Me.gVirtual.Rows - 1
    vsaldo = Val(Me.gVirtual.TextMatrix(i, 2)) - Val(Me.gVirtual.TextMatrix(i, 3))
    
    vvalores = "'" + Me.gVirtual.TextMatrix(i, 0) + "'," + Str$(vsaldo)

    vsql = "insert into ctavirtual (nombre,saldo) values (" + vvalores + ")"

    Call EjecutarScript(vsql, pathDBMySQL)

Next



End Sub

Private Sub initVirtual()
    vnrobalance = selectNrobalance(Me.vfvdesde.Value, vfvhasta.Value, vnrobalance)
End Sub
Private Sub CalcularBalanceVirtual()
Dim i As Integer
Dim vValor, vVdebe, vVhasta  As Double
Dim vtValor, vtVdebe, vtVhasta  As Double

Me.vttdebe.Caption = "0"
Me.vtthaber.Caption = "0"


b3.Max = Me.gVirtual.Rows - 1
b3.Value = 0
        vVdebe = 0
        vVhasta = 0
For i = 1 To Me.gVirtual.Rows - 1

    If Not Me.gVirtual.TextMatrix(i, 1) = "X" And Not Me.gVirtual.TextMatrix(i, 0) = "Totales:" Then
        vValor = CalBVDetalle(Me.gVirtual.TextMatrix(i, 0), vmarca)
        If Not vValor < 0 Then
            vVdebe = vValor
        Else
            vVhasta = -1 * vValor
        End If
        
        Me.gVirtual.TextMatrix(i, 2) = Format(vVdebe, "########0.00")
        Me.gVirtual.TextMatrix(i, 3) = Format(vVhasta, "########0.00")
        
        vtVdebe = vtVdebe + vVdebe
        vtVhasta = vtVhasta + vVhasta
        
        
        vVdebe = 0
        vVhasta = 0
        
        b3.Value = b3.Value + 1
        
    End If

Next

        Me.vttdebe.Caption = Format(vtVdebe, "###,###,##0.00")
        Me.vtthaber.Caption = Format(vtVhasta, "###,###,##0.00")
        Me.vresta.Caption = Format(vtVdebe - vtVhasta, "###,###,##0.00")
        
        'Me.gVirtual.AddItem ("Totales:" + vbTab + vbTab + Format(vtVdebe, "###,###,##0.00") + vbTab + Format(vtVhasta, "###,###,##0.00"))

End Sub
Function CalBVDetalle(vVirtual As String, ByVal vmarca As String) As Double
Dim i As Integer
Dim vtotal, vValor As Double


With Me.gVirtualDetalle

    
b4.Value = 0
b4.Max = .Rows - 1


For i = 1 To .Rows - 1

    If .TextMatrix(i, 0) = vVirtual Then
        vValor = CalSaldoAnteriorCtaContable(vmarca, .TextMatrix(i, 1), vnrobalance, vfvhasta.Value, vcodigoBalance, vfvdesde.Value)
    End If
    
    vtotal = vtotal + vValor
  b4.Value = b4.Value + 1
Next
End With

CalBVDetalle = vtotal

End Function






Private Sub PushButton8_Click()
gVirtual.RemoveItem (Me.gVirtual.RowSel)
End Sub

Private Sub PushButton9_Click()
Me.gVirtualDetalle.RemoveItem (gVirtualDetalle.RowSel)
End Sub

Private Sub txtCuentaContable_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
    
    If KeyAscii = 13 Then
        
        Select Case Index
    
            Case 0
                If Not txtCuentaContable(Index).Text = "" Then
                    txtCuentaContable(Index + 1).Text = TraerDato("Cuentas", "(CodigoCuenta = " & Trim(txtCuentaContable(Index).Text) & ")", "Cuenta")
                End If
                
                txtCuentaContable(Index + 2).SetFocus
            Case 1
                txtCuentaContable(Index + 1).SetFocus
            Case 2
                If Not txtCuentaContable(Index).Text = "" Then
                    txtCuentaContable(Index + 1).Text = TraerDato("Cuentas", "(CodigoCuenta = " & Trim(txtCuentaContable(Index).Text) & ")", "Cuenta")
                End If
                
                PbAcciones(0).SetFocus
            Case 3
                PbAcciones(0).SetFocus
        
        End Select
    
    End If

If Err Then GrabarLog "txtCuentaContable_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtnrobalance_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then PbAcciones(0).SetFocus
End Sub

Private Sub vcbalance_Change()
vcodigoBalance = Me.vcbalance.Text
End Sub

Private Sub vcelda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    PushButton5.SetFocus
End If
End Sub

Private Sub vCnroBalanceApertura_Change()
finitActualizaFecha
End Sub

Private Sub vCnroBalanceCierre_Change()
finitActualizaFecha
End Sub
Private Sub finitActualizaFecha()
On Error Resume Next

vfdesde = traerDatos2("select * from balances where NroBalance=" + Str$(Me.vCnroBalanceCierre.Text) + " order by idBalances desc", "FechaInicio", pathDBMySQL)
vfhasta = traerDatos2("select * from balances where NroBalance=" + Str$(Me.vCnroBalanceCierre.Text) + " order by idBalances desc", "FechaFin", pathDBMySQL)

  
    vfbdesde.Value = vfdesde
    vfbhasta.Value = CDate(vfhasta)
If Err Then Exit Sub
End Sub

Private Sub vDCtaReal_Change()
Me.vCctaReal.Text = Me.vDCtaReal.Tag
End Sub

Private Sub vDescCca_Change()
Me.vCodCca.Text = Me.vDescCca.Tag
End Sub

Private Sub vDnroBalanceApertura_Change()
Me.vCnroBalanceApertura.Text = vDnroBalanceApertura.Tag
End Sub

Private Sub vDnroBalanceCierre_Change()
Me.vCnroBalanceCierre.Text = vDnroBalanceCierre.Tag
End Sub

